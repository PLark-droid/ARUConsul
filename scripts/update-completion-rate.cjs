/**
 * WBSカテゴリの完了率を計算して更新するスクリプト
 */

const lark = require('@larksuiteoapi/node-sdk');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const client = new lark.Client({
  appId: process.env.LARK_APP_ID,
  appSecret: process.env.LARK_APP_SECRET,
  appType: lark.AppType.SelfBuild,
  domain: lark.Domain.Feishu,
});

const appToken = process.env.LARK_BASE_APP_TOKEN;
const categoryTableId = 'tbleaLCFajhx7KcN';  // WBSカテゴリ
const taskTableId = 'tblHiAuZsUyWxlAY';      // タスク

async function main() {
  console.log('=== WBSカテゴリ完了率更新 ===\n');

  try {
    // ステップ1: WBSカテゴリテーブルに完了率フィールドを追加
    console.log('ステップ1: 完了率フィールドを追加中...');

    // まず現在のフィールドを確認
    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });

    const existingFields = fieldsRes.data.items || [];
    const hasCompletionRate = existingFields.some(f => f.field_name === '完了率');
    const hasTotalTasks = existingFields.some(f => f.field_name === 'タスク数');
    const hasCompletedTasks = existingFields.some(f => f.field_name === '完了タスク数');

    // タスク数フィールド追加
    if (!hasTotalTasks) {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
        data: {
          field_name: 'タスク数',
          type: 2  // 数値
        }
      });
      console.log('  ✓ タスク数フィールド追加');
    }

    // 完了タスク数フィールド追加
    if (!hasCompletedTasks) {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
        data: {
          field_name: '完了タスク数',
          type: 2  // 数値
        }
      });
      console.log('  ✓ 完了タスク数フィールド追加');
    }

    // 完了率フィールド追加
    if (!hasCompletionRate) {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
        data: {
          field_name: '完了率',
          type: 1  // テキスト（パーセント表示用）
        }
      });
      console.log('  ✓ 完了率フィールド追加');
    }

    await new Promise(r => setTimeout(r, 500));

    // ステップ2: 全タスクを取得
    console.log('\nステップ2: タスクデータ取得中...');

    let allTasks = [];
    let pageToken = null;

    do {
      const params = { page_size: 500 };
      if (pageToken) params.page_token = pageToken;

      const tasksRes = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
        params
      });

      if (tasksRes.data.items) {
        allTasks = allTasks.concat(tasksRes.data.items);
      }
      pageToken = tasksRes.data.page_token;
    } while (pageToken);

    console.log(`  ✓ ${allTasks.length}件のタスクを取得`);

    // ステップ3: カテゴリごとに集計
    console.log('\nステップ3: カテゴリごとに完了率を計算中...');

    const categoryStats = {};

    for (const task of allTasks) {
      const wbs = task.fields['WBS番号'];
      if (!wbs || !wbs.includes('.')) continue;

      // WBS番号の先頭（カテゴリ番号）を取得
      const categoryNum = wbs.split('.')[0];

      if (!categoryStats[categoryNum]) {
        categoryStats[categoryNum] = { total: 0, completed: 0 };
      }

      categoryStats[categoryNum].total++;

      // ステータスが「完了」かチェック
      const status = task.fields['ステータス'];
      if (status === '完了') {
        categoryStats[categoryNum].completed++;
      }
    }

    console.log('\n集計結果:');
    for (const [cat, stats] of Object.entries(categoryStats)) {
      const rate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
      console.log(`  カテゴリ${cat}: ${stats.completed}/${stats.total} (${rate}%)`);
    }

    // ステップ4: WBSカテゴリテーブルを更新
    console.log('\nステップ4: WBSカテゴリテーブルを更新中...');

    // カテゴリレコードを取得
    const categoriesRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/records`,
      params: { page_size: 100 }
    });

    const categories = categoriesRes.data.items || [];

    for (const cat of categories) {
      const wbsNum = cat.fields['WBS番号'];
      const stats = categoryStats[wbsNum] || { total: 0, completed: 0 };
      const rate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;

      await client.request({
        method: 'PUT',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/records/${cat.record_id}`,
        data: {
          fields: {
            'タスク数': stats.total,
            '完了タスク数': stats.completed,
            '完了率': `${rate}%`
          }
        }
      });

      console.log(`  ✓ ${cat.fields['カテゴリ名']}: ${stats.completed}/${stats.total} (${rate}%)`);
      await new Promise(r => setTimeout(r, 200));
    }

    console.log('\n=== 完了率更新完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
