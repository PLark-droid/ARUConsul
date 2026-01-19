/**
 * 完了率フィールドのセットアップ v2
 * タスクテーブルに完了率を自動計算する数式フィールドを追加
 *
 * LarkBaseの制限:
 * - 数式フィールドは同一レコード内のフィールドのみ参照可能
 * - 他レコードの集計にはロールアップが必要
 * - ロールアップには双方向リンクが必要
 *
 * 対応策:
 * - WBSカテゴリテーブルに完了数と総数を計算するフィールドを追加
 * - タスクテーブルからルックアップで完了率を取得
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
const categoryTableId = 'tbleaLCFajhx7KcN';
const taskTableId = 'tblHiAuZsUyWxlAY';

async function main() {
  console.log('=== 完了率セットアップ v2 ===\n');

  try {
    // タスクテーブルの既存フィールドを確認
    const taskFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const taskFields = taskFieldsRes.data.items || [];

    // 既存の完了率フィールドを削除
    const oldCompletionField = taskFields.find(f => f.field_name === '完了率');
    if (oldCompletionField) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${oldCompletionField.field_id}`
      });
      console.log('既存の完了率フィールドを削除');
    }

    // WBSカテゴリリンクがあるか確認
    const linkField = taskFields.find(f => f.field_name === 'WBSカテゴリリンク');
    if (linkField) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${linkField.field_id}`
      });
      console.log('既存のリンクフィールドを削除');
    }

    await new Promise(r => setTimeout(r, 500));

    // カテゴリテーブルを取得
    const categoriesRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/records`,
      params: { page_size: 100 }
    });
    const categories = categoriesRes.data.items || [];

    // 全タスクを取得
    const tasksRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
      params: { page_size: 500 }
    });
    const tasks = tasksRes.data.items || [];

    // カテゴリごとの完了率を計算
    const categoryStats = {};
    for (const cat of categories) {
      const wbs = cat.fields['WBS番号'];
      categoryStats[wbs] = { total: 0, completed: 0, rate: '0%' };
    }

    for (const task of tasks) {
      const wbs = task.fields['WBS番号'];
      if (!wbs || !wbs.includes('.')) continue;

      const categoryNum = wbs.split('.')[0];
      if (!categoryStats[categoryNum]) continue;

      categoryStats[categoryNum].total++;
      if (task.fields['ステータス'] === '完了') {
        categoryStats[categoryNum].completed++;
      }
    }

    // 完了率を計算
    for (const [wbs, stats] of Object.entries(categoryStats)) {
      if (stats.total > 0) {
        const rate = Math.round((stats.completed / stats.total) * 100);
        stats.rate = `${rate}%`;
      }
    }

    console.log('\nカテゴリ別完了率:');
    for (const [wbs, stats] of Object.entries(categoryStats)) {
      const cat = categories.find(c => c.fields['WBS番号'] === wbs);
      if (cat) {
        console.log(`  ${cat.fields['カテゴリ名']}: ${stats.completed}/${stats.total} (${stats.rate})`);
      }
    }

    // WBSカテゴリテーブルに完了率フィールドを追加
    console.log('\nWBSカテゴリに完了率フィールドを追加中...');

    const catFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });
    const catFields = catFieldsRes.data.items || [];

    const existingCatRate = catFields.find(f => f.field_name === '完了率');
    if (!existingCatRate) {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`,
        data: {
          field_name: '完了率',
          type: 1  // テキスト
        }
      });
      console.log('✓ 完了率フィールドを追加');
    }

    await new Promise(r => setTimeout(r, 300));

    // カテゴリの完了率を更新
    console.log('\nカテゴリの完了率を更新中...');
    for (const cat of categories) {
      const wbs = cat.fields['WBS番号'];
      const stats = categoryStats[wbs] || { rate: '0%' };

      await client.request({
        method: 'PUT',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/records/${cat.record_id}`,
        data: {
          fields: { '完了率': stats.rate }
        }
      });
      await new Promise(r => setTimeout(r, 100));
    }
    console.log('✓ カテゴリの完了率を更新完了');

    // タスクテーブルに完了率フィールドを追加（カテゴリの完了率をコピー）
    console.log('\nタスクテーブルに完了率フィールドを追加中...');

    const existingTaskRate = taskFields.find(f => f.field_name === '完了率');
    if (!existingTaskRate) {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: '完了率',
          type: 1  // テキスト
        }
      });
      console.log('✓ 完了率フィールドを追加');
    }

    await new Promise(r => setTimeout(r, 300));

    // 各タスクに完了率を設定
    console.log('\n各タスクの完了率を更新中...');
    let updateCount = 0;

    for (const task of tasks) {
      const wbs = task.fields['WBS番号'];
      if (!wbs || !wbs.includes('.')) continue;

      const categoryNum = wbs.split('.')[0];
      const stats = categoryStats[categoryNum] || { rate: '0%' };

      await client.request({
        method: 'PUT',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records/${task.record_id}`,
        data: {
          fields: { '完了率': stats.rate }
        }
      });

      updateCount++;
      await new Promise(r => setTimeout(r, 50));
    }
    console.log(`✓ ${updateCount}件のタスクを更新`);

    // 自動化スクリプトのセットアップ案内
    console.log('\n=== セットアップ完了 ===');
    console.log('\nLarkBase URL: https://www.feishu.cn/base/' + appToken);
    console.log('\n【重要】ステータス変更時の自動更新について:');
    console.log('LarkBaseの自動化機能を使用してください:');
    console.log('');
    console.log('1. LarkBaseを開く');
    console.log('2. 右上の「自動化」→「新規自動化」をクリック');
    console.log('3. トリガー: 「レコードが更新されたとき」を選択');
    console.log('4. 条件: 「ステータス」フィールドが変更されたとき');
    console.log('5. アクション: 「Webhookを送信」を選択');
    console.log('');
    console.log('または、npm run update-completion を定期実行してください。');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
