/**
 * タスクテーブルのみで完了率を管理
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
  console.log('=== タスクテーブル完了率セットアップ ===\n');

  try {
    // ステップ1: WBSカテゴリテーブルから完了率フィールドを削除
    console.log('ステップ1: WBSカテゴリから完了率フィールドを削除...');

    const catFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields`
    });
    const catFields = catFieldsRes.data.items || [];

    const catRateField = catFields.find(f => f.field_name === '完了率');
    if (catRateField) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${categoryTableId}/fields/${catRateField.field_id}`
      });
      console.log('  ✓ WBSカテゴリの完了率を削除');
    } else {
      console.log('  - 削除対象なし');
    }

    // ステップ2: タスクテーブルのフィールドを確認・作成
    console.log('\nステップ2: タスクテーブルのフィールドを確認...');

    const taskFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const taskFields = taskFieldsRes.data.items || [];

    const existingField = taskFields.find(f => f.field_name === '完了率');

    if (existingField) {
      // 既存フィールドを削除
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${existingField.field_id}`
      });
      console.log('  ✓ 既存の完了率フィールドを削除');
      await new Promise(r => setTimeout(r, 500));
    }

    // 新規作成（数値型・パーセント形式）
    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: '完了率',
        type: 2,
        property: {
          formatter: '0%'
        }
      }
    });
    console.log('  ✓ 完了率フィールドを作成（数値・%形式）');

    await new Promise(r => setTimeout(r, 500));

    // ステップ3: 全タスクを取得
    console.log('\nステップ3: 全タスクを取得...');

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

    // ステップ4: カテゴリごとの完了率を計算
    console.log('\nステップ4: カテゴリごとの完了率を計算...');

    const categoryStats = {};

    for (const task of allTasks) {
      const category = task.fields['カテゴリ'];
      if (!category) continue;

      if (!categoryStats[category]) {
        categoryStats[category] = { total: 0, completed: 0 };
      }

      categoryStats[category].total++;

      if (task.fields['ステータス'] === '完了') {
        categoryStats[category].completed++;
      }
    }

    // 完了率を計算
    const categoryRates = {};
    console.log('\n  カテゴリ別完了率:');
    for (const [category, stats] of Object.entries(categoryStats)) {
      const rate = stats.total > 0 ? stats.completed / stats.total : 0;
      categoryRates[category] = rate;
      console.log(`    ${category}: ${stats.completed}/${stats.total} (${Math.round(rate * 100)}%)`);
    }

    // ステップ5: 各タスクの完了率を更新
    console.log('\nステップ5: 各タスクの完了率を更新...');

    let updateCount = 0;
    for (const task of allTasks) {
      const category = task.fields['カテゴリ'];
      if (!category) continue;

      const rate = categoryRates[category] || 0;

      await client.request({
        method: 'PUT',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records/${task.record_id}`,
        data: {
          fields: {
            '完了率': rate
          }
        }
      });

      updateCount++;
      if (updateCount % 10 === 0) {
        process.stdout.write(`\r  進捗: ${updateCount}/${allTasks.length}`);
      }
      await new Promise(r => setTimeout(r, 50));
    }

    console.log(`\n  ✓ ${updateCount}件のタスクを更新`);

    console.log('\n=== 完了 ===');
    console.log('\nLarkBase URL: https://www.feishu.cn/base/' + appToken);
    console.log('\n【自動更新の設定方法】');
    console.log('LarkBaseで自動化を設定してください:');
    console.log('1. LarkBaseを開く');
    console.log('2. 右上「自動化」→「新規」');
    console.log('3. トリガー: レコード更新時');
    console.log('4. 条件: ステータスフィールドが変更');
    console.log('5. アクション: スクリプト実行またはWebhook');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
