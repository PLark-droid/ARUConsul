/**
 * 完了率を数式フィールドで実現
 * - カテゴリ総数・カテゴリ完了数を集計
 * - 完了率 = カテゴリ完了数 / カテゴリ総数 (数式フィールド)
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
const taskTableId = 'tblHiAuZsUyWxlAY';

async function main() {
  console.log('=== 完了率更新 ===\n');

  try {
    // ステップ1: フィールドを確認
    console.log('ステップ1: フィールドを確認...');

    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    let fields = fieldsRes.data.items || [];

    // カテゴリ総数フィールドがなければ作成
    let totalField = fields.find(f => f.field_name === 'カテゴリ総数');
    if (!totalField) {
      const res = await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: { field_name: 'カテゴリ総数', type: 2 }
      });
      totalField = res.data.field;
      console.log('  ✓ カテゴリ総数フィールド作成');
      await new Promise(r => setTimeout(r, 300));
    }

    // カテゴリ完了数フィールドがなければ作成
    let completedField = fields.find(f => f.field_name === 'カテゴリ完了数');
    if (!completedField) {
      const res = await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: { field_name: 'カテゴリ完了数', type: 2 }
      });
      completedField = res.data.field;
      console.log('  ✓ カテゴリ完了数フィールド作成');
      await new Promise(r => setTimeout(r, 300));
    }

    // 完了率フィールドがなければ作成（数式）
    let rateField = fields.find(f => f.field_name === '完了率');
    if (!rateField) {
      // フィールドIDを再取得
      const fieldsRes2 = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
      });
      fields = fieldsRes2.data.items || [];
      totalField = fields.find(f => f.field_name === 'カテゴリ総数');
      completedField = fields.find(f => f.field_name === 'カテゴリ完了数');

      const formulaExpression = `IF(bitable::$table[${taskTableId}].$field[${totalField.field_id}]>0,bitable::$table[${taskTableId}].$field[${completedField.field_id}]/bitable::$table[${taskTableId}].$field[${totalField.field_id}],0)`;

      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: '完了率',
          type: 20,
          ui_type: 'Formula',
          property: { formula_expression: formulaExpression }
        }
      });
      console.log('  ✓ 完了率フィールド作成（数式）');
      await new Promise(r => setTimeout(r, 300));
    }

    // ステップ2: 全タスクを取得
    console.log('\nステップ2: タスクを取得...');

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

    console.log(`  ✓ ${allTasks.length}件のタスク`);

    // ステップ3: カテゴリごとに集計
    console.log('\nステップ3: カテゴリごとに集計...');

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

    console.log('\n  カテゴリ別:');
    for (const [cat, stats] of Object.entries(categoryStats)) {
      const rate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
      console.log(`    ${cat}: ${stats.completed}/${stats.total} (${rate}%)`);
    }

    // ステップ4: 各タスクを更新
    console.log('\nステップ4: 各タスクを更新...');

    let updateCount = 0;
    for (const task of allTasks) {
      const category = task.fields['カテゴリ'];
      if (!category) continue;

      const stats = categoryStats[category];

      await client.request({
        method: 'PUT',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records/${task.record_id}`,
        data: {
          fields: {
            'カテゴリ総数': stats.total,
            'カテゴリ完了数': stats.completed
          }
        }
      });

      updateCount++;
      if (updateCount % 10 === 0) {
        process.stdout.write(`\r  進捗: ${updateCount}/${allTasks.length}`);
      }
      await new Promise(r => setTimeout(r, 50));
    }

    console.log(`\n  ✓ ${updateCount}件更新`);

    console.log('\n=== 完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);
    console.log('\n完了率は数式フィールドで自動計算されます。');
    console.log('ステータス変更後は npm run update-completion を実行してください。');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
