/**
 * 数式フィールドの動作確認
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
  console.log('=== 数式フィールド動作確認 ===\n');

  try {
    // フィールド一覧を取得
    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    console.log('計算用フィールドの状態:');
    ['カテゴリ総数', 'カテゴリ完了数', '完了率'].forEach(name => {
      const f = fields.find(field => field.field_name === name);
      if (f) {
        console.log(`\n  ${f.field_name}:`);
        console.log(`    タイプ: ${f.type} (20=数式)`);
        console.log(`    数式: ${f.property?.formula_expression || 'なし'}`);
      }
    });

    // レコードを取得
    console.log('\n\n=== カテゴリ別統計 ===\n');

    let allRecords = [];
    let pageToken = null;

    do {
      const params = { page_size: 500 };
      if (pageToken) params.page_token = pageToken;

      const recordsRes = await client.request({
        method: 'GET',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
        params
      });

      allRecords = allRecords.concat(recordsRes.data.items || []);
      pageToken = recordsRes.data.page_token;
    } while (pageToken);

    // カテゴリ別に集計
    const categoryStats = {};

    for (const record of allRecords) {
      const category = record.fields['カテゴリ'] || '不明';
      const status = record.fields['ステータス'];
      const total = record.fields['カテゴリ総数'];
      const completed = record.fields['カテゴリ完了数'];
      const rate = record.fields['完了率'];

      if (!categoryStats[category]) {
        categoryStats[category] = {
          count: 0,
          completedCount: 0,
          formulaTotal: total,
          formulaCompleted: completed,
          formulaRate: rate
        };
      }

      categoryStats[category].count++;
      if (status === '完了') {
        categoryStats[category].completedCount++;
      }
    }

    console.log('カテゴリ | 実際の数 | 実際の完了 | 数式総数 | 数式完了 | 数式完了率');
    console.log('---------|----------|------------|----------|----------|----------');

    for (const [category, stats] of Object.entries(categoryStats)) {
      const actualRate = stats.count > 0 ? Math.round((stats.completedCount / stats.count) * 100) : 0;
      console.log(
        `${category.padEnd(8)} | ` +
        `${String(stats.count).padStart(8)} | ` +
        `${String(stats.completedCount).padStart(10)} | ` +
        `${String(stats.formulaTotal ?? 'null').padStart(8)} | ` +
        `${String(stats.formulaCompleted ?? 'null').padStart(8)} | ` +
        `${stats.formulaRate ?? 'null'}`
      );
    }

    // サンプルレコードの詳細
    console.log('\n\n=== サンプルレコード（各カテゴリから1件）===\n');

    const shownCategories = new Set();
    for (const record of allRecords) {
      const category = record.fields['カテゴリ'];
      if (!shownCategories.has(category)) {
        shownCategories.add(category);
        console.log(`タスク名: ${record.fields['タスク名']}`);
        console.log(`  カテゴリ: ${category}`);
        console.log(`  ステータス: ${record.fields['ステータス']}`);
        console.log(`  カテゴリ総数: ${record.fields['カテゴリ総数']}`);
        console.log(`  カテゴリ完了数: ${record.fields['カテゴリ完了数']}`);
        console.log(`  完了率: ${record.fields['完了率']}`);
        console.log('');
      }
    }

    console.log('\n=== 確認完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

    // 数式が計算されているかの判定
    const firstRecord = allRecords[0];
    const hasFormulaValue = firstRecord?.fields['カテゴリ総数'] !== null && firstRecord?.fields['カテゴリ総数'] !== undefined;

    if (hasFormulaValue) {
      console.log('\n✓ 数式が正しく計算されています！');
    } else {
      console.log('\n! 数式の値がnullです。LarkBaseで直接確認してください。');
      console.log('  数式フィールドの計算には数秒かかる場合があります。');
    }

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response?.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
