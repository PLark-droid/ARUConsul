/**
 * 最終セットアップ - 完了率の正しい数式
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
  console.log('=== 最終セットアップ ===\n');

  try {
    // フィールド一覧を取得
    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    // 完了率を削除して再作成（IF文で0除算対応）
    const rateField = fields.find(f => f.field_name === '完了率');
    if (rateField) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${rateField.field_id}`
      });
      console.log('既存の完了率を削除');
      await new Promise(r => setTimeout(r, 500));
    }

    // 完了率を作成（UI形式、IF文で0除算対応）
    const rateFormula = 'IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)';
    console.log(`完了率数式: ${rateFormula}`);

    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: '完了率',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: rateFormula
        }
      }
    });
    console.log('✓ 完了率を作成');

    await new Promise(r => setTimeout(r, 2000));

    // 最終確認
    console.log('\n=== カテゴリ別完了率 ===\n');

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

    // カテゴリごとに集計
    const categoryStats = {};
    for (const record of allRecords) {
      const category = record.fields['カテゴリ'] || '不明';
      if (!categoryStats[category]) {
        categoryStats[category] = {
          total: record.fields['カテゴリ総数'],
          completed: record.fields['カテゴリ完了数'],
          rate: record.fields['完了率']
        };
      }
    }

    console.log('カテゴリ        | 総数 | 完了 | 完了率');
    console.log('----------------|------|------|-------');
    for (const [category, stats] of Object.entries(categoryStats)) {
      const ratePercent = stats.rate !== null ? `${Math.round(stats.rate * 100)}%` : 'N/A';
      console.log(
        `${category.padEnd(15)} | ${String(stats.total).padStart(4)} | ${String(stats.completed).padStart(4)} | ${ratePercent}`
      );
    }

    console.log('\n\n=== 数式フィールドの最終状態 ===\n');
    const finalFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const finalFields = finalFieldsRes.data.items || [];

    ['カテゴリ総数', 'カテゴリ完了数', '完了率'].forEach(name => {
      const f = finalFields.find(field => field.field_name === name);
      if (f) {
        console.log(`${f.field_name}:`);
        console.log(`  数式: ${f.property?.formula_expression}`);
      }
    });

    console.log('\n\n=== セットアップ完了！ ===');
    console.log('\nLarkBase URL: https://www.feishu.cn/base/' + appToken);
    console.log('\n【自動計算の仕組み】');
    console.log('1. カテゴリ総数 = COUNTIF(同じカテゴリのタスク数)');
    console.log('2. カテゴリ完了数 = COUNTIF(同じカテゴリで完了のタスク数)');
    console.log('3. 完了率 = カテゴリ完了数 / カテゴリ総数');
    console.log('\nステータスを「完了」に変更すると、自動的に完了率が再計算されます！');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response?.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
