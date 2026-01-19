/**
 * 完了率数式をUI形式で修正
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

async function testFormula(formula, description) {
  console.log(`\n${description}`);
  console.log(`  数式: ${formula}`);

  const fieldsRes = await client.request({
    method: 'GET',
    url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
  });
  const rateField = fieldsRes.data.items.find(f => f.field_name === '完了率');
  if (rateField) {
    await client.request({
      method: 'DELETE',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${rateField.field_id}`
    });
    await new Promise(r => setTimeout(r, 500));
  }

  try {
    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: '完了率',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: formula
        }
      }
    });
    console.log('  ✓ 作成成功');

    await new Promise(r => setTimeout(r, 2000));

    const recordsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
      params: { page_size: 3 }
    });

    console.log('  結果:');
    recordsRes.data.items.forEach(r => {
      const total = r.fields['カテゴリ総数'];
      const completed = r.fields['カテゴリ完了数'];
      const rate = r.fields['完了率'];
      console.log(`    ${r.fields['タスク名']}: 総数=${total}, 完了=${completed}, 完了率=${rate}`);
    });

    return rate !== null;
  } catch (e) {
    console.log(`  ✗ 失敗: ${e.response?.data?.msg || e.message}`);
    return false;
  }
}

async function main() {
  console.log('=== 完了率数式修正 ===');

  // フィールドIDを取得
  const fieldsRes = await client.request({
    method: 'GET',
    url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
  });
  const fields = fieldsRes.data.items || [];

  const totalField = fields.find(f => f.field_name === 'カテゴリ総数');
  const completedField = fields.find(f => f.field_name === 'カテゴリ完了数');

  console.log(`カテゴリ総数 ID: ${totalField?.field_id}`);
  console.log(`カテゴリ完了数 ID: ${completedField?.field_id}`);

  // テスト1: UI形式（フィールド名）
  await testFormula('IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)', 'テスト1: UI形式（フィールド名参照）');

  // テスト2: 単一イコールでの比較
  await testFormula('IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)', 'テスト2: 単一イコール比較');

  // テスト3: bitable形式
  await testFormula(
    `IF(bitable::$table[${taskTableId}].$field[${totalField.field_id}]>0,bitable::$table[${taskTableId}].$field[${completedField.field_id}]/bitable::$table[${taskTableId}].$field[${totalField.field_id}],0)`,
    'テスト3: bitable形式'
  );

  // テスト4: シンプルな割り算のみ
  await testFormula('[カテゴリ完了数]/[カテゴリ総数]', 'テスト4: シンプルな割り算');

  // テスト5: DIVIDE関数
  await testFormula('DIVIDE([カテゴリ完了数],[カテゴリ総数])', 'テスト5: DIVIDE関数');

  // テスト6: 数値変換
  await testFormula('VALUE([カテゴリ完了数])/VALUE([カテゴリ総数])', 'テスト6: VALUE関数で変換');

  console.log('\n=== 完了 ===');
  console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);
}

main().catch(e => console.error(e.message));
