/**
 * フィールド参照の数式構文をテスト
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

async function deleteTestField() {
  const fieldsRes = await client.request({
    method: 'GET',
    url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
  });
  const testField = fieldsRes.data.items.find(f => f.field_name === 'テスト数式');
  if (testField) {
    await client.request({
      method: 'DELETE',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${testField.field_id}`
    });
    await new Promise(r => setTimeout(r, 500));
  }
}

async function createAndTest(formula, description) {
  console.log(`\n${description}`);
  console.log(`  数式: ${formula}`);

  await deleteTestField();

  try {
    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: 'テスト数式',
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
      console.log(`    ${r.fields['タスク名']}: ${r.fields['テスト数式']}`);
    });

    return true;
  } catch (e) {
    console.log(`  ✗ 失敗: ${e.response?.data?.msg || e.message}`);
    return false;
  }
}

async function main() {
  console.log('=== フィールド参照構文テスト ===');

  // フィールドIDを取得
  const fieldsRes = await client.request({
    method: 'GET',
    url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
  });
  const fields = fieldsRes.data.items || [];

  const wbsField = fields.find(f => f.field_name === 'WBS番号');
  const taskNameField = fields.find(f => f.field_name === 'タスク名');
  const categoryField = fields.find(f => f.field_name === 'カテゴリ');
  const statusField = fields.find(f => f.field_name === 'ステータス');

  console.log('\nフィールド情報:');
  console.log(`  WBS番号: ${wbsField?.field_id} (type: ${wbsField?.type})`);
  console.log(`  タスク名: ${taskNameField?.field_id} (type: ${taskNameField?.type})`);
  console.log(`  カテゴリ: ${categoryField?.field_id} (type: ${categoryField?.type})`);
  console.log(`  ステータス: ${statusField?.field_id} (type: ${statusField?.type})`);

  // テスト1: テキストフィールドをUI形式で参照
  await createAndTest('[タスク名]', 'テスト1: UI形式でテキストフィールド参照');

  // テスト2: テキストフィールドを$field形式で参照
  await createAndTest(`$field[${taskNameField.field_id}]`, 'テスト2: $field形式でテキストフィールド参照');

  // テスト3: テキストフィールドをbitable形式で参照
  await createAndTest(`bitable::$table[${taskTableId}].$field[${taskNameField.field_id}]`, 'テスト3: bitable形式でテキストフィールド参照');

  // テスト4: COUNTA関数（テーブル参照）
  await createAndTest('[タスク].COUNTA()', 'テスト4: COUNTA関数');

  // テスト5: 単一選択フィールドを文字列比較
  await createAndTest('IF([ステータス]="完了",1,0)', 'テスト5: 単一選択を文字列比較（=）');

  // テスト6: 単一選択フィールドをダブルクオートで比較
  await createAndTest('IF([ステータス]=="完了",1,0)', 'テスト6: 単一選択を文字列比較（==）');

  // テスト7: TOTEXTで変換してから比較
  await createAndTest('IF(TOTEXT([ステータス])=="完了",1,0)', 'テスト7: TOTEXTで変換して比較');

  // テスト8: テキストフィールド（カテゴリ）を比較
  await createAndTest('IF([カテゴリ]=="法人関連",1,0)', 'テスト8: テキストフィールドを比較');

  // テスト9: LEN関数
  await createAndTest('LEN([タスク名])', 'テスト9: LEN関数');

  // クリーンアップ
  await deleteTestField();

  console.log('\n=== テスト完了 ===');
}

main().catch(e => console.error(e.message));
