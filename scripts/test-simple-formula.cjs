/**
 * シンプルな数式でテスト
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
  console.log('=== シンプル数式テスト ===\n');

  try {
    // テスト用数式フィールドを作成（単純な定数）
    console.log('テスト1: 定数を返す数式...');

    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    // 既存のテストフィールドを削除
    const testField = fields.find(f => f.field_name === 'テスト数式');
    if (testField) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${testField.field_id}`
      });
      console.log('  既存のテストフィールドを削除');
      await new Promise(r => setTimeout(r, 500));
    }

    // 定数を返す数式
    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: 'テスト数式',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: '100'
        }
      }
    });
    console.log('  ✓ テスト数式フィールド作成（数式: 100）');

    await new Promise(r => setTimeout(r, 2000));

    // 値を確認
    const recordsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
      params: { page_size: 3 }
    });

    console.log('\nレコード確認:');
    recordsRes.data.items.forEach(r => {
      console.log(`  ${r.fields['タスク名']}: テスト数式=${r.fields['テスト数式']}`);
    });

    // テスト2: フィールド参照の数式
    console.log('\n\nテスト2: 既存フィールドを参照する数式...');

    // テストフィールドを削除
    const fieldsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const testField2 = fieldsRes2.data.items.find(f => f.field_name === 'テスト数式');
    if (testField2) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${testField2.field_id}`
      });
      await new Promise(r => setTimeout(r, 500));
    }

    // フィールドを参照する数式
    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: 'テスト数式',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: 'IF([ステータス]=="完了",1,0)'
        }
      }
    });
    console.log('  ✓ テスト数式フィールド作成（数式: IF([ステータス]=="完了",1,0)）');

    await new Promise(r => setTimeout(r, 2000));

    // 値を確認
    const recordsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
      params: { page_size: 5 }
    });

    console.log('\nレコード確認:');
    recordsRes2.data.items.forEach(r => {
      console.log(`  ${r.fields['タスク名']}: ステータス=${r.fields['ステータス']}, テスト数式=${r.fields['テスト数式']}`);
    });

    // テストフィールドを削除
    const fieldsRes3 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const testField3 = fieldsRes3.data.items.find(f => f.field_name === 'テスト数式');
    if (testField3) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${testField3.field_id}`
      });
      console.log('\nテストフィールドを削除しました');
    }

    console.log('\n=== テスト完了 ===');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response?.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
