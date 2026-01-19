/**
 * 数式の比較演算子を修正
 * == → = に変更
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
  console.log('=== 数式構文修正（== → =）===\n');

  try {
    // フィールド一覧を取得
    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    // 既存の計算フィールドを削除
    console.log('既存フィールドを削除...');
    const fieldsToDelete = ['カテゴリ総数', 'カテゴリ完了数', '完了率'];
    for (const fieldName of fieldsToDelete) {
      const field = fields.find(f => f.field_name === fieldName);
      if (field) {
        await client.request({
          method: 'DELETE',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${field.field_id}`
        });
        console.log(`  ✓ ${fieldName} を削除`);
        await new Promise(r => setTimeout(r, 500));
      }
    }

    await new Promise(r => setTimeout(r, 1000));

    // カテゴリ総数を作成（= を使用）
    console.log('\nカテゴリ総数を作成...');
    const totalFormula = '[タスク].COUNTIF(CurrentValue.[カテゴリ]=[カテゴリ])';
    console.log(`  数式: ${totalFormula}`);

    await client.request({
      method: 'POST',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
      data: {
        field_name: 'カテゴリ総数',
        type: 20,
        ui_type: 'Formula',
        property: {
          formula_expression: totalFormula
        }
      }
    });
    console.log('  ✓ 成功');

    await new Promise(r => setTimeout(r, 1000));

    // カテゴリ完了数を作成（= を使用、&& も確認が必要）
    console.log('\nカテゴリ完了数を作成...');
    // && も問題かもしれないので AND() 関数を試す
    const completedFormula = '[タスク].COUNTIF(AND(CurrentValue.[カテゴリ]=[カテゴリ],CurrentValue.[ステータス]="完了"))';
    console.log(`  数式: ${completedFormula}`);

    try {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: 'カテゴリ完了数',
          type: 20,
          ui_type: 'Formula',
          property: {
            formula_expression: completedFormula
          }
        }
      });
      console.log('  ✓ 成功');
    } catch (e) {
      console.log(`  ✗ 失敗: ${e.response?.data?.msg || e.message}`);

      // && 形式も試す（= に変更して）
      console.log('\n  代替案: && を使用...');
      const altFormula = '[タスク].COUNTIF(CurrentValue.[カテゴリ]=[カテゴリ]&&CurrentValue.[ステータス]="完了")';
      console.log(`  数式: ${altFormula}`);

      try {
        await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
          data: {
            field_name: 'カテゴリ完了数',
            type: 20,
            ui_type: 'Formula',
            property: {
              formula_expression: altFormula
            }
          }
        });
        console.log('  ✓ 成功');
      } catch (e2) {
        console.log(`  ✗ 失敗: ${e2.response?.data?.msg || e2.message}`);
      }
    }

    await new Promise(r => setTimeout(r, 1000));

    // フィールドIDを再取得
    const fieldsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields2 = fieldsRes2.data.items || [];

    const totalField = fields2.find(f => f.field_name === 'カテゴリ総数');
    const completedField = fields2.find(f => f.field_name === 'カテゴリ完了数');

    // 完了率を作成
    if (totalField && completedField) {
      console.log('\n完了率を作成...');
      const rateFormula = `IF($field[${totalField.field_id}]>0,$field[${completedField.field_id}]/$field[${totalField.field_id}],0)`;
      console.log(`  数式: ${rateFormula}`);

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
      console.log('  ✓ 成功');
    }

    await new Promise(r => setTimeout(r, 2000));

    // 結果を確認
    console.log('\n\n=== 結果確認 ===');
    const recordsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
      params: { page_size: 10 }
    });

    console.log('\nレコード:');
    recordsRes.data.items.forEach(r => {
      console.log(`  ${r.fields['タスク名']}`);
      console.log(`    カテゴリ: ${r.fields['カテゴリ']}, ステータス: ${r.fields['ステータス']}`);
      console.log(`    総数: ${r.fields['カテゴリ総数']}, 完了数: ${r.fields['カテゴリ完了数']}, 完了率: ${r.fields['完了率']}`);
    });

    console.log('\n=== 完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response?.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
