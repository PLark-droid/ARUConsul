/**
 * 数式フィールドを正しい構文で再作成
 *
 * 正しいCOUNTIF構文（ドキュメント確認済み）:
 * - [テーブル名].COUNTIF(CurrentValue.[フィールド名]==[フィールド名])
 * - API: bitable::$table[tableId].COUNTIF(CurrentValue.$field[fieldId]==$field[fieldId])
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
  console.log('=== 数式フィールド修正 ===\n');

  try {
    // ステップ1: 現在のフィールドを確認
    console.log('ステップ1: 現在のフィールドを確認...');

    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    console.log('\n現在のフィールド一覧:');
    fields.forEach(f => {
      console.log(`  - ${f.field_name} (type: ${f.type}, id: ${f.field_id})`);
      if (f.property && f.property.formula_expression) {
        console.log(`    数式: ${f.property.formula_expression}`);
      }
    });

    // 重要なフィールドIDを取得
    const categoryField = fields.find(f => f.field_name === 'カテゴリ');
    const statusField = fields.find(f => f.field_name === 'ステータス');

    if (!categoryField) {
      console.log('\nエラー: カテゴリフィールドが見つかりません');
      return;
    }
    if (!statusField) {
      console.log('\nエラー: ステータスフィールドが見つかりません');
      return;
    }

    console.log(`\nカテゴリフィールドID: ${categoryField.field_id}`);
    console.log(`ステータスフィールドID: ${statusField.field_id}`);

    // ステップ2: 既存の計算フィールドを削除
    console.log('\nステップ2: 既存の計算フィールドを削除...');

    const fieldsToDelete = ['完了率', 'カテゴリ総数', 'カテゴリ完了数'];
    for (const fieldName of fieldsToDelete) {
      const field = fields.find(f => f.field_name === fieldName);
      if (field) {
        try {
          await client.request({
            method: 'DELETE',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${field.field_id}`
          });
          console.log(`  ✓ ${fieldName} を削除`);
        } catch (e) {
          console.log(`  ! ${fieldName} 削除失敗: ${e.message}`);
        }
        await new Promise(r => setTimeout(r, 500));
      }
    }

    await new Promise(r => setTimeout(r, 1000));

    // ステップ3: カテゴリ総数を作成（COUNTIFで同じカテゴリのレコード数をカウント）
    console.log('\nステップ3: カテゴリ総数フィールドを作成...');

    // 方法1: シンプルなフィールド参照形式を試す
    // ドキュメントによると: [テーブル名].COUNTIF(CurrentValue.[フィールド名]==[フィールド名])
    // API形式: 単純な $field[fieldId] 参照

    // まず、シンプルな形式を試す
    const totalFormula1 = `[タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ])`;

    console.log('試行1: UI形式の数式');
    console.log(`  数式: ${totalFormula1}`);

    try {
      const res = await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: 'カテゴリ総数',
          type: 20,
          ui_type: 'Formula',
          property: {
            formula_expression: totalFormula1
          }
        }
      });
      console.log('  ✓ 成功!');
      console.log(`  作成されたフィールドID: ${res.data.field.field_id}`);
    } catch (e) {
      console.log(`  ✗ 失敗: ${e.message}`);
      if (e.response && e.response.data) {
        console.log(`  詳細: ${JSON.stringify(e.response.data)}`);
      }

      // 方法2: API形式を試す
      console.log('\n試行2: API形式の数式（$field参照）');
      const totalFormula2 = `$field[${categoryField.field_id}].COUNTIF(CurrentValue==$field[${categoryField.field_id}])`;
      console.log(`  数式: ${totalFormula2}`);

      try {
        const res = await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
          data: {
            field_name: 'カテゴリ総数',
            type: 20,
            ui_type: 'Formula',
            property: {
              formula_expression: totalFormula2
            }
          }
        });
        console.log('  ✓ 成功!');
      } catch (e2) {
        console.log(`  ✗ 失敗: ${e2.message}`);

        // 方法3: bitable::$table形式
        console.log('\n試行3: bitable::$table形式');
        const totalFormula3 = `bitable::$table[${taskTableId}].COUNTIF(CurrentValue.$field[${categoryField.field_id}]==$field[${categoryField.field_id}])`;
        console.log(`  数式: ${totalFormula3}`);

        try {
          const res = await client.request({
            method: 'POST',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
            data: {
              field_name: 'カテゴリ総数',
              type: 20,
              ui_type: 'Formula',
              property: {
                formula_expression: totalFormula3
              }
            }
          });
          console.log('  ✓ 成功!');
        } catch (e3) {
          console.log(`  ✗ 失敗: ${e3.message}`);

          // 方法4: 数値フィールドとして作成し、スクリプトで更新
          console.log('\n試行4: 数値フィールドとして作成');
          try {
            await client.request({
              method: 'POST',
              url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
              data: {
                field_name: 'カテゴリ総数',
                type: 2  // 数値型
              }
            });
            console.log('  ✓ 数値フィールドとして作成成功');
            console.log('  注: 数式が使えない場合は手動で数式に変更してください');
          } catch (e4) {
            console.log(`  ✗ 失敗: ${e4.message}`);
          }
        }
      }
    }

    await new Promise(r => setTimeout(r, 1000));

    // 最終確認
    console.log('\n=== 最終確認 ===');
    const finalFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const finalFields = finalFieldsRes.data.items || [];

    console.log('\n現在のフィールド:');
    finalFields.forEach(f => {
      console.log(`  - ${f.field_name} (type: ${f.type})`);
      if (f.property && f.property.formula_expression) {
        console.log(`    数式: ${f.property.formula_expression}`);
      }
    });

    console.log('\n=== 処理完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);
    console.log('\n手動での数式設定方法:');
    console.log('1. LarkBaseを開く');
    console.log('2. タスクテーブルで「カテゴリ総数」列のヘッダーをクリック');
    console.log('3. フィールドタイプを「数式」に変更');
    console.log('4. 以下の数式を入力:');
    console.log('   カテゴリ総数: [タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ])');
    console.log('   カテゴリ完了数: [タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ]&&CurrentValue.[ステータス]=="完了")');
    console.log('   完了率: IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)');

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
