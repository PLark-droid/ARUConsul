/**
 * 完了率フィールドを作成（フィールドID参照版）
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
  console.log('=== 完了率フィールド作成 ===\n');

  try {
    // フィールド一覧を取得
    console.log('フィールド一覧を取得...');
    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const fields = fieldsRes.data.items || [];

    const totalField = fields.find(f => f.field_name === 'カテゴリ総数');
    const completedField = fields.find(f => f.field_name === 'カテゴリ完了数');

    if (!totalField || !completedField) {
      console.log('エラー: カテゴリ総数またはカテゴリ完了数フィールドが見つかりません');
      return;
    }

    console.log(`カテゴリ総数 ID: ${totalField.field_id}`);
    console.log(`カテゴリ完了数 ID: ${completedField.field_id}`);

    // 既存の完了率を削除
    const existingRate = fields.find(f => f.field_name === '完了率');
    if (existingRate) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${existingRate.field_id}`
      });
      console.log('既存の完了率を削除しました');
      await new Promise(r => setTimeout(r, 500));
    }

    // 試行1: $field[fieldId]形式
    console.log('\n試行1: $field[fieldId]形式で完了率を作成...');
    const formula1 = `IF($field[${totalField.field_id}]>0,$field[${completedField.field_id}]/$field[${totalField.field_id}],0)`;
    console.log(`  数式: ${formula1}`);

    try {
      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: '完了率',
          type: 20,
          ui_type: 'Formula',
          property: {
            formatter: '0%',
            formula_expression: formula1
          }
        }
      });
      console.log('  ✓ 成功!');
    } catch (e) {
      console.log(`  ✗ 失敗: ${e.response?.data?.msg || e.message}`);

      // 試行2: フィールド名を直接使用（スペースなし）
      console.log('\n試行2: フィールド名参照（別形式）...');
      const formula2 = 'IF(カテゴリ総数>0,カテゴリ完了数/カテゴリ総数,0)';
      console.log(`  数式: ${formula2}`);

      try {
        await client.request({
          method: 'POST',
          url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
          data: {
            field_name: '完了率',
            type: 20,
            ui_type: 'Formula',
            property: {
              formatter: '0%',
              formula_expression: formula2
            }
          }
        });
        console.log('  ✓ 成功!');
      } catch (e2) {
        console.log(`  ✗ 失敗: ${e2.response?.data?.msg || e2.message}`);

        // 試行3: フォーマッタなしで試す
        console.log('\n試行3: フォーマッタなしで作成...');
        const formula3 = `IF($field[${totalField.field_id}]>0,$field[${completedField.field_id}]/$field[${totalField.field_id}],0)`;
        console.log(`  数式: ${formula3}`);

        try {
          await client.request({
            method: 'POST',
            url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
            data: {
              field_name: '完了率',
              type: 20,
              ui_type: 'Formula',
              property: {
                formula_expression: formula3
              }
            }
          });
          console.log('  ✓ 成功!');
        } catch (e3) {
          console.log(`  ✗ 失敗: ${e3.response?.data?.msg || e3.message}`);

          // 試行4: bitable::形式
          console.log('\n試行4: bitable::$table形式...');
          const formula4 = `IF(bitable::$table[${taskTableId}].$field[${totalField.field_id}]>0,bitable::$table[${taskTableId}].$field[${completedField.field_id}]/bitable::$table[${taskTableId}].$field[${totalField.field_id}],0)`;
          console.log(`  数式: ${formula4}`);

          try {
            await client.request({
              method: 'POST',
              url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
              data: {
                field_name: '完了率',
                type: 20,
                ui_type: 'Formula',
                property: {
                  formula_expression: formula4
                }
              }
            });
            console.log('  ✓ 成功!');
          } catch (e4) {
            console.log(`  ✗ 失敗: ${e4.response?.data?.msg || e4.message}`);

            // 試行5: 数値フィールドとして作成し、手動で数式に変換する案内
            console.log('\n試行5: 数値フィールドとして作成...');
            try {
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
              console.log('  ✓ 数値フィールドとして作成成功');
              console.log('\n  【手動設定が必要です】');
              console.log('  LarkBaseで「完了率」列を数式フィールドに変更し、以下の数式を入力:');
              console.log('  IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)');
            } catch (e5) {
              console.log(`  ✗ 失敗: ${e5.response?.data?.msg || e5.message}`);
            }
          }
        }
      }
    }

    // 最終確認
    console.log('\n=== 最終確認 ===');
    const finalFieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    const finalFields = finalFieldsRes.data.items || [];

    console.log('\n計算用フィールド:');
    ['カテゴリ総数', 'カテゴリ完了数', '完了率'].forEach(name => {
      const f = finalFields.find(field => field.field_name === name);
      if (f) {
        console.log(`  - ${f.field_name} (type: ${f.type})`);
        if (f.property?.formula_expression) {
          console.log(`    数式: ${f.property.formula_expression}`);
        }
      }
    });

    console.log('\n=== 処理完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response?.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
