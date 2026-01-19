/**
 * 完了率計算用の数式フィールドを作成
 *
 * 正しい構文:
 * - カテゴリ総数: [タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ])
 * - カテゴリ完了数: [タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ]&&CurrentValue.[ステータス]=="完了")
 * - 完了率: IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)
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
  console.log('=== 完了率計算フィールド作成 ===\n');

  try {
    // ステップ1: 現在のフィールドを確認
    console.log('ステップ1: 現在のフィールドを確認...');

    const fieldsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    let fields = fieldsRes.data.items || [];

    // カテゴリ総数が存在するか確認
    const totalField = fields.find(f => f.field_name === 'カテゴリ総数');
    if (!totalField) {
      console.log('カテゴリ総数フィールドがありません。作成します...');

      await client.request({
        method: 'POST',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`,
        data: {
          field_name: 'カテゴリ総数',
          type: 20,
          ui_type: 'Formula',
          property: {
            formula_expression: '[タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ])'
          }
        }
      });
      console.log('✓ カテゴリ総数 作成成功');
      await new Promise(r => setTimeout(r, 1000));
    } else {
      console.log('✓ カテゴリ総数 既存');
    }

    // ステップ2: カテゴリ完了数を作成
    console.log('\nステップ2: カテゴリ完了数フィールドを作成...');

    // 既存のカテゴリ完了数を削除
    const existingCompleted = fields.find(f => f.field_name === 'カテゴリ完了数');
    if (existingCompleted) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${existingCompleted.field_id}`
      });
      console.log('  既存のカテゴリ完了数を削除');
      await new Promise(r => setTimeout(r, 500));
    }

    // 新規作成
    const completedFormula = '[タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ]&&CurrentValue.[ステータス]=="完了")';
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
      console.log('✓ カテゴリ完了数 作成成功');
    } catch (e) {
      console.log(`✗ カテゴリ完了数 作成失敗: ${e.message}`);
      if (e.response && e.response.data) {
        console.log(`  詳細: ${JSON.stringify(e.response.data)}`);
      }

      // 代替案: 単一選択フィールドの値比較を別の方法で試す
      console.log('\n  代替案を試します...');

      // ステータスが単一選択(type 3)の場合、値の比較方法が異なる可能性
      // 試行: 文字列 "完了" を使わず、選択肢の名前を直接参照
      const altFormula = '[タスク].COUNTIF(CurrentValue.[カテゴリ]==[カテゴリ] && CurrentValue.[ステータス]==[ステータス])';
      console.log(`  代替数式: ${altFormula}`);

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
        console.log('  ✓ 代替案で成功');
      } catch (e2) {
        console.log(`  ✗ 代替案も失敗: ${e2.message}`);
      }
    }

    await new Promise(r => setTimeout(r, 1000));

    // ステップ3: 完了率フィールドを作成
    console.log('\nステップ3: 完了率フィールドを作成...');

    // フィールド一覧を再取得
    const fieldsRes2 = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields`
    });
    fields = fieldsRes2.data.items || [];

    // 既存の完了率を削除
    const existingRate = fields.find(f => f.field_name === '完了率');
    if (existingRate) {
      await client.request({
        method: 'DELETE',
        url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/fields/${existingRate.field_id}`
      });
      console.log('  既存の完了率を削除');
      await new Promise(r => setTimeout(r, 500));
    }

    // 新規作成
    const rateFormula = 'IF([カテゴリ総数]>0,[カテゴリ完了数]/[カテゴリ総数],0)';
    console.log(`  数式: ${rateFormula}`);

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
            formula_expression: rateFormula
          }
        }
      });
      console.log('✓ 完了率 作成成功');
    } catch (e) {
      console.log(`✗ 完了率 作成失敗: ${e.message}`);
      if (e.response && e.response.data) {
        console.log(`  詳細: ${JSON.stringify(e.response.data)}`);
      }
    }

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
      if (f.property && f.property.formatter) {
        console.log(`    書式: ${f.property.formatter}`);
      }
    });

    // サンプルデータを取得して確認
    console.log('\n=== サンプルデータ確認 ===');
    const recordsRes = await client.request({
      method: 'GET',
      url: `/open-apis/bitable/v1/apps/${appToken}/tables/${taskTableId}/records`,
      params: { page_size: 5 }
    });
    const records = recordsRes.data.items || [];

    console.log('\n最初の5件のタスク:');
    records.forEach(r => {
      console.log(`  - ${r.fields['タスク名']}`);
      console.log(`    カテゴリ: ${r.fields['カテゴリ']}`);
      console.log(`    ステータス: ${r.fields['ステータス']}`);
      console.log(`    カテゴリ総数: ${r.fields['カテゴリ総数']}`);
      console.log(`    カテゴリ完了数: ${r.fields['カテゴリ完了数']}`);
      console.log(`    完了率: ${r.fields['完了率']}`);
    });

    console.log('\n=== 処理完了 ===');
    console.log('LarkBase URL: https://www.feishu.cn/base/' + appToken);

  } catch (error) {
    console.error('エラー:', error.message);
    if (error.response && error.response.data) {
      console.log('詳細:', JSON.stringify(error.response.data, null, 2));
    }
  }
}

main();
