/**
 * 期限超過・進捗遅延タスク検出スクリプト
 * Phase 2: 自動化
 *
 * 機能:
 *   - 期限超過タスクの検出
 *   - 進捗遅延タスク（期限3日以内で未着手）の検出
 *   - Lark Webhook経由での通知
 *
 * 使用方法:
 *   node scripts/check-overdue-tasks.cjs [--notify]
 *   --notify: Lark Webhookに通知を送信
 */

const lark = require('@larksuiteoapi/node-sdk');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const client = new lark.Client({
  appId: process.env.LARK_APP_ID,
  appSecret: process.env.LARK_APP_SECRET,
  appType: lark.AppType.SelfBuild,
  domain: process.env.LARK_DOMAIN === 'larksuite' ? lark.Domain.Lark : lark.Domain.Feishu,
});

const APP_TOKEN = process.env.LARK_BASE_APP_TOKEN;
const TASK_TABLE_ID = process.env.TASK_TABLE_ID;
const CUSTOMER_TABLE_ID = process.env.CUSTOMER_TABLE_ID;
const WEBHOOK_URL = process.env.LARK_WEBHOOK_URL;

const SEND_NOTIFICATION = process.argv.includes('--notify');

// ========================================
// メイン処理
// ========================================
async function main() {
  console.log('=== 期限超過・進捗遅延チェック ===\n');
  console.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  console.log(`通知送信: ${SEND_NOTIFICATION ? 'ON' : 'OFF'}\n`);

  try {
    // 顧客情報を取得
    const customers = await getCustomers();
    console.log(`アクティブ顧客数: ${customers.length}\n`);

    // 全タスクを取得
    const allTasks = await getAllTasks();
    console.log(`全タスク数: ${allTasks.length}\n`);

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayTime = today.getTime();

    const threeDaysLater = new Date(today);
    threeDaysLater.setDate(threeDaysLater.getDate() + 3);
    const threeDaysLaterTime = threeDaysLater.getTime();

    // タスクを分類
    const overdueTasks = [];      // 期限超過
    const urgentTasks = [];       // 期限3日以内で未着手
    const warningTasks = [];      // 期限7日以内で未着手

    for (const task of allTasks) {
      const status = task.fields['ステータス'];
      const dueDate = task.fields['期限'];

      // 完了・保留タスクはスキップ
      if (status === '完了' || status === '保留') continue;

      if (!dueDate) continue;

      const dueDateTime = typeof dueDate === 'number' ? dueDate : new Date(dueDate).getTime();

      if (dueDateTime < todayTime) {
        // 期限超過
        overdueTasks.push({
          ...task,
          daysOverdue: Math.floor((todayTime - dueDateTime) / (1000 * 60 * 60 * 24))
        });
      } else if (dueDateTime <= threeDaysLaterTime && status === '未着手') {
        // 期限3日以内で未着手
        urgentTasks.push({
          ...task,
          daysRemaining: Math.floor((dueDateTime - todayTime) / (1000 * 60 * 60 * 24))
        });
      } else if (dueDateTime <= todayTime + 7 * 24 * 60 * 60 * 1000 && status === '未着手') {
        // 期限7日以内で未着手
        warningTasks.push({
          ...task,
          daysRemaining: Math.floor((dueDateTime - todayTime) / (1000 * 60 * 60 * 24))
        });
      }
    }

    // 顧客情報をマッピング
    const customerMap = {};
    for (const c of customers) {
      customerMap[c.record_id] = c.fields['会社名'];
    }

    // 結果を表示
    console.log('========================================');
    console.log('【期限超過タスク】', overdueTasks.length, '件');
    console.log('========================================');
    for (const task of overdueTasks.sort((a, b) => b.daysOverdue - a.daysOverdue)) {
      const customerId = task.fields['顧客']?.[0]?.record_ids?.[0];
      const customerName = customerId ? (customerMap[customerId] || '不明') : '未設定';
      console.log(`  ❌ [${task.fields['WBS番号']}] ${task.fields['タスク名']}`);
      console.log(`     顧客: ${customerName} | ${task.daysOverdue}日超過 | 担当: ${task.fields['担当者'] || '-'}`);
    }

    console.log('\n========================================');
    console.log('【緊急タスク（3日以内・未着手）】', urgentTasks.length, '件');
    console.log('========================================');
    for (const task of urgentTasks.sort((a, b) => a.daysRemaining - b.daysRemaining)) {
      const customerId = task.fields['顧客']?.[0]?.record_ids?.[0];
      const customerName = customerId ? (customerMap[customerId] || '不明') : '未設定';
      console.log(`  ⚠️  [${task.fields['WBS番号']}] ${task.fields['タスク名']}`);
      console.log(`     顧客: ${customerName} | 残り${task.daysRemaining}日 | 担当: ${task.fields['担当者'] || '-'}`);
    }

    console.log('\n========================================');
    console.log('【注意タスク（7日以内・未着手）】', warningTasks.length, '件');
    console.log('========================================');
    for (const task of warningTasks.sort((a, b) => a.daysRemaining - b.daysRemaining)) {
      const customerId = task.fields['顧客']?.[0]?.record_ids?.[0];
      const customerName = customerId ? (customerMap[customerId] || '不明') : '未設定';
      console.log(`  📋 [${task.fields['WBS番号']}] ${task.fields['タスク名']}`);
      console.log(`     顧客: ${customerName} | 残り${task.daysRemaining}日 | 担当: ${task.fields['担当者'] || '-'}`);
    }

    // サマリー
    console.log('\n========================================');
    console.log('【サマリー】');
    console.log('========================================');
    console.log(`  期限超過:    ${overdueTasks.length} 件`);
    console.log(`  緊急（3日）: ${urgentTasks.length} 件`);
    console.log(`  注意（7日）: ${warningTasks.length} 件`);

    // 通知送信
    if (SEND_NOTIFICATION && (overdueTasks.length > 0 || urgentTasks.length > 0)) {
      console.log('\n通知を送信中...');
      await sendNotification(overdueTasks, urgentTasks, customerMap);
      console.log('✓ 通知送信完了');
    }

    // 結果をJSONで出力（他システム連携用）
    const result = {
      timestamp: new Date().toISOString(),
      summary: {
        overdue: overdueTasks.length,
        urgent: urgentTasks.length,
        warning: warningTasks.length
      },
      overdueTasks: overdueTasks.map(t => ({
        wbs: t.fields['WBS番号'],
        name: t.fields['タスク名'],
        customer: customerMap[t.fields['顧客']?.[0]?.record_ids?.[0]] || '未設定',
        daysOverdue: t.daysOverdue
      })),
      urgentTasks: urgentTasks.map(t => ({
        wbs: t.fields['WBS番号'],
        name: t.fields['タスク名'],
        customer: customerMap[t.fields['顧客']?.[0]?.record_ids?.[0]] || '未設定',
        daysRemaining: t.daysRemaining
      }))
    };

    return result;

  } catch (error) {
    console.error('\nエラー:', error.message);
    process.exit(1);
  }
}

// ========================================
// 顧客情報取得
// ========================================
async function getCustomers() {
  const res = await client.bitable.appTableRecord.list({
    path: { app_token: APP_TOKEN, table_id: CUSTOMER_TABLE_ID },
    params: { page_size: 100 }
  });

  if (res.code !== 0) {
    throw new Error(`顧客取得エラー: ${res.msg}`);
  }

  return res.data.items.filter(c =>
    c.fields['ステータス'] === '進行中' || c.fields['ステータス'] === '準備中'
  );
}

// ========================================
// 全タスク取得
// ========================================
async function getAllTasks() {
  let allTasks = [];
  let pageToken = null;

  do {
    const params = { page_size: 500 };
    if (pageToken) params.page_token = pageToken;

    const res = await client.bitable.appTableRecord.list({
      path: { app_token: APP_TOKEN, table_id: TASK_TABLE_ID },
      params
    });

    if (res.code !== 0) {
      throw new Error(`タスク取得エラー: ${res.msg}`);
    }

    allTasks = allTasks.concat(res.data.items || []);
    pageToken = res.data.page_token;

  } while (pageToken);

  return allTasks;
}

// ========================================
// Lark Webhook通知
// ========================================
async function sendNotification(overdueTasks, urgentTasks, customerMap) {
  if (!WEBHOOK_URL) {
    console.log('  ⚠ LARK_WEBHOOK_URLが設定されていません');
    return;
  }

  // 顧客別にグループ化
  const overdueByCustomer = groupByCustomer(overdueTasks, customerMap);
  const urgentByCustomer = groupByCustomer(urgentTasks, customerMap);

  let content = `📊 **タスクアラート** (${new Date().toLocaleDateString('ja-JP')})\n\n`;

  if (overdueTasks.length > 0) {
    content += `🔴 **期限超過: ${overdueTasks.length}件**\n`;
    for (const [customer, tasks] of Object.entries(overdueByCustomer)) {
      content += `\n**${customer}** (${tasks.length}件)\n`;
      for (const t of tasks.slice(0, 5)) {
        content += `• [${t.fields['WBS番号']}] ${t.fields['タスク名']} (${t.daysOverdue}日超過)\n`;
      }
      if (tasks.length > 5) {
        content += `  ...他${tasks.length - 5}件\n`;
      }
    }
  }

  if (urgentTasks.length > 0) {
    content += `\n🟡 **緊急（3日以内）: ${urgentTasks.length}件**\n`;
    for (const [customer, tasks] of Object.entries(urgentByCustomer)) {
      content += `\n**${customer}** (${tasks.length}件)\n`;
      for (const t of tasks.slice(0, 5)) {
        content += `• [${t.fields['WBS番号']}] ${t.fields['タスク名']} (残${t.daysRemaining}日)\n`;
      }
      if (tasks.length > 5) {
        content += `  ...他${tasks.length - 5}件\n`;
      }
    }
  }

  content += `\n[📋 LarkBaseで確認](https://www.larksuite.com/base/${APP_TOKEN})`;

  // Webhook送信
  const response = await fetch(WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      msg_type: 'text',
      content: { text: content }
    })
  });

  if (!response.ok) {
    console.log(`  ⚠ Webhook送信エラー: ${response.status}`);
  }
}

function groupByCustomer(tasks, customerMap) {
  const grouped = {};
  for (const task of tasks) {
    const customerId = task.fields['顧客']?.[0]?.record_ids?.[0];
    const customerName = customerId ? (customerMap[customerId] || '不明') : '未設定';
    if (!grouped[customerName]) grouped[customerName] = [];
    grouped[customerName].push(task);
  }
  return grouped;
}

main();
