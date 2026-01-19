/**
 * 週次進捗レポート生成スクリプト
 * Phase 2: 自動化
 *
 * 機能:
 *   - 全顧客の進捗サマリー
 *   - 顧客別進捗率
 *   - 今週完了タスク
 *   - 来週期限タスク
 *   - Lark Webhook経由での送信
 *
 * 使用方法:
 *   node scripts/weekly-report.cjs [--notify]
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

async function main() {
  console.log('=== 週次進捗レポート ===\n');
  const reportDate = new Date().toLocaleDateString('ja-JP');
  console.log(`レポート日: ${reportDate}\n`);

  try {
    // データ取得
    const customers = await getCustomers();
    const allTasks = await getAllTasks();

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const oneWeekAgo = new Date(today);
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);

    const oneWeekLater = new Date(today);
    oneWeekLater.setDate(oneWeekLater.getDate() + 7);

    // 顧客マップ作成
    const customerMap = {};
    for (const c of customers) {
      customerMap[c.record_id] = {
        name: c.fields['会社名'],
        openingDate: c.fields['開業予定日'],
        status: c.fields['ステータス']
      };
    }

    // 顧客別タスク集計
    const customerStats = {};
    for (const customerId of Object.keys(customerMap)) {
      customerStats[customerId] = {
        total: 0,
        completed: 0,
        overdue: 0,
        thisWeekCompleted: 0,
        nextWeekDue: []
      };
    }

    // タスク集計
    let totalTasks = 0;
    let completedTasks = 0;
    let overdueTasks = 0;
    const thisWeekCompleted = [];
    const nextWeekDue = [];

    for (const task of allTasks) {
      const customerId = task.fields['顧客']?.[0]?.record_ids?.[0];
      const status = task.fields['ステータス'];
      const dueDate = task.fields['期限'];
      const updatedTime = task.fields['更新時間'];

      if (!customerId || !customerStats[customerId]) continue;

      totalTasks++;
      customerStats[customerId].total++;

      if (status === '完了') {
        completedTasks++;
        customerStats[customerId].completed++;

        // 今週完了したタスク
        if (updatedTime && updatedTime >= oneWeekAgo.getTime()) {
          thisWeekCompleted.push({
            ...task,
            customerName: customerMap[customerId]?.name
          });
          customerStats[customerId].thisWeekCompleted++;
        }
      } else if (status !== '保留') {
        // 期限超過チェック
        if (dueDate && dueDate < today.getTime()) {
          overdueTasks++;
          customerStats[customerId].overdue++;
        }

        // 来週期限タスク
        if (dueDate && dueDate >= today.getTime() && dueDate <= oneWeekLater.getTime()) {
          nextWeekDue.push({
            ...task,
            customerName: customerMap[customerId]?.name
          });
          customerStats[customerId].nextWeekDue.push(task);
        }
      }
    }

    // レポート生成
    console.log('========================================');
    console.log('【全体サマリー】');
    console.log('========================================');
    console.log(`  アクティブ顧客数: ${customers.length}`);
    console.log(`  全タスク数: ${totalTasks}`);
    console.log(`  完了タスク数: ${completedTasks}`);
    console.log(`  全体進捗率: ${totalTasks > 0 ? ((completedTasks / totalTasks) * 100).toFixed(1) : 0}%`);
    console.log(`  期限超過タスク: ${overdueTasks}`);
    console.log(`  今週完了: ${thisWeekCompleted.length}`);
    console.log(`  来週期限: ${nextWeekDue.length}`);

    console.log('\n========================================');
    console.log('【顧客別進捗】（進捗率低い順）');
    console.log('========================================');

    const sortedCustomers = Object.entries(customerStats)
      .map(([id, stats]) => ({
        id,
        name: customerMap[id]?.name,
        ...stats,
        progressRate: stats.total > 0 ? (stats.completed / stats.total) * 100 : 0
      }))
      .filter(c => c.total > 0)
      .sort((a, b) => a.progressRate - b.progressRate);

    for (const c of sortedCustomers) {
      const bar = generateProgressBar(c.progressRate);
      console.log(`\n  ${c.name}`);
      console.log(`    ${bar} ${c.progressRate.toFixed(1)}%`);
      console.log(`    完了: ${c.completed}/${c.total} | 超過: ${c.overdue} | 来週期限: ${c.nextWeekDue.length}`);
    }

    console.log('\n========================================');
    console.log('【来週期限タスク】');
    console.log('========================================');
    for (const task of nextWeekDue.sort((a, b) => (a.fields['期限'] || 0) - (b.fields['期限'] || 0))) {
      const dueDate = new Date(task.fields['期限']).toLocaleDateString('ja-JP');
      console.log(`  📅 ${dueDate} | ${task.customerName}`);
      console.log(`     [${task.fields['WBS番号']}] ${task.fields['タスク名']}`);
    }

    // 通知送信
    if (SEND_NOTIFICATION) {
      console.log('\n通知を送信中...');
      await sendWeeklyReport(reportDate, {
        totalCustomers: customers.length,
        totalTasks,
        completedTasks,
        overdueTasks,
        thisWeekCompletedCount: thisWeekCompleted.length,
        nextWeekDueCount: nextWeekDue.length,
        customerStats: sortedCustomers,
        nextWeekDue
      });
      console.log('✓ 送信完了');
    }

    console.log('\n========================================');
    console.log('レポート生成完了');
    console.log('========================================');

  } catch (error) {
    console.error('\nエラー:', error.message);
    process.exit(1);
  }
}

function generateProgressBar(percentage) {
  const filled = Math.round(percentage / 10);
  const empty = 10 - filled;
  return '█'.repeat(filled) + '░'.repeat(empty);
}

async function getCustomers() {
  const res = await client.bitable.appTableRecord.list({
    path: { app_token: APP_TOKEN, table_id: CUSTOMER_TABLE_ID },
    params: { page_size: 100 }
  });
  if (res.code !== 0) throw new Error(`顧客取得エラー: ${res.msg}`);
  return res.data.items.filter(c =>
    c.fields['ステータス'] === '進行中' || c.fields['ステータス'] === '準備中'
  );
}

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
    if (res.code !== 0) throw new Error(`タスク取得エラー: ${res.msg}`);
    allTasks = allTasks.concat(res.data.items || []);
    pageToken = res.data.page_token;
  } while (pageToken);
  return allTasks;
}

async function sendWeeklyReport(reportDate, data) {
  if (!WEBHOOK_URL) {
    console.log('  ⚠ LARK_WEBHOOK_URLが設定されていません');
    return;
  }

  const progressRate = data.totalTasks > 0
    ? ((data.completedTasks / data.totalTasks) * 100).toFixed(1)
    : 0;

  let content = `📊 **週次進捗レポート** (${reportDate})\n\n`;
  content += `**■ 全体サマリー**\n`;
  content += `• アクティブ顧客: ${data.totalCustomers}社\n`;
  content += `• 全体進捗率: ${progressRate}% (${data.completedTasks}/${data.totalTasks})\n`;
  content += `• 期限超過: ${data.overdueTasks}件\n`;
  content += `• 今週完了: ${data.thisWeekCompletedCount}件\n`;
  content += `• 来週期限: ${data.nextWeekDueCount}件\n\n`;

  content += `**■ 顧客別進捗（低い順）**\n`;
  for (const c of data.customerStats.slice(0, 5)) {
    const bar = generateProgressBar(c.progressRate);
    content += `${c.name}: ${bar} ${c.progressRate.toFixed(0)}%\n`;
  }

  if (data.nextWeekDue.length > 0) {
    content += `\n**■ 来週期限タスク（抜粋）**\n`;
    for (const task of data.nextWeekDue.slice(0, 5)) {
      const dueDate = new Date(task.fields['期限']).toLocaleDateString('ja-JP');
      content += `• ${dueDate} | ${task.customerName} | ${task.fields['タスク名']}\n`;
    }
  }

  content += `\n[📋 LarkBaseで詳細確認](https://www.larksuite.com/base/${APP_TOKEN})`;

  await fetch(WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      msg_type: 'text',
      content: { text: content }
    })
  });
}

main();
