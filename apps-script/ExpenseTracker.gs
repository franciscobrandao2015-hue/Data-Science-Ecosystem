// ============================================================
// ENTRY POINT
// ============================================================

function doGet(e) {
  var p = e.parameter;
  if (p.type) return handleShortcutEntry(e);
  if (p.quick) return showQuickEntry();

  return showApp();
}


// ============================================================
// SHORTCUT ENTRY LOGGING (Expenses / Income)
// ============================================================

function toNumber(text) {
  if (!text) return 0;
  var cleaned = String(text).replace(",", ".");
  var n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function sortByDate(sheet, numCols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  sheet.getRange(2, 1, lastRow - 1, numCols).sort({ column: 1, ascending: true });
}

function handleShortcutEntry(e) {
  var params = e.parameter;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entryDate = data.date ? new Date(data.date + "T12:00:00") : new Date();

  if (params.type === "income") {
    var incomeSheet = ss.getSheetByName("Income");
    if (!incomeSheet) throw new Error("Sheet 'Income' not found");
    var incomeDate = params.date ? new Date(params.date) : now;
    incomeSheet.appendRow([
      entryDate,
      params.source || "",
      params.description || "",
      toNumber(params.amount),
      params.note || ""
    ]);
    sortByDate(incomeSheet, 5);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "ok",
        logged: "income"
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = ss.getSheetByName("Transactions");
  if (!sheet) throw new Error("Sheet 'Transactions' not found");

  var expenseDate = params.date ? new Date(params.date) : now;

  sheet.appendRow([
    entryDate,
    params.description || "",
    params.category || "",
    params.account || "",
    toNumber(params.amount),
    params.note || ""
  ]);

  sortByDate(sheet, 6);

  var rows = sheet.getDataRange().getValues();
  var total = 0;
  var m = now.getMonth();
  var y = now.getFullYear();

  for (var i = 1; i < rows.length; i++) {
    var d = new Date(rows[i][0]);

    if (d.getMonth() == m && d.getFullYear() == y) {
      total += Number(rows[i][4]) || 0;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "ok",
      month_total: total
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ANDROID SHORTCUT
// ============================================================

function saveQuickEntry(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();

  if (data.type === "income") {
    var incomeSheet = ss.getSheetByName("Income");
    if (!incomeSheet) throw new Error("Sheet 'Income' not found");

    incomeSheet.appendRow([
      now,
      data.source || "",
      data.description || "",
      toNumber(data.amount),
      ""
    ]);

    sortByDate(incomeSheet, 5);

    return {
      status: "ok",
      logged: "income"
    };
  }

  var sheet = ss.getSheetByName("Transactions");
  if (!sheet) throw new Error("Sheet 'Transactions' not found");

  sheet.appendRow([
    now,
    data.description || "",
    data.category || "",
    data.account || "",
    toNumber(data.amount),
    ""
  ]);

  sortByDate(sheet, 6);

  return {
    status: "ok",
    logged: "expense"
  };
}

function showQuickEntry() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setup = getSetupData(ss);
  var template = HtmlService.createTemplate(QUICK_ENTRY_HTML);
  template.categories = JSON.stringify(setup.categories);
  template.accounts = JSON.stringify(setup.accounts);
  return template.evaluate()
    .setTitle("Quick Entry")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ============================================================
// SHARED HELPERS
// ============================================================

function safe(fn, fallback) {
  try {
    return fn();
  } catch (err) {
    return fallback;
  }
}

var MONTH_ABBR = [
  "Jan","Feb","Mar","Apr","May","Jun",
  "Jul","Aug","Sep","Oct","Nov","Dec"
];

// ============================================================
// APP SHELL
// ============================================================

function showApp() {

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var data =
    getInitialAppData(ss);

  var template =
    HtmlService.createTemplate(APP_HTML);

  template.data =
    JSON.stringify(data);

  return template
    .evaluate()
    .setTitle("Expense Tracker")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1"
    );

}



function getInitialAppData(ss) {

  var now = new Date();

  return {

    dashboard:
      getDashboardData(
        ss,
        now.getMonth() + 1,
        now.getFullYear()
      ),

    monthlybudget:
      getMonthlyBudgetData(ss),

    income:
      getIncomeData(ss),

    bills:
      getBillsData(ss),

    debt:
      getDebtData(ss),

    goals:
      getGoalsData(ss),

    networth:
      getNetWorthData(ss),

    setup:
      getSetupData(ss)

  };

}

function getPageData(page, params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  switch (page) {
    case "dashboard":
      return getDashboardData(ss, params.month, params.year);

    case "monthlybudget":
      return getMonthlyBudgetData(ss);

    case "transactions":
      return getTransactionsData(ss);

    case "income":
      return getIncomeData(ss);

    case "bills":
      return getBillsData(ss);

    case "debt":
      return getDebtData(ss);

    case "goals":
      return getGoalsData(ss);

    case "networth":
      return getNetWorthData(ss);

    case "setup":
      return getSetupData(ss);
  }

  return {};
}

// ============================================================
// DASHBOARD
// ============================================================

function getDashboardData(ss, month1, year) {

  var selectedMonth =
    parseInt(month1) - 1;

  year =
    parseInt(year);


  // ============================================================
  // MONTHLY DATA
  // ============================================================

  var months = [];

  for (var m = 0; m < 12; m++) {

    months.push({

      income: 0,
      expenses: 0,
      categories: {},
      transactions: [],
      incomeRows: []

    });

  }


  // ============================================================
  // TRANSACTIONS
  // ============================================================

  safe(function () {

    var tx =
      ss
        .getSheetByName("Transactions")
        .getDataRange()
        .getValues();


    for (var i = 1; i < tx.length; i++) {

      var d =
        new Date(tx[i][0]);

      var cat =
        tx[i][2];

      var amt =
        Number(tx[i][4]) || 0;


      if (!(d instanceof Date) || isNaN(d))
        continue;


      if (d.getFullYear() !== year)
        continue;


      var m =
        d.getMonth();


      months[m].expenses += amt;


      months[m].categories[cat] =
        (months[m].categories[cat] || 0) + amt;


      months[m].transactions.push({

        date: Utilities.formatDate(
          d,
          Session.getScriptTimeZone(),
          "dd/MM"
        ),

        description: tx[i][1],

        category: cat,

        amount: amt

      });

    }

  }, null);


  // ============================================================
  // INCOME
  // ============================================================

  safe(function () {

    var inc =
      ss
        .getSheetByName("Income")
        .getDataRange()
        .getValues();


    for (var i = 1; i < inc.length; i++) {

      var d =
        new Date(inc[i][0]);


      if (!(d instanceof Date) || isNaN(d))
        continue;


      if (d.getFullYear() !== year)
        continue;


      var m =
        d.getMonth();


      var amt =
        Number(inc[i][3]) || 0;


      months[m].income += amt;


      months[m].incomeRows.push({

        date: Utilities.formatDate(
          d,
          Session.getScriptTimeZone(),
          "dd/MM"
        ),

        source: inc[i][1],

        description: inc[i][2],

        amount: amt

      });

    }

  }, null);


  // ============================================================
  // MONTHLY ARRAYS
  // ============================================================

  var monthly =
    months.map(function (m) {
      return m.expenses;
    });


  var incMonthly =
    months.map(function (m) {
      return m.income;
    });


  // ============================================================
  // PROJECTED INCOME
  // ============================================================

  var projectedIncome =
    safe(function () {

      var sheet =
        ss.getSheetByName("Income");

      var vals =
        sheet
          .getRange(2, 8, 12, 2)
          .getValues();

      var byMonth =
        new Array(12).fill(0);


      vals.forEach(function (row, i) {

        byMonth[i] =
          Number(row[1]) || 0;

      });


      return byMonth;

    }, new Array(12).fill(0));


  // ============================================================
  // NEXT BILL
  // ============================================================

  var nextBill =
    safe(function () {

      var bills =
        ss
          .getSheetByName("Bill Calendar")
          .getDataRange()
          .getValues();


      var best = null;


      for (var i = 4; i < bills.length; i++) {

        if (
          bills[i][5] === "Unpaid" &&
          bills[i][2]
        ) {

          if (
            !best ||
            bills[i][2] < best.dueDay
          ) {

            best = {

              name: bills[i][0],

              amount: bills[i][4],

              dueDay: bills[i][2]

            };

          }

        }

      }


      return best;

    }, null);


  // ============================================================
  // DEBT SERIES
  // ============================================================

  var debtSeries =
    safe(function () {

      return ss
        .getSheetByName("Debt Tracker")
        .getRange(31, 11, 1, 12)
        .getValues()[0]
        .map(function (v) {

          return Number(v) || 0;

        });

    }, new Array(12).fill(0));


  // ============================================================
  // SAVINGS GOALS
  // ============================================================

  var goals =
    safe(function () {

      var rows =
        ss
          .getSheetByName("Savings Goals")
          .getRange(5, 1, 15, 6)
          .getValues();


      var list = [];


      rows.forEach(function (r) {

        if (!r[0])
          return;


        list.push({

          name: r[0],

          target: Number(r[1]) || 0,

          current: Number(r[2]) || 0,

          progress: Number(r[4]) || 0,

          monthsLeft: r[5]

        });

      });


      return list;

    }, []);


  // ============================================================
  // NET WORTH — ALL 12 MONTHS
  // ============================================================

  var netWorthSeries =
    safe(function () {

      var rows =
        ss
          .getSheetByName("Net Worth")
          .getRange(5, 1, 12, 4)
          .getValues();


      return rows.map(function (nw, i) {

        var cash =
          Number(nw[1]) || 0;


        var assets =
          cash +
          (Number(nw[2]) || 0) +
          (Number(nw[3]) || 0);


        var liabilities =
          debtSeries[i] || 0;


        return {

          cash: cash,

          assets: assets,

          liabilities: liabilities,

          net: assets - liabilities

        };

      });

    }, new Array(12).fill(null));


  // ============================================================
  // BUILD DASHBOARD DATA FOR EACH MONTH
  // ============================================================

  var dashboardMonths =
    months.map(function (m, i) {

      var pieCategories = {};

      var allCategories = [];


      for (var k in m.categories) {

        allCategories.push({

          name: k,

          amount: m.categories[k]

        });


        if (m.categories[k] > 0) {

          pieCategories[k] =
            m.categories[k];

        }

      }


      allCategories.sort(function (a, b) {

        return b.amount - a.amount;

      });


      return {

        selectedMonth: i + 1,

        selectedYear: year,

        income: m.income,

        expenses: m.expenses,

        net: m.income - m.expenses,

        nextBill: nextBill,

        pieCategories: pieCategories,

        allCategories: allCategories,

        monthly: monthly,

        debtSeries: debtSeries,

        goals: goals,

        netWorth:
          netWorthSeries[i],

        transactions: m.transactions,

        incomeRows: m.incomeRows,

        projectedIncome: projectedIncome,

        incomeMonthly: incMonthly

      };

    });


  // ============================================================
  // CURRENT / SELECTED MONTH
  // ============================================================

  var current =
    dashboardMonths[selectedMonth] ||
    dashboardMonths[0];


  return {

    selectedMonth:
      current.selectedMonth,

    selectedYear:
      year,

    income:
      current.income,

    expenses:
      current.expenses,

    net:
      current.net,

    nextBill:
      current.nextBill,

    pieCategories:
      current.pieCategories,

    allCategories:
      current.allCategories,

    monthly:
      monthly,

    debtSeries:
      debtSeries,

    goals:
      goals,

    netWorth:
      current.netWorth,

    transactions:
      current.transactions,

    incomeRows:
      current.incomeRows,

    projectedIncome:
      projectedIncome,

    incomeMonthly:
      incMonthly,

    dashboardMonths:
      dashboardMonths

  };

}

// ============================================================
// MONTHLY BUDGET PAGE
// ============================================================

function getMonthlyBudgetData(ss) {

  return safe(function () {

    var rows =
      ss
        .getSheetByName("Monthly Budget")
        .getRange(1, 1, 18, 27)
        .getValues();

    var months = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December"
    ];

    var data = {};

    months.forEach(function(month, m){

      var budgetCol = 1 + (m * 2);
      var actualCol = budgetCol + 1;

      data[month] = {
        budget: Number(rows[17][budgetCol]) || 0,
        actual: Number(rows[17][actualCol]) || 0,
        categories: []
      };

      for (var i = 2; i <= 16; i++) {

        var budget =
          Number(rows[i][budgetCol]) || 0;

        var actual =
          Number(rows[i][actualCol]) || 0;

        data[month].categories.push({

          category: rows[i][0],

          budget: budget,

          actual: actual,

          difference: actual - budget

        });

      }

    });

    return {
      months: months,
      data: data
    };

  }, {
    months: [],
    data: {}
  });
}
// ============================================================
// TRANSACTIONS PAGE
// ============================================================

function getTransactionsData(ss) {

  return safe(function () {

    var rows = ss
      .getSheetByName("Transactions")
      .getDataRange()
      .getValues();

    var list = [];
    var categoryTotals = {};
    var monthMap = {};

    for (var i = 1; i < rows.length; i++) {

      var d = new Date(rows[i][0]);

      if (!(d instanceof Date) || isNaN(d)) continue;

      var amt = Number(rows[i][4]) || 0;
      var cat = rows[i][2];

      list.push({
        date: Utilities.formatDate(
          d,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
        ),
        description: rows[i][1],
        category: cat,
        account: rows[i][3],
        amount: amt,
        note: rows[i][5]
      });

      if (amt > 0) {
        categoryTotals[cat] =
          (categoryTotals[cat] || 0) + amt;
      }

      var key =
        d.getFullYear() +
        "-" +
        ("0" + (d.getMonth() + 1)).slice(-2);

      monthMap[key] =
        (monthMap[key] || 0) + amt;
    }

    list.reverse();

    var monthKeys =
      Object.keys(monthMap)
        .sort()
        .slice(-12);

    var year = safe(function () { return ss.getSheetByName("Setup").getRange("B6").getValue(); }, "");
    return {
      rows: list, categoryTotals: categoryTotals, year: year,
      monthly: monthKeys.map(function (k) { return { label: MONTH_ABBR[parseInt(k.split("-")[1], 10) - 1], amount: monthMap[k] }; })
    };


  }, {
    rows: [],
    categoryTotals: {},
    monthly: []
  });
}

// ============================================================
// INCOME PAGE
// ============================================================

function getIncomeData(ss) {

  return safe(function () {

    var rows = ss
      .getSheetByName("Income")
      .getDataRange()
      .getValues();

    var list = [];
    var total = 0;
    var monthMap = {};

    for (var i = 1; i < rows.length; i++) {

      var d = new Date(rows[i][0]);

      if (!(d instanceof Date) || isNaN(d)) continue;

      var amt = Number(rows[i][3]) || 0;

      list.push({
        date: Utilities.formatDate(
          d,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
        ),
        source: rows[i][1],
        description: rows[i][2],
        amount: amt,
        note: rows[i][4]
      });

      total += amt;

      var key =
        d.getFullYear() +
        "-" +
        ("0" + (d.getMonth() + 1)).slice(-2);

      monthMap[key] =
        (monthMap[key] || 0) + amt;
    }

    list.reverse();

    var monthKeys =
      Object.keys(monthMap)
        .sort()
        .slice(-12);

    return {
      rows: list,
      total: total,
      monthly: monthKeys.map(function (k) {

        return {
          label:
            MONTH_ABBR[
              parseInt(
                k.split("-")[1],
                10
              ) - 1
            ],
          amount: monthMap[k]
        };

      })
    };

  }, {
    rows: [],
    total: 0,
    monthly: []
  });
}

// ============================================================
// BILL CALENDAR PAGE
// ============================================================

function getBillsData(ss) {

  return safe(function () {

    var rows = ss
      .getSheetByName("Bill Calendar")
      .getRange(5, 1, 60, 7)
      .getValues();

    var list = [];
    var paid = 0;
    var unpaid = 0;
    var totalDue = 0;

    rows.forEach(function (r) {

      if (!r[0]) return;

      list.push({
        name: r[0],
        category: r[1],
        dueDay: r[2],
        frequency: r[3],
        amount: r[4],
        status: r[5]
      });

      if (r[5] === "Paid") {
        paid++;
      } else if (r[5] === "Unpaid") {
        unpaid++;
        totalDue += Number(r[4]) || 0;
      }
    });

    return {
      rows: list,
      paid: paid,
      unpaid: unpaid,
      totalDue: totalDue
    };

  }, {
    rows: [],
    paid: 0,
    unpaid: 0,
    totalDue: 0
  });
}

// ============================================================
// DEBT TRACKER PAGE
// ============================================================

function getDebtData(ss) {

  return safe(function () {

    var sheet =
      ss.getSheetByName("Debt Tracker");

    var rows =
      sheet.getRange(5, 1, 25, 9)
        .getValues();

    var monthlyRows =
      sheet.getRange(5, 11, 25, 12)
        .getValues();

    var list = [];
    var schedule = [];

    rows.forEach(function (r, i) {

      if (!r[0]) return;

      list.push({
        name: r[0],
        type: r[1],
        balance: Number(r[2]) || 0,
        rate: Number(r[3]) || 0,
        minPayment: Number(r[4]) || 0,
        months: r[6],
        totalInterest: Number(r[7]) || 0
      });

      schedule.push({
        name: r[0],
        values: monthlyRows[i].map(function (v) {
          return Number(v) || 0;
        })
      });
    });

    var totalSeries =
      sheet
        .getRange(31, 11, 1, 12)
        .getValues()[0]
        .map(function (v) {
          return Number(v) || 0;
        });

    return {
      rows: list,
      totalSeries: totalSeries,
      schedule: schedule
    };

  }, {
    rows: [],
    totalSeries: new Array(12).fill(0),
    schedule: []
  });
}

// ============================================================
// SAVINGS GOALS PAGE
// ============================================================

function getGoalsData(ss) {

  return safe(function () {

    var rows =
      ss
        .getSheetByName("Savings Goals")
        .getRange(5, 1, 15, 7)
        .getValues();

    var list = [];

    rows.forEach(function (r) {

      if (!r[0]) return;

      list.push({
        name: r[0],
        target: Number(r[1]) || 0,
        current: Number(r[2]) || 0,
        progress: Number(r[4]) || 0,
        monthsLeft: r[5],
        neededPerMonth: Number(r[6]) || 0
      });
    });

    return {
      rows: list
    };

  }, {
    rows: []
  });
}

// ============================================================
// NET WORTH PAGE
// ============================================================

function getNetWorthData(ss) {

  return safe(function () {

    var nwRows =
      ss
        .getSheetByName("Net Worth")
        .getRange(5, 1, 12, 4)
        .getValues();

    var debtSeries =
      ss
        .getSheetByName("Debt Tracker")
        .getRange(31, 11, 1, 12)
        .getValues()[0]
        .map(function (v) {
          return Number(v) || 0;
        });

    var list =
      nwRows.map(function (r, i) {

        var cash = Number(r[1]) || 0;
        var inv = Number(r[2]) || 0;
        var prop = Number(r[3]) || 0;

        var assets =
          cash + inv + prop;

        var liab =
          debtSeries[i] || 0;

        return {
          month: r[0],
          cash: cash,
          investments: inv,
          property: prop,
          assets: assets,
          liabilities: liab,
          net: assets - liab
        };
      });

    return {
      rows: list
    };

  }, {
    rows: []
  });
}

// ============================================================
// SETUP PAGE (read-only)
// ============================================================

function getSetupData(ss) {

  return safe(function () {

    var sheet =
      ss.getSheetByName("Setup");

    var currency =
      sheet.getRange("B5").getValue();

    var year =
      sheet.getRange("B6").getValue();

    var startingCashBalance =
      sheet.getRange("B14").getValue() ;

    var categories =
      sheet
        .getRange(6, 4, 15, 1)
        .getValues()
        .flat()
        .filter(String);

    var accounts =
      sheet
        .getRange(6, 7, 10, 1)
        .getValues()
        .flat()
        .filter(String);

    return {
      currency: currency,
      year: year,
      categories: categories,
      accounts: accounts,
      startingCashBalance: startingCashBalance
    };

  }, {
    currency: "",
    year: "",
    categories: [],
    accounts: []
  });
}

// ============================================================
// HTML APP
// ============================================================

var APP_HTML = `<!DOCTYPE html>
<html>
<head>
<base target="_top">

<meta name="viewport" content="width=device-width, initial-scale=1">

<link
  href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap"
  rel="stylesheet"
>

<script src="https://www.gstatic.com/charts/loader.js"></script>

<style>

:root{
  --bg:#0D0F13;
  --surface:#15171D;
  --surface-alt:#1A1D24;
  --border:#23262E;
  --text:#EDEEF2;
  --text-muted:#868C99;
  --accent:#6C8EEF;
  --accent-soft:#161B2C;
  --positive:#5FB88C;
  --negative:#D97A82;
}

*{
  box-sizing:border-box;
  -webkit-text-size-adjust:100%;
  text-size-adjust:100%;
}

body{
  font-family:'Inter',-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
  font-variant-numeric:tabular-nums;
  background:var(--bg);
  color:var(--text);
  margin:0;
  display:flex;
  min-height:100vh;
  align-items:flex-start;
}

.sidebar{
  width:190px;
  background:var(--surface);
  border-right:1px solid var(--border);
  padding:16px 0;
  flex-shrink:0;
  position:sticky;
  top:0;
  height:100vh;
  overflow-y:auto;
  transition:width .2s;
}

.sidebar.collapsed{
  width:56px;
}

.sidebar.collapsed .nav-label,
.sidebar.collapsed .brand-text{
  display:none;
}

.collapse-btn{
  background:none;
  border:none;
  cursor:pointer;
  font-size:16px;
  padding:6px 16px;
  color:var(--text-muted);
  width:100%;
  text-align:left;
}

.brand{
  font-weight:600;
  font-size:14px;
  letter-spacing:.02em;
  color:var(--text);
  padding:0 16px 16px;
}

.nav-item{
  display:flex;
  align-items:center;
  gap:11px;
  width:100%;
  text-align:left;
  background:none;
  border:none;
  padding:9px 16px;
  font-size:13px;
  font-weight:500;
  cursor:pointer;
  color:var(--text-muted);
  border-left:2px solid transparent;
}

.nav-item:hover{
  background:var(--surface-alt);
  color:var(--text);
}

.nav-item.active{
  color:var(--accent);
  border-left-color:var(--accent);
  background:var(--accent-soft);
  font-weight:600;
}

.nav-icon{
  display:inline-flex;
  flex-shrink:0;
}

.nav-icon svg{
  width:17px;
  height:17px;
  display:block;
}

.sidebar.collapsed .nav-item{
  justify-content:center;
}

.content{
  flex:1;
  padding:32px;
  max-width:1100px;
  position:relative;
}

.page{
  display:none;
}

.page.active{
  display:block;
}

h1{
  font-size:20px;
  margin:0 0 4px;
  font-weight:600;
  letter-spacing:-0.01em;
  color:var(--text);
}

.sub{
  color:var(--text-muted);
  font-size:13px;
  margin:0 0 18px;
}

select{
  font-size:14px;
  padding:6px 10px;
  margin-bottom:20px;
  border:1px solid var(--border);
  border-radius:8px;
  background:var(--surface);
  color:var(--text);
}

.cards{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
  gap:16px;
  margin-bottom:32px;
}

.card{
  background:var(--surface);
  border-radius:12px;
  padding:16px 18px;
  border:1px solid var(--border);
}

.card .lbl{
  font-size:11px;
  color:var(--text-muted);
  font-weight:600;
  letter-spacing:.04em;
  text-transform:uppercase;
}

.card .val{
  font-size:21px;
  font-weight:600;
  letter-spacing:-0.01em;
  margin-top:8px;
  color:var(--text);
}

.row{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:24px;
  margin-bottom:32px;
}

.panel{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:12px;
  padding:18px;
}

.panel p{
  margin:0 0 8px;
  font-size:12px;
  line-height:1.5;
  color:var(--text);
}

.bar-bg{
  height:8px;
  background:var(--surface-alt);
  border-radius:4px;
  overflow:hidden;
  margin-top:10px;
}

.bar-fill{
  height:100%;
  background:var(--accent);
  border-radius:4px;
}

h2{
  font-size:13px;
  letter-spacing:.01em;
  border-bottom:1px solid var(--border);
  padding-bottom:10px;
  margin:0 0 14px;
  font-weight:600;
  color:var(--text);
}

table{
  width:100%;
  border-collapse:collapse;
  font-size:13px;
}

td,
th{
  padding:6px 8px;
  text-align:left;
  color:var(--text);
}

th{
  color:var(--text-muted);
  font-size:11px;
  font-weight:600;
}

tr:nth-child(even){
  background:var(--surface-alt);
}

.tables-wrap{
  max-height:400px;
  overflow-y:auto;
}

.empty{
  color:var(--text-muted);
  font-size:13px;
  font-style:italic;
  padding:20px 0;
  text-align:center;
}

.cat-list{
  font-size:12px;
  margin-top:10px;
}

.cat-list div{
  display:flex;
  justify-content:space-between;
  padding:3px 0;
}

.category-table-wrap{
  width:100%;
  margin-top:8px;
  overflow-x:auto;
}

.category-table-wrap table{
  width:100%;
  min-width:0;
}

.category-table-wrap th{
  padding:6px 8px;
  border-bottom:1px solid var(--border);
}

.category-table-wrap td{
  padding:7px 8px;
}

.category-table-wrap td:last-child,
.category-table-wrap th:last-child{
  text-align:right;
}

.category-table-wrap tr:nth-child(even){
  background:var(--surface-alt);
}

.tables-wrap{
  width:100%;
  max-width:100%;
  overflow-x:auto;
  overflow-y:auto;
  -webkit-overflow-scrolling:touch;
}

#de-schedule{
  width:max-content;
  min-width:100%;
  border-collapse:collapse;
}

#de-schedule th,
#de-schedule td{
  white-space:normal;
}

.neg{
  color:var(--negative);
}

.chip{
  display:inline-block;
  background:var(--accent-soft);
  color:var(--accent);
  padding:4px 10px;
  border-radius:12px;
  font-size:12px;
  margin:0 6px 6px 0;
}

.loading{
  position:absolute;
  top:0;
  left:0;
  right:0;
  bottom:0;
  background:rgba(13,15,19,0.85);
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:13px;
  color:var(--accent);
  z-index:10;
}

.dashboard-trend-panel{
  display:flex;
  flex-direction:column;
}

.dashboard-trend-panel #d-trend{
  flex:1;
  min-height:360px;
}

/* ============================================================
   MOBILE ONLY
   Desktop permanece inalterado.
   ============================================================ */

.mobile-app-title {
  display:none;
}

@media (max-width:800px) {

  /* ----------------------------------------------------------
     BASE
     ---------------------------------------------------------- */

  html,
  body {
    width:100%;
    min-width:0;
    max-width:100%;
    margin:0;
    padding:0;
    overflow-x:hidden;
  }



  body {
    display:flex;
    align-items:flex-start;
  }


  /* ----------------------------------------------------------
     NAV MOBILE

     Sidebar lateral, inicialmente colapsada.
     O utilizador pode expandi-la através do botão.
     ---------------------------------------------------------- */

  .sidebar {
    width:56px;
    min-width:56px;
    max-width:56px;

    height:100vh;

    display:flex;
    flex-direction:column;
    align-items:stretch;

    position:fixed;
    top:0;
    left:0;
    bottom:0;
    z-index:1000;

    overflow-x:hidden;
    overflow-y:auto;

    padding:16px 0;

    background:var(--surface);

    border-right:1px solid var(--border);
    border-bottom:none;

    white-space:normal;

    transition:width .2s ease,
               min-width .2s ease,
               max-width .2s ease;

    scrollbar-width:none;
  }

  .sidebar::-webkit-scrollbar {
    display:none;
  }

  /* Quando o utilizador expande a barra */

  .sidebar:not(.collapsed) {
    width:190px;
    min-width:190px;
    max-width:190px;
  }

  /* ----------------------------------------------------------
     MARCA
     ---------------------------------------------------------- */

  .brand {
    display:none;
    padding:0 16px 16px;
  }

  .sidebar:not(.collapsed) .brand {
    display:block;
  }

  /* ----------------------------------------------------------
     BOTÃO DE COLAPSAR
     ---------------------------------------------------------- */

  .collapse-btn {
    display:block;

    width:100%;

    background:none;
    border:none;

    cursor:pointer;

    font-size:16px;

    padding:6px 16px;

    color:var(--text-muted);

    text-align:left;

    flex-shrink:0;
  }

  .sidebar.collapsed .collapse-btn {
    padding-left:16px;
    padding-right:16px;
    text-align:center;
  }

  /* ----------------------------------------------------------
     NAV ITEMS
     ---------------------------------------------------------- */

  .nav-item {
    display:flex;
    align-items:center;
    gap:11px;

    flex:0 0 auto;

    width:100%;
    min-width:0;

    text-align:left;

    background:none;
    border:none;

    border-left:2px solid transparent;
    border-bottom:none;

    padding:9px 16px;

    font-size:13px;
    font-weight:500;
    line-height:1.2;

    white-space:nowrap;

    cursor:pointer;

    color:var(--text-muted);
  }

  .nav-item:hover {
    background:var(--surface-alt);
  }

  .nav-item.active {
    color:var(--accent);

    border-left-color:var(--accent);
    border-bottom:none;

    background:var(--accent-soft);

    font-weight:600;
  }

  /* Na barra colapsada mostramos apenas os ícones */

  .sidebar.collapsed .nav-item {
    padding-left:0;
    padding-right:0;
    justify-content:center;
  }

  .sidebar.collapsed .nav-label {
    display:none;
  }


  /* ----------------------------------------------------------
     TÍTULO MOBILE
     ---------------------------------------------------------- */

  .mobile-app-title {
    display:block;

    font-size:15px;
    font-weight:600;

    color:var(--text);

    margin:0 0 8px;
  }


  /* ----------------------------------------------------------
     CONTEÚDO
     ---------------------------------------------------------- */

  .content {
    flex:1;

    width:auto;
    max-width:none;
    min-width:0;

    margin-left:56px;

    padding:16px 12px;

    overflow:hidden;
  }

  .sidebar:not(.collapsed) ~ .content {
    margin-left:190px;
  }
  .page {
    width:100%;
    max-width:100%;
    min-width:0;
  }


  /* ----------------------------------------------------------
     TÍTULOS
     ---------------------------------------------------------- */

  h1 {
    font-size:18px;
    line-height:1.25;
    margin-bottom:4px;
  }

  h2 {
    font-size:13px;
  }

  .sub {
    font-size:12px;
  }


  /* ----------------------------------------------------------
     SELECT
     ---------------------------------------------------------- */

  select {
    width:100%;
    max-width:100%;

    margin-bottom:16px;

    font-size:13px;
    padding:8px 10px;
  }


  /* ----------------------------------------------------------
     CARDS
     2 COLUNAS = 2x2 NO DASHBOARD
     ---------------------------------------------------------- */

  .cards {
    width:100%;
    max-width:100%;
    min-width:0;

    grid-template-columns:repeat(2,minmax(0,1fr));

    gap:10px;

    margin-bottom:18px;
  }

  .card {
    width:100%;
    min-width:0;

    padding:12px;

    border-radius:10px;
  }

  .card .lbl {
    font-size:9px;
    line-height:1.2;
  }

  .card .val {
    min-width:0;

    font-size:17px;
    line-height:1.25;

    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
  }


  /* ----------------------------------------------------------
     ROWS

     Um painel/gráfico por linha.
     ---------------------------------------------------------- */

  .row {
    width:100%;
    max-width:100%;
    min-width:0;

    grid-template-columns:minmax(0,1fr);

    gap:12px;

    margin-bottom:16px;
  }


  /* ----------------------------------------------------------
     PAINÉIS
     ---------------------------------------------------------- */

  .panel {
    width:100%;
    max-width:100%;
    min-width:0;

    padding:13px;

    border-radius:10px;

    overflow:hidden;
  }

  /* ----------------------------------------------------------
   GRÁFICOS MOBILE
   Mantêm uma largura confortável.
   Se não couberem, o painel permite scroll horizontal.
   ---------------------------------------------------------- */

  .chart-panel {
    width:100%;
    max-width:100%;
    min-width:0;

    overflow-x:auto;
    overflow-y:hidden;

    -webkit-overflow-scrolling:touch;
  }

  /* O gráfico tem espaço suficiente para não ficar achatado */
  .chart-panel > div[id] {
    width:520px !important;
    min-width:520px !important;
    max-width:none !important;
    height: 300px !important;
    min-height: 300px !important;
  }

  /* Altura confortável dos gráficos */
  #d-donut{width:100% !important;
    min-width:0 !important;
    max-width:100% !important;
    height:250px !important;
    min-height:250px !important;
  }

 #d-debt,

  #de-chart,
  #n-chart {
  height:300px !important;
  min-height:300px !important;
  }

  #t-donut { height:260px !important; min-height:260px !important; }

  .chart-panel:has(#d-donut){
    text-align:center;
  }
  /* ----------------------------------------------------------
     TABLE
     ---------------------------------------------------------- */

  .category-table-wrap{
    width:100%;
    max-width:100%;
    overflow-x:auto;
    -webkit-overflow-scrolling:touch;
  }

  .category-table-wrap table{
    width:100%;
    min-width:0;
  }

    .tables-wrap{
    width:100%;
    max-width:100%;
    overflow-x:auto;
    -webkit-overflow-scrolling:touch;
  }

  #de-schedule{
    width:max-content !important;
    min-width:100% !important;
    max-width:none !important;
  }

  #de-schedule td:not(:first-child){
  white-space:nowrap;
  }

  #de-schedule td:first-child{
    white-space:normal;
    overflow-wrap:break-word;
  }

  /* ----------------------------------------------------------
     ACCOUNTS
     ---------------------------------------------------------- */

  #d-accounts {
    grid-template-columns:repeat(2,minmax(0,1fr));
  }


  /* ----------------------------------------------------------
     CATEGORIAS
     ---------------------------------------------------------- */

  .cat-list {
    width:100%;
    max-width:100%;

    overflow:hidden;
  }

  .cat-list div {
    gap:8px;
  }

  .cat-list div span:first-child {
    min-width:0;

    overflow:hidden;

    text-overflow:ellipsis;

    white-space:nowrap;
  }


  /* ----------------------------------------------------------
     TABELAS
     ---------------------------------------------------------- */

  .tables-wrap {
    width:100%;
    max-width:100%;

    overflow-x:auto;
    overflow-y:auto;

    -webkit-overflow-scrolling:touch;
  }



  td,
  th {
    padding:7px 8px;
    white-space:nowrap;
  }


  /* ----------------------------------------------------------
     SAVINGS GOALS
     ---------------------------------------------------------- */

  #g-list {
    width:100%;
    max-width:100%;
  }

  #g-list .panel {
    overflow-x:hidden;
  }


  /* ----------------------------------------------------------
     SETUP
     ---------------------------------------------------------- */

  #s-categories,
  #s-accounts {
    width:100%;
    max-width:100%;
  }

  .chip {
    margin-right:4px;
  }


  /* ----------------------------------------------------------
     LOADING
     ---------------------------------------------------------- */

  .loading {
    position:fixed;
    inset:0;
    z-index:100;
  }
}

</style>
</head>

<body>

<nav class="sidebar" id="sidebar">

  <div class="brand">
    <span class="brand-text">Expense Tracker</span>
  </div>

  <button class="collapse-btn" id="collapseBtn">
    \u{2190}
  </button>

  <button class="nav-item active" data-page="dashboard">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 11.5 12 4l8 7.5"/><path d="M6 10v9a1 1 0 0 0 1 1h3v-6h4v6h3a1 1 0 0 0 1-1v-9"/></svg></span>
    <span class="nav-label">Dashboard</span>
  </button>

  <button class="nav-item" data-page="monthlybudget">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 20V11"/><path d="M12 20V4"/><path d="M19 20v-6"/></svg></span>
    <span class="nav-label">Monthly Budget</span>
  </button>

  <button class="nav-item" data-page="transactions">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M8.5 8h7M8.5 12h7M8.5 16h4"/></svg></span>
    <span class="nav-label">Transactions</span>
  </button>

  <button class="nav-item" data-page="income">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7a2 2 0 0 1 2-2h11a1 1 0 0 1 1 1v2"/><path d="M4 7v10a2 2 0 0 0 2 2h13a1 1 0 0 0 1-1v-7a1 1 0 0 0-1-1H6a2 2 0 0 1-2-2Z"/><circle cx="16.3" cy="13.5" r="1.1" fill="currentColor" stroke="none"/></svg></span>
    <span class="nav-label">Income</span>
  </button>

  <button class="nav-item" data-page="bills">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 9.5h16M8 3v3.5M16 3v3.5"/></svg></span>
    <span class="nav-label">Bill Calendar</span>
  </button>

  <button class="nav-item" data-page="debt">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="6" width="18" height="12" rx="2"/><path d="M3 10h18"/></svg></span>
    <span class="nav-label">Debt Tracker</span>
  </button>

  <button class="nav-item" data-page="goals">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="4.3"/><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none"/></svg></span>
    <span class="nav-label">Savings Goals</span>
  </button>

  <button class="nav-item" data-page="networth">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 16l5.5-5.5 4 4L20 7"/><path d="M14.5 7H20v5.5"/></svg></span>
    <span class="nav-label">Net Worth</span>
  </button>

  <button class="nav-item" data-page="setup">
    <span class="nav-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h9M17 7h3"/><circle cx="14" cy="7" r="2"/><path d="M4 12h3M11 12h9"/><circle cx="8" cy="12" r="2"/><path d="M4 17h9M17 17h3"/><circle cx="14" cy="17" r="2"/></svg></span>
    <span class="nav-label">Setup</span>
  </button>

</nav>

<main class="content">

<div id="loading" class="loading" style="display:none;">
  Loading…
</div>


<section id="page-dashboard" class="page active">

  <h1>Income &amp; expense summary</h1>

  <select id="monthSel"></select>

  <div class="cards">

    <div class="card">
      <div class="lbl">INCOME</div>
      <div class="val" id="d-income" style="color:var(--positive)"></div>
    </div>

    <div class="card">
      <div class="lbl">EXPENSES</div>
      <div class="val" id="d-exp" style="color:var(--negative)"></div>
    </div>

    <div class="card">
      <div class="lbl">NET</div>
      <div class="val" id="d-net"></div>
    </div>

    <div class="card">
      <div class="lbl">NEXT BILL DUE</div>
      <div class="val" id="d-bill" style="font-size:14px"></div>
    </div>

  </div>


  <div class="row">

    <div class="panel chart-panel">

      <div id="d-donut" style="height:250px;"></div>

      <div class="category-table-wrap">
        <table id="d-catlist"></table>
      </div>

    </div>

    <div class="panel dashboard-trend-panel">
      <div id="d-trend"></div>
    </div>

  </div>


  <h2>Accounts</h2>


  <div class="cards" id="d-accounts">

    <div class="card">
      <div class="lbl">CASH</div>
      <div class="val" id="d-cash"></div>
    </div>

    <div class="card">
      <div class="lbl">ASSETS</div>
      <div class="val" id="d-assets" style="color:var(--positive)"></div>
    </div>

    <div class="card">
      <div class="lbl">LIABILITIES</div>
      <div class="val" id="d-liab" style="color:var(--negative)"></div>
    </div>

    <div class="card">
      <div class="lbl">NET WORTH</div>
      <div class="val" id="d-netw"></div>
    </div>

  </div>


  <div class="row">

    <div class="panel chart-panel">
      <div id="d-debt" style="height:220px;"></div>
    </div>

    <div class="panel" id="d-goal-section">

      <h2>Savings goals</h2>

      <div id="d-goals-list"></div>

    </div>

  </div>



    <h2>
      Projected vs actual income (12 months)
    </h2>

    <div id="d-proj" style="height:220px;"></div>



  <div class="row">

    <div class="panel">

      <h2>Transactions this month</h2>

      <div class="tables-wrap">
        <table id="d-tx"></table>

        <div
          class="empty"
          id="d-tx-empty"
          style="display:none;"
        >
          No transactions this month
        </div>
      </div>

    </div>


    <div class="panel">

      <h2>Income this month</h2>

      <div class="tables-wrap">
        <table id="d-inc"></table>

        <div
          class="empty"
          id="d-inc-empty"
          style="display:none;"
        >
          No income logged this month
        </div>
      </div>

    </div>

  </div>

</section>

<section id="page-monthlybudget" class="page">

  <h1>Monthly Budget</h1>

  <select id="mb-month"></select>

  <div class="cards">

    <div class="card">
      <div class="lbl">MONTHLY BUDGET</div>
      <div
        class="val"
        id="mb-budget"
      ></div>
    </div>

    <div class="card">
      <div class="lbl">MONTHLY ACTUAL</div>
      <div
        class="val"
        id="mb-actual"
        style="color:var(--negative)"
      ></div>
    </div>

  </div>


  <div class="panel chart-panel">
    <div id="mb-chart" style="height:400px;"></div>
  </div>

  <div
    class="panel chart-panel"
    style="margin-top:20px;"
  >
    <table id="mb-table"></table>
  </div>

</section>

<section id="page-transactions" class="page">

  <h1>Transactions</h1>

  <p class="sub">
    Every expense you've logged, most recent first.
  </p>

  <div class="row">
    <div class="panel">
      <div id="t-donut" style="height:250px;"></div>
      <div id="t-donut-legend" class="donut-legend"></div>
    </div>
    <div class="panel"><div id="t-trend" style="height:250px;"></div></div>
  </div>


  <div class="panel">

    <div
      class="tables-wrap"
      style="max-height:500px;"
    >

      <table id="t-table"></table>

      <div
        class="empty"
        id="t-empty"
        style="display:none;"
      >
        No transactions yet
      </div>

    </div>

  </div>

</section>


<section id="page-income" class="page">

  <h1>Income</h1>

  <p class="sub" id="i-total"></p>

  <div class="panel">
    <div id="i-trend" style="height:250px;"></div>
  </div>

  <div
    class="panel"
    style="margin-top:20px;"
  >

    <div
      class="tables-wrap"
      style="max-height:500px;"
    >

      <table id="i-table"></table>

      <div
        class="empty"
        id="i-empty"
        style="display:none;"
      >
        No income yet
      </div>

    </div>

  </div>

</section>


<section id="page-bills" class="page">

  <h1>Bill Calendar</h1>

  <div class="cards">

    <div class="card">
      <div class="lbl">PAID</div>
      <div class="val" id="b-paid" style="color:var(--positive)"></div>
    </div>

    <div class="card">
      <div class="lbl">UNPAID</div>
      <div class="val" id="b-unpaid" style="color:var(--negative)"></div>
    </div>

    <div class="card">
      <div class="lbl">TOTAL DUE</div>
      <div class="val" id="b-due" style="color:var(--negative)"></div>
    </div>

  </div>


  <div class="row">

    <div class="panel">
      <div id="b-donut" style="height:220px;"></div>
    </div>

    <div class="panel">

      <div class="tables-wrap">
        <table id="b-table"></table>
      </div>

    </div>

  </div>

</section>


<section id="page-debt" class="page">

  <h1>Debt Tracker</h1>

  <div class="panel chart-panel">
    <div id="de-chart" style="height:250px;"></div>
  </div>

  <div
    class="panel chart-panel"
    style="margin-top:20px;"
  >

    <table id="de-table"></table>

    <div
      class="empty"
      id="de-empty"
      style="display:none;"
    >
      No debts logged
    </div>

  </div>

  <div
    class="panel"
    style="margin-top:20px;"
  >

    <h2>Monthly amortization schedule</h2>

    <div
      class="tables-wrap"
      style="max-height:400px;"
    >
      <table id="de-schedule"></table>
    </div>

  </div>

</section>


<section id="page-goals" class="page">

  <h1>Savings Goals</h1>

  <div id="g-list"></div>

  <div
    class="empty"
    id="g-empty"
    style="display:none;"
  >
    No savings goals yet
  </div>

</section>


<section id="page-networth" class="page">

  <h1>Net Worth</h1>

  <div class="panel chart-panel">
    <div id="n-chart" style="height:260px;"></div>
  </div>

  <div
    class="panel chart-panel"
    style="margin-top:20px;"
  >
    <table id="n-table"></table>
  </div>

</section>


<section id="page-setup" class="page">

  <h1>Setup</h1>

  <p class="sub">
    Read-only overview — edit these in the Google Sheet itself.
  </p>

  <div
    class="cards"
    style="margin-bottom:20px;"
  >

    <div class="card">
      <div class="lbl">CURRENCY</div>
      <div class="val" id="s-currency"></div>
    </div>

    <div class="card">
      <div class="lbl">TRACKING YEAR</div>
      <div class="val" id="s-year"></div>
    </div>

    <div class="card">
      <div class="lbl">STARTING CASH BALANCE</div>
      <div class="val" id="s-cash"></div>
    </div>

  </div>


  <h2>Categories</h2>

  <div
    id="s-categories"
    style="margin-bottom:20px;"
  ></div>


  <h2>Accounts</h2>

  <div
    id="s-accounts"
    style="margin-bottom:20px;"
  ></div>


  <h2>How each page works</h2>

  <div class="panel">

    <p>
      <b>Dashboard</b>
      — pick a month and see income, expenses, next bill due,
      spending by category, accounts, debt payoff and savings progress.
    </p>
      <p>
      <b>Monthly Budget</b>
      — Set a monthly budget per category and compare it to what you actually spent.
    </p>

    <p>
      <b>Transactions</b>
      — every expense you've logged, all-time,
      with a full category breakdown and 12-month trend.
    </p>

    <p>
      <b>Income</b>
      — every payment you've logged, all-time,
      with a 12-month trend.
    </p>

    <p>
      <b>Bill Calendar</b>
      — your recurring bills, due dates and payment status.
    </p>

    <p>
      <b>Debt Tracker</b>
      — each debt's balance, rate and projected month-by-month payoff.
    </p>

    <p>
      <b>Savings Goals</b>
      — progress toward each goal,
      and how much you still need to save monthly.
    </p>

    <p>
      <b>Net Worth</b>
      — assets, liabilities and net worth,
      tracked monthly.
    </p>

    <p class="sub">
      To add a category, account, or change currency/year,
      edit the Setup tab directly in the Google Sheet
      — this page is read-only.
    </p>

  </div>

</section>

</main>


<script>

var APP_DATA = <?!= data ?>;

var DATA = APP_DATA.dashboard;

var PAGE_CACHE = {

  monthlybudget: APP_DATA.monthlybudget,

  income: APP_DATA.income,

  bills: APP_DATA.bills,

  debt: APP_DATA.debt,

  goals: APP_DATA.goals,

  networth: APP_DATA.networth,

  setup: APP_DATA.setup

};


var CAT_COLORS = {

  "\u{1F3E0} Housing":"#7C93D8",
  "\u{1F6D2} Groceries":"#D98899",
  "\u{1F37D}\u{FE0F} Dining Out":"#D9A972",
  "\u{1F6CD}\u{FE0F} Shopping":"#C98F6B",
  "\u{1F68C} Transportation":"#6FAF9B",
  "\u{26FD} Fuel":"#9C8FD1",
  "\u{1F4A1} Utilities":"#93B27A",
  "\u{1F3AC} Entertainment":"#CE8A8A",
  "\u{1FA7A} Health & Wellness":"#93989E",
  "\u{1F3CB}\u{FE0F} Fitness":"#74A6C4",
  "\u{2708}\u{FE0F} Travel":"#C084A8",
  "\u{1F4F1} Subscriptions":"#C79768",
  "\u{1F486} Personal Care":"#7EBBA8",
  "\u{1F381} Gifts & Donations":"#A594D1",
  "\u{1F4E6} Other":"#97AD74"
};


function fmt(n){

  return (n||0).toLocaleString(
    'pt-PT',
    {
      minimumFractionDigits:2,
      maximumFractionDigits:2
    }
  )+" €";

}


var axisStyle = {

  textStyle:{
    color:'#868C99',
    fontName:'Inter',
    fontSize:10
  },

  gridlines:{
    color:'#1B1E24'
  },

  format:'short'
};


var axisStyleNoGrid = {

  textStyle:{
    color:'#868C99',
    fontName:'Inter',
    fontSize:10
  },

  gridlines:{
    color:'transparent'
  },

  slantedText:true,
  slantedTextAngle:45
};


var titleStyle = {

  fontName:'Inter',
  fontSize:12,
  bold:true,
  color:'#EDEEF2'

};


var legendTextStyle = {
  color:'#9AA0AC',
  fontName:'Inter'
};




function showLoading(v){

  document.getElementById("loading").style.display =
    v ? "flex" : "none";

}


document
  .getElementById("collapseBtn")
  .addEventListener("click",function(){

    var sb =
      document.getElementById("sidebar");

    sb.classList.toggle("collapsed");

    this.textContent =
      sb.classList.contains("collapsed")
        ? "→"
        : "←";

  });


// ============================================================
// MOBILE: START SIDEBAR COLLAPSED
// ============================================================

if (window.matchMedia("(max-width:800px)").matches) {

  var sb = document.getElementById("sidebar");
  var btn = document.getElementById("collapseBtn");

  sb.classList.add("collapsed");

  btn.textContent = "→";

}


document
  .querySelectorAll(".nav-item")
  .forEach(function(btn){

    btn.addEventListener("click",function(){
      switchPage(btn.dataset.page);
    });

  });


function switchPage(page){

  window.scrollTo(0, 0);

  document
    .querySelectorAll(".nav-item")
    .forEach(function(b){

      b.classList.toggle(
        "active",
        b.dataset.page === page
      );

    });


  document
    .querySelectorAll(".page")
    .forEach(function(s){

      s.classList.toggle(
        "active",
        s.id === "page-"+page
      );

    });


  if (
    page !== "dashboard" &&
    PAGE_CACHE[page]
  ){

    renderPage(
      page,
      PAGE_CACHE[page]
    );

    return;
  }


  if (page === "dashboard"){

    renderDashboard(DATA);

    return;
  }


  showLoading(true);


  google.script.run

    .withSuccessHandler(function(data){

      showLoading(false);

      PAGE_CACHE[page] = data;

      renderPage(page,data);

    })

    .withFailureHandler(function(err){

      showLoading(false);

      alert("Error: "+err.message);

    })

    .getPageData(page,{});

}


function renderPage(page,data){

  if(page === "monthlybudget")
    renderMonthlyBudget(data);

    else if(page === "transactions")
      renderTransactions(data);

    else if(page === "income")
      renderIncome(data);

    else if(page === "bills")
      renderBills(data);

    else if(page === "debt")
      renderDebt(data);

    else if(page === "goals")
      renderGoals(data);

    else if(page === "networth")
      renderNetWorth(data);

    else if(page === "setup")
      renderSetup(data);

  }


  var months = [

    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"

  ];


  var sel =
    document.getElementById("monthSel");


  months.forEach(function(m,i){

  var opt =
    document.createElement("option");

  opt.value = i+1;

  opt.textContent =
    m+" "+DATA.selectedYear;

  if(i+1 === DATA.selectedMonth)
    opt.selected = true;

  sel.appendChild(opt);

});


sel.addEventListener("change",function(){

  var month =
    parseInt(this.value) - 1;

  var data =
    DATA.dashboardMonths[month];

  if(!data) return;

  renderDashboard(data);

});


function renderDashboard(DATA){

  document.getElementById("d-income").textContent =
    fmt(DATA.income);

  document.getElementById("d-exp").textContent =
    fmt(DATA.expenses);

  document.getElementById("d-net").textContent =
    fmt(DATA.net);


  document.getElementById("d-bill").textContent =
    DATA.nextBill
      ? DATA.nextBill.name+" · "+fmt(DATA.nextBill.amount)
      : "No unpaid bills";


  document.getElementById("d-accounts").style.display =
    DATA.netWorth ? "grid" : "none";


  if(DATA.netWorth){

    document.getElementById("d-cash").textContent =
      fmt(DATA.netWorth.cash);

    document.getElementById("d-assets").textContent =
      fmt(DATA.netWorth.assets);

    document.getElementById("d-liab").textContent =
      fmt(DATA.netWorth.liabilities);

    document.getElementById("d-netw").textContent =
      fmt(DATA.netWorth.net);

  }


  document.getElementById("d-goal-section").style.display =
    DATA.goals && DATA.goals.length
      ? "block"
      : "none";


  var goalsList =
    document.getElementById("d-goals-list");

  goalsList.innerHTML = "";


  (DATA.goals || []).forEach(function(g){

    var div =
      document.createElement("div");

    div.style.marginBottom = "16px";


    div.innerHTML =

      "<div style='font-weight:600;'>"+
      g.name+
      "</div>"+

      "<div class='sub'>Saved "+
      fmt(g.current)+
      " of "+
      fmt(g.target)+
      " · "+
      g.monthsLeft+
      " months left</div>"+

      "<div class='bar-bg'>"+
      "<div class='bar-fill' style='width:"+
      Math.min(g.progress*100,100)+
      "%'></div>"+
      "</div>"+

      "<div class='val' style='color:#6C8EEF;margin-top:6px;font-size:15px;'>"+
      (g.progress*100).toFixed(1)+
      "%</div>";


    goalsList.appendChild(div);

  });


  fillTable(
    "d-tx",
    ["Date","Description","Category","Amount"],

    DATA.transactions.map(function(t){

      return [
        t.date,
        t.description,
        t.category,
        fmt(t.amount)
      ];

    }),

    "d-tx-empty"
  );


  fillTable(
    "d-inc",
    ["Date","Source","Description","Amount"],

    DATA.incomeRows.map(function(t){

      return [
        t.date,
        t.source,
        t.description,
        fmt(t.amount)
      ];

    }),

    "d-inc-empty"
  );

  renderCategoryTable(DATA.allCategories);

  google.charts.setOnLoadCallback(function(){

    drawDonut(
      "d-donut",
      DATA.pieCategories,
      "Spending by category"
    );


    drawHorizontalBar("d-trend", monthLabels(), DATA.monthly, "Monthly spending", ["#6C8EEF"]);


    drawLine(
      "d-debt",
      monthLabels(),
      DATA.debtSeries,
      "Debt payoff",
      "#6C8EEF"
    );


    if (isMobileLayout()) {
      var pValues = monthLabels().map(function(m,i){ return [Number(DATA.projectedIncome[i])||0, Number(DATA.incomeMonthly[i])||0]; });
      drawHorizontalBar("d-proj", monthLabels(), pValues, "", ["#565C69","#5FB88C"], ["Projected","Actual"]);
    } else {
      var pArr=[["Month","Projected","Actual"]];
      monthLabels().forEach(function(m,i){ pArr.push([m, Number(DATA.projectedIncome[i])||0, Number(DATA.incomeMonthly[i])||0]); });
      drawProjectedIncome("d-proj", pArr);
    }

  });

}

function renderMonthlyBudget(data){

  var sel =
    document.getElementById("mb-month");

  var months =
    data.months;

  sel.innerHTML = "";

  months.forEach(function(month){

    var opt =
      document.createElement("option");

    opt.value = month;
    opt.textContent = month;

    sel.appendChild(opt);

  });


  function update(){

    var month =
      sel.value;

    var d =
      data.data[month];

    if(!d) return;


    document.getElementById("mb-budget").textContent =
      fmt(d.budget);

    document.getElementById("mb-actual").textContent =
      fmt(d.actual);


    var labels =
      d.categories.map(function(r){
        return r.category;
      });


    var values =
      d.categories.map(function(r){
        return [
          r.budget,
          r.actual
        ];
      });


    google.charts.setOnLoadCallback(function(){

      drawHorizontalBar(
        "mb-chart",
        labels,
        values,
        "Budget vs Actual",
        ["#6C8EEF","#D97A82"],
        ["Budget","Actual"]
      );

    });


    fillTable(
      "mb-table",

      [
        "Category",
        "Budget",
        "Actual",
        "Difference"
      ],

      d.categories.map(function(r){

        return [
          r.category,
          fmt(r.budget),
          fmt(r.actual),
          fmt(r.difference)
        ];

      })
    );

  }


  sel.onchange =
    update;


  var currentMonth =
    new Date().getMonth();

  sel.value =
    months[currentMonth];


  update();

}
function renderTransactions(data){
  fillTable("t-table", ["Date","Description","Category","Account","Amount"],
    data.rows.map(function(t){ return [t.date,t.description,t.category,t.account,fmt(t.amount)]; }), "t-empty");

  google.charts.setOnLoadCallback(function(){
    drawDonut("t-donut", data.categoryTotals, "Spending by category (" + data.year + ")");
    if (data.monthly.length===0){ document.getElementById("t-trend").innerHTML='<div class="empty">No data yet</div>'; return; }

    var labels = data.monthly.map(function(m){return m.label;});
    var values = data.monthly.map(function(m){return m.amount;});

    if (isMobileLayout()) {
      drawHorizontalBar("t-trend", labels, values, "Last 12 months", ["#6C8EEF"]);
    } else {
      drawColumn("t-trend", labels, values, "Last 12 months", "#6C8EEF");
    }
  });
}


function renderIncome(data){

  document.getElementById("i-total").textContent =
    "Total logged: "+fmt(data.total);


  fillTable(
    "i-table",
    ["Date","Source","Description","Amount"],

    data.rows.map(function(t){

      return [
        t.date,
        t.source,
        t.description,
        fmt(t.amount)
      ];

    }),

    "i-empty"
  );


  google.charts.setOnLoadCallback(function(){

    if(data.monthly.length === 0){

      document.getElementById("i-trend").innerHTML =
        '<div class="empty">No data yet</div>';

      return;
    }

    var labels = data.monthly.map(function(m){
      return m.label;
    });

    var values = data.monthly.map(function(m){
      return m.amount;
    });

    if (isMobileLayout()) {

      drawHorizontalBar(
        "i-trend",
        labels,
        values,
        "Last 12 months",
        ["#6C8EEF"]
      );

    } else {

      drawColumn(
        "i-trend",
        labels,
        values,
        "Last 12 months",
        "#6C8EEF"
    );

  }

});

}


function renderBills(data){

  document.getElementById("b-paid").textContent =
    data.paid;

  document.getElementById("b-unpaid").textContent =
    data.unpaid;

  document.getElementById("b-due").textContent =
    fmt(data.totalDue);


  fillTable(
    "b-table",
    [
      "Bill",
      "Category",
      "Due Day",
      "Frequency",
      "Amount",
      "Status"
    ],

    data.rows.map(function(r){

      return [
        r.name,
        r.category,
        r.dueDay,
        r.frequency,
        fmt(r.amount),
        r.status
      ];

    })
  );


  google.charts.setOnLoadCallback(function(){

    var el =
      document.getElementById("b-donut");


    if(data.paid + data.unpaid === 0){

      el.innerHTML =
        '<div class="empty">No bills logged</div>';

      return;
    }


    el.innerHTML = "";


    drawBillDonut(
      "b-donut",
      data.paid,
      data.unpaid
    );

  });

}


function renderDebt(data){

  document.getElementById("de-empty").style.display =
    data.rows.length === 0
      ? "block"
      : "none";


  fillTable(
    "de-table",
    [
      "Debt",
      "Type",
      "Balance",
      "Rate",
      "Min. Payment",
      "Months to Payoff"
    ],

    data.rows.map(function(r){

      return [
        r.name,
        r.type,
        fmt(r.balance),
        (r.rate*100).toFixed(1)+"%",
        fmt(r.minPayment),
        r.months
      ];

    })
  );


  var headers =
    ["Month"].concat(
      data.schedule.map(function(s){
        return s.name;
      })
    );


  var rows =
    monthLabels().map(function(m,i){

      return [
        m
      ].concat(
        data.schedule.map(function(s){
          return fmt(s.values[i]);
        })
      );

    });


  fillTable(
    "de-schedule",
    headers,
    rows
  );


  google.charts.setOnLoadCallback(function(){

    drawLine(
      "de-chart",
      monthLabels(),
      data.totalSeries,
      "Total debt payoff, 12 months",
      "#6C8EEF"
    );

  });

}


function renderGoals(data){

  var el =
    document.getElementById("g-list");

  el.innerHTML = "";


  document.getElementById("g-empty").style.display =
    data.rows.length === 0
      ? "block"
      : "none";


  data.rows.forEach(function(g){

    var div =
      document.createElement("div");

    div.className = "panel";

    div.style.marginBottom = "16px";


    div.innerHTML =

      "<h2>"+g.name+"</h2>"+

      "<div class='bar-bg'>"+
      "<div class='bar-fill' style='width:"+
      Math.min(g.progress*100,100)+
      "%'></div>"+
      "</div>"+

      "<div class='val' style='color:#6C8EEF;margin:10px 0 14px;'>"+
      (g.progress*100).toFixed(1)+
      "%</div>"+

      "<table>"+
      "<tr>"+
      "<th>Target</th>"+
      "<th>Current</th>"+
      "<th>Needed / month</th>"+
      "<th>Months left</th>"+
      "</tr>"+

      "<tr>"+
      "<td>"+fmt(g.target)+"</td>"+
      "<td>"+fmt(g.current)+"</td>"+
      "<td>"+fmt(g.neededPerMonth)+"</td>"+
      "<td>"+g.monthsLeft+"</td>"+
      "</tr>"+

      "</table>";


    el.appendChild(div);

  });

}


function renderNetWorth(data){

  fillTable(
    "n-table",

    [
      "Month",
      "Cash",
      "Investments",
      "Property",
      "Assets",
      "Liabilities",
      "Net Worth"
    ],

    data.rows.map(function(r){

      return [
        r.month,
        fmt(r.cash),
        fmt(r.investments),
        fmt(r.property),
        fmt(r.assets),
        fmt(r.liabilities),
        fmt(r.net)
      ];

    })
  );


  google.charts.setOnLoadCallback(function(){

    var arr = [
      ["Month","Assets","Liabilities"]
    ];


    data.rows.forEach(function(r){

      arr.push([
        r.month,
        r.assets,
        r.liabilities
      ]);

    });


    drawNetWorth(
      "n-chart",
      arr
    );

  });

}


function renderSetup(data){

  document.getElementById("s-currency").textContent =
    data.currency;

  document.getElementById("s-year").textContent =
    data.year;

  document.getElementById("s-cash").textContent =
  fmt(data.startingCashBalance);

  document.getElementById("s-categories").innerHTML =
    data.categories.map(function(c){

      return "<span class='chip'>"+c+"</span>";

    }).join("");


  document.getElementById("s-accounts").innerHTML =
    data.accounts.map(function(a){

      return "<span class='chip'>"+a+"</span>";

    }).join("");

}


function monthLabels(){

  return [
    "Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"
  ];

}


function fillTable(
  id,
  headers,
  rows,
  emptyId
){

  var t =
    document.getElementById(id);


  t.innerHTML =
    "<tr>"+

    headers.map(function(h){

      return "<th>"+h+"</th>";

    }).join("")+

    "</tr>";


  if(emptyId){

    document.getElementById(emptyId).style.display =
      rows.length === 0
        ? "block"
        : "none";
  }


  rows.forEach(function(r){

    var tr =
      t.insertRow();


    r.forEach(function(v,i){

      var td =
        tr.insertCell(i);

      td.textContent = v;


      if(i === r.length-1){

        td.style.textAlign = "right";

      }

    });

  });

}


/* ============================================================
   MOBILE CHART HELPERS
   ============================================================ */

function isMobileLayout(){

  return window.matchMedia(
    "(max-width:800px)"
  ).matches;

}


function chartOptions(baseOptions){

  var options = {};

  Object.keys(baseOptions).forEach(function(key){
    options[key] = baseOptions[key];
  });


  if(isMobileLayout()){

    options.chartArea = {
      left: "8%",
      top: "18%",
      width: "84%",
      height: "65%"
    };

    options.titleTextStyle = {
      fontName:"Inter",
      fontSize:12,
      bold:true,
      color:"#EDEEF2"
    };

  }


  return options;

}

function drawDonut(elId, rows, title){

  var el = document.getElementById(elId);

  if (!rows || Object.keys(rows).length === 0) {

    el.innerHTML =
      '<div class="empty">No data yet</div>';

    return;
  }


  // ----------------------------------------------------------
  // Ordenar categorias por valor
  // ----------------------------------------------------------

  var categories = Object.keys(rows)
    .map(function(category){

      return {
        name: category,
        amount: Number(rows[category]) || 0
      };

    })
    .filter(function(item){

      return item.amount > 0;

    })
    .sort(function(a,b){

      return b.amount - a.amount;

    });


  if(categories.length === 0){

    el.innerHTML =
      '<div class="empty">No data yet</div>';

    return;
  }


  // ----------------------------------------------------------
  // Dados para o Google Charts
  // ----------------------------------------------------------

  var arr = [
    ["Categoria","Valor"]
  ];

  var colors = [];
  var total = 0;


  categories.forEach(function(item){

    arr.push([
      item.name,
      item.amount
    ]);

    colors.push(
      CAT_COLORS[item.name] || "#565C69"
    );

    total += item.amount;

  });


  // ----------------------------------------------------------
  // Opções do donut
  // ----------------------------------------------------------

  var options = {

    title:title,

    titleTextStyle:{
      fontName:"Inter",
      fontSize:13,
      bold:true,
      color:"#EDEEF2"
    },

    pieHole:0.55,

    colors:colors,

    pieSliceText:"percentage",

    pieSliceTextStyle:{
      color:"#0D0F13"
    },

    pieSliceBorderColor:"#15171D",

    // SEM LEGENDA
    legend:"none",

    chartArea:{
      left:"5%",
      top:"8%",
      width:"90%",
      height:"86%"
    },

    tooltip:{
      text:"percentage"
    },

    backgroundColor:"transparent"

  };


  // ----------------------------------------------------------
  // Ajustes mobile
  // ----------------------------------------------------------

  if(isMobileLayout()){

    options.chartArea = {

      left:"5%",
      top:"15%",
      width:"90%",
      height:"78%"

    };

    options.titleTextStyle = {

      fontName:"Inter",
      fontSize:12,
      bold:true,
      color:"#EDEEF2"

    };

  }


  // ----------------------------------------------------------
  // Desenhar donut
  // ----------------------------------------------------------

  var chart =
    new google.visualization.PieChart(el);


  chart.draw(
    google.visualization.arrayToDataTable(arr),
    options
  );

}

function renderCategoryTable(categories){

  var table =
    document.getElementById("d-catlist");

  if(!table) return;


  if(!categories || categories.length === 0){

    table.innerHTML =
      "<tr>"+
        "<td class='empty'>No expenses yet</td>"+
      "</tr>";

    return;
  }


  var total =
    categories.reduce(function(sum,item){

      return sum + (Number(item.amount) || 0);

    },0);


  var html =

    "<tr>"+
      "<th>Category</th>"+
      "<th>Amount</th>"+
      "<th>%</th>"+
    "</tr>";


  categories.forEach(function(item){

    var amount =
      Number(item.amount) || 0;

    var percentage =
      total > 0
        ? (amount / total * 100)
        : 0;


    html +=

      "<tr>"+

        "<td>"+

          "<span style='display:inline-flex;align-items:center;gap:7px;'>"+

            "<span style='width:9px;height:9px;border-radius:50%;background:"+
              (CAT_COLORS[item.name] || "#565C69")+
              ";display:inline-block;flex-shrink:0;'></span>"+

            item.name+

          "</span>"+

        "</td>"+

        "<td>"+

          fmt(amount)+

        "</td>"+

        "<td>"+

          percentage.toFixed(1)+
          "%"+

        "</td>"+

      "</tr>";

  });


  table.innerHTML = html;

}

function drawColumn(
    elId,
    labels,
    values,
    title,
    color
  ){

  var arr = [
    ["Month","Value"]
  ];


  labels.forEach(function(l,i){

    arr.push([
      l,
      values[i]
    ]);

  });


  var options = {

    title:title,

    titleTextStyle:titleStyle,

    colors:[color],

    legend:"none",

    bar:{
      groupWidth:"55%"
    },

    hAxis:axisStyleNoGrid,

    vAxis:axisStyle,

    chartArea:{
      width:"85%",
      height:"65%"
    },

    backgroundColor:"transparent"

  };


  if(isMobileLayout()){

    options.chartArea = {
      left:"10%",
      top:"18%",
      width:"84%",
      height:"62%"
    };

    options.hAxis = {
    textStyle:{
      color:"#868C99",
      fontName:"Inter",
      fontSize:9
    },
    gridlines:{
      color:"transparent"
    },
    slantedText:true,
    slantedTextAngle:30,
    showTextEvery:1
  };

    options.vAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:9
      },
      gridlines:{
        color:"#1B1E24"
      },
      format:"short"
    };

  }


  new google.visualization.ColumnChart(
    document.getElementById(elId)
  ).draw(
    google.visualization.arrayToDataTable(arr),
    options
  );

}


function drawHorizontalBar(elId, labels, values, title, colors, seriesNames){

  var el = document.getElementById(elId);
  var isMulti = Array.isArray(values[0]);

  var arr;

  if (isMulti) {

    arr = [["Month"].concat(seriesNames)];

    labels.forEach(function(l,i){
      arr.push([l].concat(values[i]));
    });

  } else {

    arr = [["Month","Value"]];

    labels.forEach(function(l,i){
      arr.push([l, Number(values[i])||0]);
    });

  }

  var fontSize = isMobileLayout() ? 9 : 10;

  var options = {
    title:title,
    titleTextStyle:titleStyle,
    colors:colors,
    legend:isMulti
      ? {
          position:"top",
          textStyle:{
            fontSize:fontSize,
            color:"#9AA0AC"
          }
        }
      : "none",

    bar:{
      groupWidth:isMulti ? "60%" : "65%"
    },

    chartArea:{
      left:"18%",
      top:isMulti ? "10%" : "8%",
      width:"78%",
      height:isMulti ? "75%" : "84%"
    },

    backgroundColor:"transparent"
  };


  if(isMobileLayout()){

    // MOBILE → horizontal bars

    options.hAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:fontSize
      },
      gridlines:{
        color:"#1B1E24"
      },
      format:"short"
    };

    options.vAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:fontSize
      },
      gridlines:{
        color:"transparent"
      }
    };

  } else {

    // PC → vertical bars

    options.hAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:fontSize
      },
      gridlines:{
        color:"transparent"
      }
    };

    options.vAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:fontSize
      },
      gridlines:{
        color:"#1B1E24"
      },
      format:"short"
    };

  }


  var chartType =
    isMobileLayout()
      ? google.visualization.BarChart
      : google.visualization.ColumnChart;


  new chartType(el).draw(
    google.visualization.arrayToDataTable(arr),
    options
  );

}

function drawLine(
  elId,
  labels,
  values,
  title,
  color
){

  var arr = [
    ["Month","Value"]
  ];


  labels.forEach(function(l,i){

    arr.push([
      l,
      values[i]
    ]);

  });


  var options = {

    title:title,

    titleTextStyle:titleStyle,

    colors:[color],

    legend:"none",

    curveType:"function",

    hAxis:axisStyleNoGrid,

    vAxis:axisStyle,

    chartArea:{
      width:"85%",
      height:"65%"
    },

    backgroundColor:"transparent"

  };


  if(isMobileLayout()){

    options.chartArea = {
      left:"10%",
      top:"18%",
      width:"84%",
      height:"62%"
    };

    options.hAxis = {
    textStyle:{
      color:"#868C99",
      fontName:"Inter",
      fontSize:9
    },
    gridlines:{
      color:"transparent"
    },
    slantedText:true,
    slantedTextAngle:30,
    showTextEvery:1
  };

    options.vAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:9
      },
      gridlines:{
        color:"#1B1E24"
      },
      format:"short"
    };

  }


  new google.visualization.LineChart(
    document.getElementById(elId)
  ).draw(
    google.visualization.arrayToDataTable(arr),
    options
  );

}


function drawProjectedIncome(
  elId,
  dataArray
){

  var options = {

    colors:[
      "#565C69",
      "#5FB88C"
    ],

    legend:{
      position:"top",
      textStyle:{
        color:"#9AA0AC"
      }
    },

    bar:{
      groupWidth:"65%"
    },

    hAxis:axisStyleNoGrid,

    vAxis:axisStyle,

    chartArea:{
      width:"85%",
      height:"65%"
    },

    backgroundColor:"transparent"

  };


  if(isMobileLayout()){

    options.legend = {
      position:"bottom",
      alignment:"center",
      textStyle:{
        fontSize:9,
        color:"#9AA0AC"
      }
    };

    options.chartArea = {
      left:"10%",
      top:"12%",
      width:"84%",
      height:"60%"
    };

    options.hAxis = {
    textStyle:{
      color:"#868C99",
      fontName:"Inter",
      fontSize:9
    },
    gridlines:{
      color:"transparent"
    },
    slantedText:true,
    slantedTextAngle:30,
    showTextEvery:1
  };

    options.vAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:9
      },
      gridlines:{
        color:"#1B1E24"
      },
      format:"short"
    };

  }


  var chartType =
    isMobileLayout()
      ? google.visualization.BarChart
      : google.visualization.ColumnChart;

  new chartType(
    document.getElementById(elId)
  ).draw(
    google.visualization.arrayToDataTable(dataArray),
    options
  );

}




function drawBillDonut(
  elId,
  paid,
  unpaid
){

  var el =
    document.getElementById(elId);

  var options = {

    pieHole:0.5,

    colors:[
      "#5FB88C",
      "#D97A82"
    ],

    pieSliceTextStyle:{
      color:"#0D0F13"
    },

    pieSliceBorderColor:"#15171D",

    legend:{
      position:"right",
      textStyle:{
        color:"#9AA0AC"
      }
    },

    chartArea:{
      width:"85%",
      height:"80%"
    },

    backgroundColor:"transparent"

  };


  if(isMobileLayout()){

    options.legend = {
      position:"bottom",
      alignment:"center",
      textStyle:{
        fontSize:9,
        color:"#9AA0AC"
      }
    };

    options.chartArea = {
      left:"4%",
      top:"10%",
      width:"92%",
      height:"62%"
    };

  }


  new google.visualization.PieChart(el).draw(
    google.visualization.arrayToDataTable([
      ["Status","Count"],
      ["Paid",paid],
      ["Unpaid",unpaid]
    ]),
    options
  );

}


function drawNetWorth(
  elId,
  dataArray
){

  var options = {

    colors:[
      "#5FB88C",
      "#D97A82"
    ],

    legend:{
      position:"top",
      textStyle:{
        color:"#9AA0AC"
      }
    },

    hAxis:axisStyleNoGrid,

    vAxis:axisStyle,

    chartArea:{
      width:"85%",
      height:"65%"
    },

    backgroundColor:"transparent"

  };


  if(isMobileLayout()){

    options.legend = {
      position:"bottom",
      alignment:"center",
      textStyle:{
        fontSize:9,
        color:"#9AA0AC"
      }
    };

    options.chartArea = {
      left:"10%",
      top:"12%",
      width:"84%",
      height:"60%"
    };

    options.hAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:9
      },
      gridlines:{
        color:"#1B1E24"
      },
      format:"short"
    };

    options.vAxis = {
      textStyle:{
        color:"#868C99",
        fontName:"Inter",
        fontSize:9
      },
      gridlines:{
        color:"transparent"
      }
    };

  }


  new google.visualization.AreaChart(
    document.getElementById(elId)
  ).draw(
    google.visualization.arrayToDataTable(dataArray),
    options
  );

}


google.charts.load(
  "current",
  {
    packages:["corechart"]
  }
);


google.charts.setOnLoadCallback(function(){

  renderDashboard(DATA);

});

</script>

</body>
</html>`;

var QUICK_ENTRY_HTML = `<!DOCTYPE html><html><head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;}
:root{
  --bg:#0D0F13;
  --surface:#15171D;
  --surface-alt:#1A1D24;
  --border:#23262E;
  --text:#EDEEF2;
  --text-muted:#868C99;
  --accent:#6C8EEF;
  --positive:#5FB88C;
  --negative:#D97A82;
}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:var(--bg);margin:0;padding:24px;color:var(--text);}
h1{font-size:18px;font-weight:600;margin:0 0 20px;}
.toggle{display:flex;gap:8px;margin-bottom:20px;}
.toggle button{flex:1;padding:12px;border:1px solid var(--border);border-radius:8px;background:var(--surface);font-size:15px;cursor:pointer;color:var(--text-muted);}
.toggle button.active{border-color:var(--accent);color:var(--accent);font-weight:600;background:#161B2C;}
label{font-size:12px;color:var(--text-muted);font-weight:600;display:block;margin:14px 0 4px;}
input,select{width:100%;padding:12px;font-size:16px;border:1px solid var(--border);border-radius:8px;background:var(--surface);color:var(--text);}
.submit{width:100%;padding:15px;margin-top:24px;background:var(--accent);color:#0D0F13;border:none;border-radius:8px;font-size:16px;font-weight:600;cursor:pointer;}
.submit:disabled{background:#3A4256;color:var(--text-muted);}
.toast{position:fixed;bottom:24px;left:24px;right:24px;padding:14px;border-radius:8px;text-align:center;font-size:14px;display:none;}
.toast.ok{background:var(--positive);color:#0D0F13;}
.toast.err{background:var(--negative);color:#0D0F13;}
.income-only, .expense-only{display:none;}


</style></head><body>
<h1>Quick Entry</h1>
<div class="toggle">
  <button id="btnExpense" class="active">Expense</button>
  <button id="btnIncome">Income</button>
</div>

<div id="expenseFields">
  <div class="field">
    <label for="entryDate">Date</label>
    <input type="date" id="entryDate">
  </div>
  <label>Amount</label>
  <input type="number" id="amount" inputmode="decimal" placeholder="0.00">
  <label>Category</label>
  <select id="category"></select>
  <label>Account</label>
  <select id="account"></select>
  <label>Description</label>
  <input type="text" id="description" placeholder="Optional">
</div>

<div id="incomeFields" style="display:none;">
  <div class="field">
    <label for="entryDate">Date</label>
    <input type="date" id="entryDate">
  </div>
  <label>Amount</label>
  <input type="number" id="amountInc" inputmode="decimal" placeholder="0.00">
  <label>Source</label>
  <input type="text" id="source" placeholder="e.g. Salary">
  <label>Description</label>
  <input type="text" id="descriptionInc" placeholder="Optional">
</div>

<button class="submit" id="submitBtn">Log it</button>
<div class="toast" id="toast"></div>

<script>

var today = new Date();
var yyyy = today.getFullYear();
var mm = String(today.getMonth() + 1).padStart(2, "0");
var dd = String(today.getDate()).padStart(2, "0");

document.getElementById("entryDate").value = yyyy + "-" + mm + "-" + dd;

var categories = <?!= categories ?>;
var accounts = <?!= accounts ?>;
var mode = "expense";

var catSel = document.getElementById("category");
categories.forEach(function(c){ var o=document.createElement("option"); o.value=c; o.textContent=c; catSel.appendChild(o); });
var accSel = document.getElementById("account");
accounts.forEach(function(a){ var o=document.createElement("option"); o.value=a; o.textContent=a; accSel.appendChild(o); });

document.getElementById("btnExpense").addEventListener("click", function(){
  mode = "expense";
  this.classList.add("active");
  document.getElementById("btnIncome").classList.remove("active");
  document.getElementById("expenseFields").style.display = "block";
  document.getElementById("incomeFields").style.display = "none";
});
document.getElementById("btnIncome").addEventListener("click", function(){
  mode = "income";
  this.classList.add("active");
  document.getElementById("btnExpense").classList.remove("active");
  document.getElementById("incomeFields").style.display = "block";
  document.getElementById("expenseFields").style.display = "none";
});

function showToast(msg, ok){
  var t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast " + (ok ? "ok" : "err");
  t.style.display = "block";
  setTimeout(function(){ t.style.display = "none"; }, 2500);
}

document.getElementById("submitBtn").addEventListener("click", function(){
  var btn = this;
  btn.disabled = true;

  var data;

  if (mode === "expense") {
    data = {
      type: "expense",
      date: document.getElementById("entryDate").value,
      amount: document.getElementById("amount").value,
      category: document.getElementById("category").value,
      account: document.getElementById("account").value,
      description: document.getElementById("description").value
    };
  } else {
    data = {
      type: "income",
      date: document.getElementById("entryDate").value,
      amount: document.getElementById("amountInc").value,
      source: document.getElementById("source").value,
      description: document.getElementById("descriptionInc").value
};
  }

  google.script.run
    .withSuccessHandler(function(result){
      btn.disabled = false;

      if (result.status === "ok") {
        showToast("Logged!", true);

        document.getElementById("amount").value = "";
        document.getElementById("description").value = "";
        document.getElementById("amountInc").value = "";
        document.getElementById("source").value = "";
        document.getElementById("descriptionInc").value = "";
      } else {
        showToast("Error: " + (result.message || "unknown"), false);
      }
    })
    .withFailureHandler(function(error){
      btn.disabled = false;
      showToast("Error: " + error.message, false);
    })
    .saveQuickEntry(data);
});

</script></body></html>`;
