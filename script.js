$("#run").on("click", () => tryCatch(run));
$("#run1").on("click", () => tryCatch(run1));
$("#run2").on("click", () => tryCatch(run2));
$("#run3").on("click", () => tryCatch(run3));
$("#run4").on("click", () => tryCatch(run4));

async function run() {
  await Excel.run(async (context) => {
    const xlWorkbook = context.workbook;
    let sheet = context.workbook.worksheets.getItem("Data");
    let expensesTable = sheet.tables.getItem("Table1");
    let a = $("#add_location").val();
    const xlSheet = xlWorkbook.worksheets;
    xlWorkbook.load("properties/");
    xlSheet.load(["name", "items", "id"]);
    let worksheetCount = xlSheet.getCount();
    await context.sync();

    for (let i in xlSheet.items) {
      if (xlSheet.items[i].name === "Data") {
        if (a !== "") {
          expensesTable.rows.add(null, [[`${a}`]], true);
        }
      }
    }
    $("#add_location").val("");
  });
}
async function run1() {
  await Excel.run(async (context) => {
    const xlWorkbook = context.workbook;
    let sheet = context.workbook.worksheets.getItem("Data");
    let expensesTable = sheet.tables.getItem("Table2");
    let a = $("#add_floor").val();
    const xlSheet = xlWorkbook.worksheets;
    xlWorkbook.load("properties/");
    xlSheet.load(["name", "items", "id"]);
    let worksheetCount = xlSheet.getCount();
    await context.sync();

    for (let i in xlSheet.items) {
      if (xlSheet.items[i].name === "Data") {
        if (a !== "") {
          expensesTable.rows.add(null, [[`${a}`]], true);
        }
      }
    }
    $("#add_floor").val("");
  });
}
async function run2() {
  await Excel.run(async (context) => {
    const xlWorkbook = context.workbook;
    let sheet = context.workbook.worksheets.getItem("Data");
    let expensesTable = sheet.tables.getItem("Table14");
    let a = $("#add_tiur").val();
    const xlSheet = xlWorkbook.worksheets;
    xlWorkbook.load("properties/");
    xlSheet.load(["name", "items", "id"]);
    let worksheetCount = xlSheet.getCount();
    await context.sync();

    for (let i in xlSheet.items) {
      if (xlSheet.items[i].name === "Data") {
        if (a !== "") {
          expensesTable.rows.add(null, [[`${a}`]], true);
        }
      }
    }
    $("#add_tiur").val("");
  });
}
async function run3() {
  await Excel.run(async (context) => {
    const xlWorkbook = context.workbook;
    let sheet = context.workbook.worksheets.getItem("Data");
    let expensesTable = sheet.tables.getItem("Table25");
    let a = $("#add_tiur1").val();
    const xlSheet = xlWorkbook.worksheets;
    xlWorkbook.load("properties/");
    xlSheet.load(["name", "items", "id"]);
    let worksheetCount = xlSheet.getCount();
    await context.sync();

    for (let i in xlSheet.items) {
      if (xlSheet.items[i].name === "Data") {
        if (a !== "") {
          expensesTable.rows.add(null, [[`${a}`]], true);
        }
      }
    }
    $("#add_tiur1").val("");
  });
}
async function run4() {
  await Excel.run(async (context) => {
    const xlWorkbook = context.workbook;
    let sheet = context.workbook.worksheets.getItem("Data");
    let expensesTable = sheet.tables.getItem("Table3");
    let a = $("#add_tiur2").val();
    const xlSheet = xlWorkbook.worksheets;
    xlWorkbook.load("properties/");
    xlSheet.load(["name", "items", "id"]);
    let worksheetCount = xlSheet.getCount();
    await context.sync();

    for (let i in xlSheet.items) {
      if (xlSheet.items[i].name === "Data") {
        if (a !== "") {
          expensesTable.rows.add(null, [[`${a}`]], true);
        }
      }
    }
    $("#add_tiur2").val("");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
