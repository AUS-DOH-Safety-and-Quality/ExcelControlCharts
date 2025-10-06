/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
  }
});

async function createTable() {
    await Excel.run(async (context) => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const dummyTable = currentWorksheet.tables.add("A1:C1", true /*hasHeaders*/);
        dummyTable.name = "DummyTable";

        dummyTable.getHeaderRowRange().values = [["Date", "Numerator", "Denominator"]];

        dummyTable.rows.add(null /*add at the end*/, [
            ["1/1/2017", "120", "180"],
            ["1/2/2017", "142.33", "200"],
            ["1/5/2017", "27.9", "50"],
            ["1/10/2017", "33", "60"],
            ["1/11/2017", "350.1", "400"],
            ["1/15/2017", "135", "150"],
            ["1/15/2017", "97.88", "100"]
        ]);

        dummyTable.columns.getItemAt(1).getRange().numberFormat = [['##0.00']];
        dummyTable.columns.getItemAt(2).getRange().numberFormat = [['##0.00']];
        dummyTable.getRange().format.autofitColumns();
        dummyTable.getRange().format.autofitRows();

        await context.sync();
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
