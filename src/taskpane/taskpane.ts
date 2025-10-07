import { makeConstructorArgs, makeUpdateValues } from "../utilities/commonUtils";
import { Visual } from "../PowerBI-SPC/src/visual";
import { defaultSettings, type defaultSettingsType } from "../PowerBI-SPC/src/settings";
import * as d3 from "../PowerBI-SPC/src/D3 Plotting Functions/D3 Modules"

const spcDiv = d3.select(document.body)
                  .append('div')
                  .classed('spc-container', true)
                  .attr("hidden", true)
                  .node();

const spcVisual = new Visual(makeConstructorArgs(spcDiv));

const inputSettings = Object.fromEntries(Object.keys(defaultSettings).map((settingGroupName) => {
  return [settingGroupName, Object.fromEntries(Object.keys(defaultSettings[settingGroupName]).map((settingName) => {
    return [settingName, defaultSettings[settingGroupName][settingName]["default"]];
  }))];
})) as defaultSettingsType;
inputSettings.canvas.left_padding += 50;
inputSettings.canvas.lower_padding += 50;
const aggregations = { numerators: "sum" };

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-plot").onclick = () => tryCatch(createPlot);
    // Populate table selector when dropdown is clicked
    document.getElementById("table-selector").onclick = () => tryCatch(updateTableSelector);
    tryCatch(updateTableSelector);
  }
});

function fromExcelDate(excelDate: number): Date {
  return new Date((excelDate - (25567 + 2)) * 86400 * 1000);
}

async function updateTableSelector() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const tables = currentWorksheet.tables.load("items/name");
    await context.sync();
    const tableSelector = document.getElementById("table-selector") as HTMLSelectElement;
    tableSelector.innerHTML = '<option value="" disabled selected>Select a table</option>';
    tables.items.forEach(table => {
      const option = document.createElement("option");
      option.value = table.name;
      option.text = table.name;
      tableSelector.appendChild(option);
    });
    tableSelector.onchange = () => {
      const selectedTable = tableSelector.value;
      if (selectedTable) {
        document.getElementById("create-plot").removeAttribute("disabled");
      } else {
        document.getElementById("create-plot").setAttribute("disabled", "true");
      }
    };
    if (tables.items.length > 0) {
      tableSelector.value = tables.items[0].name;
      document.getElementById("create-plot").removeAttribute("disabled");
    }
  });
}

async function createPlot() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const table = currentWorksheet.tables.getItem("Table1");
    const categoryColumn = table.columns.getItem("categories").getDataBodyRange().load("values");
    const numeratorsColumn = table.columns.getItem("numerators").getDataBodyRange().load("values");
    await context.sync();

    const rawData = categoryColumn.values.flat().map((cat, i) => ({categories: fromExcelDate(cat), numerators: numeratorsColumn.values.flat()[i]}));

    var updateArgs = {
      dataViews: makeUpdateValues(rawData, inputSettings, aggregations).dataViews,
      viewport: { width: 640, height: 480 },
      type: 2,
      headless: true,
      frontend: true
    };

    spcVisual.update(updateArgs as any);
    spcVisual.svg
              .append("rect")
              .attr("width", "100%")
              .attr("height", "100%")
              .attr("fill", "white")
              .lower();

    var image = currentWorksheet.shapes.addImage(btoa(spcVisual.svg.node().outerHTML));
    image.name = "Image";
    image.top = 10;
    image.left = 200;

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
