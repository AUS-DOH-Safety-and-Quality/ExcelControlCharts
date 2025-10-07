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
const aggregations = { numerators: "sum", denominators: "sum" };

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-plot").onclick = () => tryCatch(createPlot);
    // Populate table selector when dropdown is clicked
    // Populate column selectors when table selection changes
    document.getElementById("table-selector").onclick = () => {tryCatch(updateTableSelector); tryCatch(updateColumnSelectors)};
    // Initial population of table selector
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

async function updateColumnSelectors() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement).value;
    if (!selectedTableName) {
      throw new Error("No table selected");
    }
    const table = currentWorksheet.tables.getItem(selectedTableName);
    const columns = table.columns.load("items/name");
    await context.sync();
    const categorySelector = document.getElementById("category-selector") as HTMLSelectElement;
    const numeratorSelector = document.getElementById("numerator-selector") as HTMLSelectElement;
    const denominatorSelector = document.getElementById("denominator-selector") as HTMLSelectElement;
    categorySelector.innerHTML = '<option value="" disabled selected>Select category</option>';
    numeratorSelector.innerHTML = '<option value="" disabled selected>Select numerator</option>';
    denominatorSelector.innerHTML = '<option value="" disabled selected>Select denominator</option>';
    columns.items.forEach(column => {
      const option1 = document.createElement("option");
      option1.value = column.name;
      option1.text = column.name;
      categorySelector.appendChild(option1);

      const option2 = document.createElement("option");
      option2.value = column.name;
      option2.text = column.name;
      numeratorSelector.appendChild(option2);

      const option3 = document.createElement("option");
      option3.value = column.name;
      option3.text = column.name;
      denominatorSelector.appendChild(option3);
    });
  });
}

async function createPlot() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement).value;
    if (!selectedTableName) {
      throw new Error("No table selected");
    }
    const table = currentWorksheet.tables.getItem(selectedTableName);
    const selectedCategoryColumn = (document.getElementById("category-selector") as HTMLSelectElement).value;
    const selectedNumeratorColumn = (document.getElementById("numerator-selector") as HTMLSelectElement).value;
    const selectedDenominatorColumn = (document.getElementById("denominator-selector") as HTMLSelectElement).value;

    const categoryColumn = table.columns.getItem(selectedCategoryColumn).getDataBodyRange().load("values");
    const numeratorsColumn = table.columns.getItem(selectedNumeratorColumn).getDataBodyRange().load("values");
    const denominatorsColumn = table.columns.getItem(selectedDenominatorColumn).getDataBodyRange().load("values");
    await context.sync();

    const rawData = categoryColumn.values.flat().map((cat, i) => ({
      categories: fromExcelDate(cat), numerators: numeratorsColumn.values.flat()[i], denominators: denominatorsColumn.values.flat()[i]
    }));

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
