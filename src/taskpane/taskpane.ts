import { makeConstructorArgs, makeUpdateValues } from "../utilities/commonUtils";
import { renderSpcDataSettings } from "../utilities/renderSpcDataSettings";
import { Visual as spcVisualClass } from "../PowerBI-SPC/src/visual";
import { Visual as funnelVisualClass } from "../PowerBI-Funnels/src/visual";
import { defaultSettings as spcDefaultSettings, type defaultSettingsType as spcDefaultSettingsType } from "../PowerBI-SPC/src/settings";
import { defaultSettings as funnelDefaultSettings, type defaultSettingsType as funnelDefaultSettingsType } from "../PowerBI-Funnels/src/settings";


const spcDiv = document.createElement('div');
spcDiv.className = 'spc-container';
spcDiv.setAttribute("hidden", "true");

const funnelDiv = document.createElement('div');
funnelDiv.className = 'funnel-container';
funnelDiv.setAttribute("hidden", "true");

const spcVisual = new spcVisualClass(makeConstructorArgs(spcDiv));
const funnelVisual = new funnelVisualClass(makeConstructorArgs(funnelDiv));

const spcInputSettings = Object.fromEntries(Object.keys(spcDefaultSettings).map((settingGroupName) => {
  return [settingGroupName, Object.fromEntries(Object.keys(spcDefaultSettings[settingGroupName]).map((settingName) => {
    return [settingName, spcDefaultSettings[settingGroupName][settingName]["default"]];
  }))];
})) as spcDefaultSettingsType;
spcInputSettings.canvas.left_padding += 50;
spcInputSettings.canvas.lower_padding += 50;

const funnelInputSettings = Object.fromEntries(Object.keys(funnelDefaultSettings).map((settingGroupName) => {
  return [settingGroupName, Object.fromEntries(Object.keys(funnelDefaultSettings[settingGroupName]).map((settingName) => {
    return [settingName, funnelDefaultSettings[settingGroupName][settingName]["default"]];
  }))];
})) as funnelDefaultSettingsType;
funnelInputSettings.canvas.left_padding += 50;
funnelInputSettings.canvas.lower_padding += 25;

const aggregations: Record<string, string> = { numerators: "sum", denominators: "sum", xbar_sds: "first" };

function getSelectedSpcChartType(): string {
  const el = document.getElementById("spc-chart-type") as HTMLSelectElement | null;
  return el?.value || "i";
}

function resetSelectToPlaceholder(id: string) {
  const el = document.getElementById(id) as HTMLSelectElement | null;
  if (!el) return;
  // Prefer selecting the explicit placeholder option (value="")
  el.value = "";
  // Some browsers keep prior selection if value doesn't match; force index 0 as fallback
  if (el.selectedIndex > 0) {
    el.selectedIndex = 0;
  }
}

function updateSdSelectorVisibility() {
  const chartFamily = (document.getElementById("controlchart-selector") as HTMLInputElement | null)?.value;
  const isSpc = chartFamily === "spc";
  const isXbar = getSelectedSpcChartType() === "xbar";
  const sdField = document.getElementById("sd-selector-field") as HTMLElement | null;
  if (sdField) {
    const shouldShow = isSpc && isXbar;
    sdField.hidden = !shouldShow;
    sdField.style.display = shouldShow ? "" : "none";
    if (!shouldShow) {
      resetSelectToPlaceholder("sd-selector");
    }
  }
}

function isDenominatorRequired(): boolean {
  const chartFamily = (document.getElementById("controlchart-selector") as HTMLInputElement | null)?.value;
  if (chartFamily !== "spc") return true;
  // Denominators required for ratio-based charts. MR can also be ratio-based (MR of rates),
  // so we require it here to match expected UX/workflows.
  const chartType = getSelectedSpcChartType();
  return ["p", "pp", "u", "up", "xbar", "s", "mr"].includes(chartType);
}

function updateDenominatorSelectorVisibility() {
  const denomField = document.getElementById("denominator-selector-field") as HTMLElement | null;
  if (!denomField) return;
  const shouldShow = isDenominatorRequired();
  denomField.hidden = !shouldShow;
  denomField.style.display = shouldShow ? "" : "none";
  if (!shouldShow) {
    resetSelectToPlaceholder("denominator-selector");
  }
}

function parseBoolean(value: string | undefined | null, fallback: boolean): boolean {
  if (value === "true") return true;
  if (value === "false") return false;
  return fallback;
}

function parseNumber(value: string | undefined | null, fallback: number, opts?: { min?: number; max?: number }): number {
  const raw = (value ?? "").trim();
  const parsed = raw.length ? Number(raw) : NaN;
  let next = Number.isFinite(parsed) ? parsed : fallback;
  if (opts?.min !== undefined) next = Math.max(opts.min, next);
  if (opts?.max !== undefined) next = Math.min(opts.max, next);
  return next;
}

function parseOptionalNumber(value: string | undefined | null, opts?: { min?: number; max?: number }): number | null {
  const raw = (value ?? "").trim();
  if (!raw.length) return null;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return null;
  let next = parsed;
  if (opts?.min !== undefined) next = Math.max(opts.min, next);
  if (opts?.max !== undefined) next = Math.min(opts.max, next);
  return next;
}

function updateSpcInputSettingsFromUi() {
  const chartTypeSel = document.getElementById("spc-chart-type") as HTMLSelectElement | null;
  const outliersInLimitsSel = document.getElementById("spc-outliers-in-limits") as HTMLSelectElement | null;
  const multiplierInput = document.getElementById("spc-multiplier") as HTMLInputElement | null;
  const sigFigsInput = document.getElementById("spc-sig-figs") as HTMLInputElement | null;
  const percLabelsSel = document.getElementById("spc-perc-labels") as HTMLSelectElement | null;
  const splitOnClickSel = document.getElementById("spc-split-on-click") as HTMLSelectElement | null;
  const numPointsSubsetInput = document.getElementById("spc-num-points-subset") as HTMLInputElement | null;
  const subsetPointsFromSel = document.getElementById("spc-subset-points-from") as HTMLSelectElement | null;
  const showDateSel = document.getElementById("spc-ttip-show-date") as HTMLSelectElement | null;
  const labelDateInput = document.getElementById("spc-ttip-label-date") as HTMLInputElement | null;
  const showNumeratorSel = document.getElementById("spc-ttip-show-numerator") as HTMLSelectElement | null;
  const labelNumeratorInput = document.getElementById("spc-ttip-label-numerator") as HTMLInputElement | null;
  const showDenominatorSel = document.getElementById("spc-ttip-show-denominator") as HTMLSelectElement | null;
  const labelDenominatorInput = document.getElementById("spc-ttip-label-denominator") as HTMLInputElement | null;
  const showValueSel = document.getElementById("spc-ttip-show-value") as HTMLSelectElement | null;
  const labelValueInput = document.getElementById("spc-ttip-label-value") as HTMLInputElement | null;
  const llTruncateInput = document.getElementById("spc-ll-truncate") as HTMLInputElement | null;
  const ulTruncateInput = document.getElementById("spc-ul-truncate") as HTMLInputElement | null;

  if (!spcInputSettings?.spc) {
    return;
  }

  if (chartTypeSel) {
    spcInputSettings.spc.chart_type = chartTypeSel.value as any;
  }
  if (outliersInLimitsSel) {
    spcInputSettings.spc.outliers_in_limits = parseBoolean(outliersInLimitsSel.value, spcInputSettings.spc.outliers_in_limits);
  }
  if (multiplierInput) {
    spcInputSettings.spc.multiplier = parseNumber(multiplierInput.value, spcInputSettings.spc.multiplier, { min: 0 });
  }
  if (sigFigsInput) {
    spcInputSettings.spc.sig_figs = parseNumber(sigFigsInput.value, spcInputSettings.spc.sig_figs, { min: 0, max: 20 });
  }
  if (percLabelsSel) {
    spcInputSettings.spc.perc_labels = percLabelsSel.value as any;
  }
  if (splitOnClickSel) {
    spcInputSettings.spc.split_on_click = parseBoolean(splitOnClickSel.value, spcInputSettings.spc.split_on_click);
  }
  if (numPointsSubsetInput) {
    spcInputSettings.spc.num_points_subset = parseOptionalNumber(numPointsSubsetInput.value, { min: 1 }) as any;
  }
  if (subsetPointsFromSel) {
    spcInputSettings.spc.subset_points_from = subsetPointsFromSel.value as any;
  }
  if (showDateSel) {
    spcInputSettings.spc.ttip_show_date = parseBoolean(showDateSel.value, spcInputSettings.spc.ttip_show_date);
  }
  if (labelDateInput) {
    const next = labelDateInput.value.trim();
    if (next.length) spcInputSettings.spc.ttip_label_date = next as any;
  }
  if (showNumeratorSel) {
    spcInputSettings.spc.ttip_show_numerator = parseBoolean(showNumeratorSel.value, spcInputSettings.spc.ttip_show_numerator);
  }
  if (labelNumeratorInput) {
    const next = labelNumeratorInput.value.trim();
    if (next.length) spcInputSettings.spc.ttip_label_numerator = next as any;
  }
  if (showDenominatorSel) {
    spcInputSettings.spc.ttip_show_denominator = parseBoolean(showDenominatorSel.value, spcInputSettings.spc.ttip_show_denominator);
  }
  if (labelDenominatorInput) {
    const next = labelDenominatorInput.value.trim();
    if (next.length) spcInputSettings.spc.ttip_label_denominator = next as any;
  }
  if (showValueSel) {
    spcInputSettings.spc.ttip_show_value = parseBoolean(showValueSel.value, spcInputSettings.spc.ttip_show_value);
  }
  if (labelValueInput) {
    const next = labelValueInput.value.trim();
    if (next.length) spcInputSettings.spc.ttip_label_value = next as any;
  }
  if (llTruncateInput) {
    spcInputSettings.spc.ll_truncate = parseOptionalNumber(llTruncateInput.value) as any;
  }
  if (ulTruncateInput) {
    spcInputSettings.spc.ul_truncate = parseOptionalNumber(ulTruncateInput.value) as any;
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Change the display of sideload message so it is hidden
    document.getElementById("sideload-msg").style.display = "none";
    // Show the app body. This is the main form
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("create-plot").onclick = () => tryCatch(createPlot);
    document.getElementById("preview-plot").onclick = () => tryCatch(previewPlot);

    // Move our rendering containers inside the preview area
    const previewHost = document.getElementById("preview-container");
    if (previewHost) {
      previewHost.appendChild(spcDiv);
      previewHost.appendChild(funnelDiv);
      // Ensure containers expand to fit preview area
      (spcDiv as HTMLElement).style.width = '100%';
      (spcDiv as HTMLElement).style.height = '100%';
      (funnelDiv as HTMLElement).style.width = '100%';
      (funnelDiv as HTMLElement).style.height = '100%';
    }

  // Render the Data Settings UI programmatically (reduces hard-coded HTML)
  renderSpcDataSettings();
  // Populate worksheet selector when dropdown is clicked; tables/columns depend on worksheet
  document.getElementById("worksheet-selector").onclick = () => { tryCatch(updateWorksheetSelector); };
  // Populate table selector when dropdown is clicked; columns update after table change
  document.getElementById("table-selector").onclick = () => {
    tryCatch(async () => {
      await updateTableSelector();
      await updateColumnSelectors();
    });
  };

  // React to field changes to control button availability
  const worksheetSel = document.getElementById("worksheet-selector") as HTMLSelectElement;
  const tableSel = document.getElementById("table-selector") as HTMLSelectElement;
  const catSel = document.getElementById("category-selector") as HTMLSelectElement;
  const numSel = document.getElementById("numerator-selector") as HTMLSelectElement;
  const denSel = document.getElementById("denominator-selector") as HTMLSelectElement;
  const sdSel = document.getElementById("sd-selector") as HTMLSelectElement | null;

  worksheetSel.onchange = () => {
    tryCatch(async () => {
      await updateTableSelector();
      // Only try to populate columns if a table is actually selected
      const nextTable = (document.getElementById("table-selector") as HTMLSelectElement | null)?.value;
      if (nextTable) {
        await updateColumnSelectors();
      } else {
        clearColumnSelectors();
      }
      updateActionButtonsEnabledState();
    });
  };
  tableSel.onchange = () => {
    tryCatch(async () => {
      const nextTable = (document.getElementById("table-selector") as HTMLSelectElement | null)?.value;
      if (nextTable) {
        await updateColumnSelectors();
      } else {
        clearColumnSelectors();
      }
      updateActionButtonsEnabledState();
    });
  };
  catSel.onchange = () => updateActionButtonsEnabledState();
  numSel.onchange = () => updateActionButtonsEnabledState();
  denSel.onchange = () => updateActionButtonsEnabledState();
  sdSel && (sdSel.onchange = () => updateActionButtonsEnabledState());
  
  // Tabs: Data/Inputs vs Settings
  const tabData = document.getElementById("tab-data") as HTMLButtonElement;
  const tabSettings = document.getElementById("tab-settings") as HTMLButtonElement;
  const panelData = document.getElementById("panel-data") as HTMLElement;
  const panelSettings = document.getElementById("panel-settings") as HTMLElement;
  const chartTypeHidden = document.getElementById("controlchart-selector") as HTMLInputElement;
  const toggleSpc = document.getElementById("toggle-spc") as HTMLButtonElement;
  const toggleFunnel = document.getElementById("toggle-funnel") as HTMLButtonElement;
  const chartTitleInput = document.getElementById("setting-chart-title") as HTMLInputElement;
  const chartTitleSizeInput = document.getElementById("setting-title-size") as HTMLInputElement;
  const chartTitleColorInput = document.getElementById("setting-title-color") as HTMLInputElement;

  function activateTab(which: 'data' | 'settings') {
    const isData = which === 'data';
    tabData.classList.toggle('tab--active', isData);
    tabSettings.classList.toggle('tab--active', !isData);
    tabData.setAttribute('aria-selected', String(isData));
    tabSettings.setAttribute('aria-selected', String(!isData));
    tabData.setAttribute('tabindex', isData ? '0' : '-1');
    tabSettings.setAttribute('tabindex', !isData ? '0' : '-1');
    panelData.hidden = !isData;
    panelSettings.hidden = isData;
  }

  tabData?.addEventListener('click', () => activateTab('data'));
  tabSettings?.addEventListener('click', () => activateTab('settings'));
  // Keyboard support: left/right to switch
  [tabData, tabSettings].forEach(tab => {
    tab?.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
        const isData = document.activeElement === tabData;
        if (e.key === 'ArrowRight') {
          (isData ? tabSettings : tabData).focus();
          activateTab(isData ? 'settings' : 'data');
        } else {
          (isData ? tabSettings : tabData).focus();
          activateTab(isData ? 'settings' : 'data');
        }
        e.preventDefault();
      }
    });
  });
  // Default to Data tab
  activateTab('data');

  function setChartType(value: 'spc' | 'funnel') {
    chartTypeHidden.value = value;
    const isSpc = value === 'spc';
    toggleSpc.classList.toggle('is-active', isSpc);
    toggleFunnel.classList.toggle('is-active', !isSpc);
    toggleSpc.setAttribute('aria-pressed', String(isSpc));
    toggleFunnel.setAttribute('aria-pressed', String(!isSpc));
    updateSdSelectorVisibility();
    updateDenominatorSelectorVisibility();
    updateActionButtonsEnabledState();
  }
  toggleSpc?.addEventListener('click', () => setChartType('spc'));
  toggleFunnel?.addEventListener('click', () => setChartType('funnel'));
  // Initialize hidden value
  setChartType('spc');

  // Live preview update on title input (debounced)
  let titleDebounce: number | undefined;
  function queuePreviewRefresh() {
    if (titleDebounce) { clearTimeout(titleDebounce); }
    titleDebounce = window.setTimeout(() => {
      // Only update if preview is active (buttons enabled and maybe already rendered)
      const previewEnabled = !document.getElementById('preview-plot')?.hasAttribute('disabled');
      if (previewEnabled) { tryCatch(previewPlot); }
    }, 250);
  }
  chartTitleInput?.addEventListener('input', queuePreviewRefresh);
  chartTitleSizeInput?.addEventListener('input', queuePreviewRefresh);
  chartTitleColorInput?.addEventListener('input', queuePreviewRefresh);

  // Live preview update on Data Settings controls
  const dataSettingsIds = [
    "spc-chart-type",
    "spc-outliers-in-limits",
    "spc-multiplier",
    "spc-sig-figs",
    "spc-perc-labels",
    "spc-split-on-click",
    "spc-num-points-subset",
    "spc-subset-points-from",
    "spc-ttip-show-date",
    "spc-ttip-label-date",
    "spc-ttip-show-numerator",
    "spc-ttip-label-numerator",
    "spc-ttip-show-denominator",
    "spc-ttip-label-denominator",
    "spc-ttip-show-value",
    "spc-ttip-label-value",
    "spc-ll-truncate",
    "spc-ul-truncate"
  ];
  dataSettingsIds.forEach((id) => {
    const el = document.getElementById(id) as (HTMLInputElement | HTMLSelectElement | null);
    el?.addEventListener('input', queuePreviewRefresh);
    el?.addEventListener('change', queuePreviewRefresh);
  });

  // Ensure SD selector visibility/required-state tracks chart type
  const spcChartTypeSel = document.getElementById("spc-chart-type") as HTMLSelectElement | null;
  spcChartTypeSel?.addEventListener("change", () => {
    updateSdSelectorVisibility();
    updateDenominatorSelectorVisibility();
    updateActionButtonsEnabledState();
  });
  updateSdSelectorVisibility();
  updateDenominatorSelectorVisibility();
    // Initial population of worksheet selector, then tables/columns
    tryCatch(async () => {
      await updateWorksheetSelector();
      await updateTableSelector();
      const nextTable = (document.getElementById("table-selector") as HTMLSelectElement | null)?.value;
      if (nextTable) {
        await updateColumnSelectors();
      } else {
        clearColumnSelectors();
      }
      updateActionButtonsEnabledState();
    });
    updateActionButtonsEnabledState();
  }
});

function clearColumnSelectors() {
  const categorySelector = document.getElementById("category-selector") as HTMLSelectElement | null;
  const numeratorSelector = document.getElementById("numerator-selector") as HTMLSelectElement | null;
  const denominatorSelector = document.getElementById("denominator-selector") as HTMLSelectElement | null;
  const sdSelector = document.getElementById("sd-selector") as HTMLSelectElement | null;

  if (categorySelector) categorySelector.innerHTML = '<option value="" disabled selected>Select category</option>';
  if (numeratorSelector) numeratorSelector.innerHTML = '<option value="" disabled selected>Select numerator</option>';
  if (denominatorSelector) denominatorSelector.innerHTML = '<option value="" disabled selected>Select denominator</option>';
  if (sdSelector) sdSelector.innerHTML = '<option value="" disabled selected>Select SD (Xbar)</option>';
  updateSdSelectorVisibility();
  updateDenominatorSelectorVisibility();
}

function fromExcelDate(excelDate: number): Date {
  return new Date((excelDate - (25567 + 2)) * 86400 * 1000);
}

async function updateTableSelector() {
  await Excel.run(async (context) => {
    const selectedWorksheetName = (document.getElementById("worksheet-selector") as HTMLSelectElement | null)?.value;
    if (!selectedWorksheetName) {
      throw new Error("No worksheet selected");
    }
    const worksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const tables = worksheet.tables.load("items/name");
    await context.sync();
    const tableSelector = document.getElementById("table-selector") as HTMLSelectElement;
    tableSelector.innerHTML = '<option value="" disabled selected>Select a table</option>';
    tables.items.forEach(table => {
      const option = document.createElement("option");
      option.value = table.name;
      option.text = table.name;
      tableSelector.appendChild(option);
    });
    if (tables.items.length > 0) {
      tableSelector.value = tables.items[0].name;
      // Automatically populate columns for the first table, but keep actions disabled
      tryCatch(updateColumnSelectors);
    } else {
      clearColumnSelectors();
    }
    updateActionButtonsEnabledState();
  });
}

async function updateWorksheetSelector() {
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets.load("items/name");
    const activeWorksheet = context.workbook.worksheets.getActiveWorksheet().load("name");
    await context.sync();

    const worksheetSelector = document.getElementById("worksheet-selector") as HTMLSelectElement;
    worksheetSelector.innerHTML = '<option value="" disabled selected>Select a worksheet</option>';

    worksheets.items.forEach(ws => {
      const option = document.createElement("option");
      option.value = ws.name;
      option.text = ws.name;
      worksheetSelector.appendChild(option);
    });

    // Default to active worksheet if present
    const activeName = activeWorksheet.name;
    const activeExists = worksheets.items.some(ws => ws.name === activeName);
    if (activeExists) {
      worksheetSelector.value = activeName;
    } else if (worksheets.items.length > 0) {
      worksheetSelector.value = worksheets.items[0].name;
    }
  });
}

async function updateColumnSelectors() {
  await Excel.run(async (context) => {
    const selectedWorksheetName = (document.getElementById("worksheet-selector") as HTMLSelectElement | null)?.value;
    if (!selectedWorksheetName) {
      throw new Error("No worksheet selected");
    }
    const worksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement).value;
    if (!selectedTableName) {
      clearColumnSelectors();
      updateActionButtonsEnabledState();
      return;
    }
    const table = worksheet.tables.getItem(selectedTableName);
    const columns = table.columns.load("items/name");
    await context.sync();
    const categorySelector = document.getElementById("category-selector") as HTMLSelectElement;
    const numeratorSelector = document.getElementById("numerator-selector") as HTMLSelectElement;
    const denominatorSelector = document.getElementById("denominator-selector") as HTMLSelectElement;
    const sdSelector = document.getElementById("sd-selector") as HTMLSelectElement | null;
    categorySelector.innerHTML = '<option value="" disabled selected>Select category</option>';
    numeratorSelector.innerHTML = '<option value="" disabled selected>Select numerator</option>';
    denominatorSelector.innerHTML = '<option value="" disabled selected>Select denominator</option>';
    if (sdSelector) {
      sdSelector.innerHTML = '<option value="" disabled selected>Select SD (Xbar)</option>';
    }
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

      if (sdSelector) {
        const option4 = document.createElement("option");
        option4.value = column.name;
        option4.text = column.name;
        sdSelector.appendChild(option4);
      }
    });
    // Columns reset, so ensure buttons reflect incomplete selection
    updateSdSelectorVisibility();
    updateActionButtonsEnabledState();
  });
}

function updateActionButtonsEnabledState() {
  const chartFamily = (document.getElementById("controlchart-selector") as HTMLInputElement | null)?.value;
  const isSpc = chartFamily === "spc";
  const isXbar = isSpc && getSelectedSpcChartType() === "xbar";
  const requiredIds = ["worksheet-selector", "table-selector", "category-selector", "numerator-selector"];
  if (isDenominatorRequired()) {
    requiredIds.push("denominator-selector");
  }
  if (isXbar) {
    requiredIds.push("sd-selector");
  }
  const allSelected = requiredIds.every((id) => {
    const el = document.getElementById(id) as HTMLSelectElement;
    return el && typeof el.value === 'string' && el.value.length > 0;
  });
  const createBtn = document.getElementById("create-plot");
  const previewBtn = document.getElementById("preview-plot");
  if (allSelected) {
    createBtn?.removeAttribute("disabled");
    previewBtn?.removeAttribute("disabled");
  } else {
    createBtn?.setAttribute("disabled", "true");
    previewBtn?.setAttribute("disabled", "true");
  }
}

async function createPlot() {
  await Excel.run(async (context) => {
    const selectedWorksheetName = (document.getElementById("worksheet-selector") as HTMLSelectElement | null)?.value;
    if (!selectedWorksheetName) {
      throw new Error("No worksheet selected");
    }
    const currentWorksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement).value;
    if (!selectedTableName) {
      throw new Error("No table selected");
    }
    const table = currentWorksheet.tables.getItem(selectedTableName);
    const selectedCategoryColumn = (document.getElementById("category-selector") as HTMLSelectElement).value;
    const selectedNumeratorColumn = (document.getElementById("numerator-selector") as HTMLSelectElement).value;
    const selectedDenominatorColumn = (document.getElementById("denominator-selector") as HTMLSelectElement).value;
    const selectedSdColumn = (document.getElementById("sd-selector") as HTMLSelectElement | null)?.value;

    const categoryColumn = table.columns.getItem(selectedCategoryColumn).getDataBodyRange().load("values");
    const numeratorsColumn = table.columns.getItem(selectedNumeratorColumn).getDataBodyRange().load("values");
    const controlChartType = (document.getElementById("controlchart-selector") as HTMLInputElement).value;
    if (controlChartType === "spc") {
      updateSpcInputSettingsFromUi();
    }

    const denomRequired = isDenominatorRequired();
    if (denomRequired && !selectedDenominatorColumn) {
      throw new Error("This chart type requires a Denominator column. Please select a Denominator under Data / Inputs.");
    }
    const denominatorsColumn = (denomRequired && selectedDenominatorColumn)
      ? table.columns.getItem(selectedDenominatorColumn).getDataBodyRange().load("values")
      : null;

    const needsXbarSd = controlChartType === "spc" && spcInputSettings.spc.chart_type === "xbar";
    if (needsXbarSd && !selectedSdColumn) {
      throw new Error("Xbar requires an SD column. Please select an SD column (Xbar) under Data / Inputs.");
    }

    const sdColumnRange = needsXbarSd ? table.columns.getItem(selectedSdColumn!).getDataBodyRange().load("values") : null;
    await context.sync();
    if (controlChartType === "spc") {
      updateSpcInputSettingsFromUi();
    }

    const rawData = categoryColumn.values.flat().map((cat, i) => {
      const row: any = {
        categories: controlChartType === "spc" ? fromExcelDate(cat) : cat,
        numerators: numeratorsColumn.values.flat()[i],
      };
      if (denominatorsColumn) {
        row.denominators = denominatorsColumn.values.flat()[i];
      }
      if (needsXbarSd && sdColumnRange) {
        row.xbar_sds = (sdColumnRange.values.flat() as any[])[i];
      }
      return row;
    });

    var updateArgs = {
      dataViews: makeUpdateValues(rawData, controlChartType === "spc" ? spcInputSettings : funnelInputSettings, aggregations).dataViews,
      viewport: { width: 640, height: 480 },
      type: 2//,
      //headless: true,
      //frontend: true
    };

    var currVisual = controlChartType === "spc" ? spcVisual : funnelVisual;

    currVisual.update(updateArgs as any);
    currVisual.svg.selectAll('.chart-title').remove();
    currVisual.svg
      .append("rect")
      .attr("width", "100%")
      .attr("height", "100%")
      .attr("fill", "white")
      .lower();

  const titleTextCreate = (document.getElementById('setting-chart-title') as HTMLInputElement)?.value?.trim();
  const titleSizeRaw = (document.getElementById('setting-title-size') as HTMLInputElement)?.value || '16';
  const titleSize = Math.min(48, Math.max(10, parseInt(titleSizeRaw, 10) || 16));
  const titleColor = (document.getElementById('setting-title-color') as HTMLInputElement)?.value || '#111111';

    if (titleTextCreate) {
      const plotProps: any = currVisual.viewModel?.plotProperties || {};
      const topPad = plotProps.yAxis?.end_padding ?? 0;
      const descender = 0.2 * titleSize; // rough descender height
      const gap = 2; // small gap above plot area
      const computedY = topPad > 0 ? (topPad - descender - gap) : (titleSize + 4);
      const titleY = Math.max(titleSize + 2, computedY);
      currVisual.svg
        .append('text')
        .attr('class', 'chart-title')
        .attr('x', plotProps.xAxis?.start_padding || 20)
        .attr('y', titleY)
        .attr('font-size', titleSize)
        .attr('fill', titleColor)
        .text(titleTextCreate);
    }

    var image = currentWorksheet.shapes.addImage(btoa(currVisual.svg.node().outerHTML));
    image.name = "Image";
    image.top = 10;
    image.left = 200;

    await context.sync();
  });
}

async function previewPlot() {
  // Render the chart in the side-pane preview area without inserting into main Excel area
  const selectedWorksheetName = (document.getElementById("worksheet-selector") as HTMLSelectElement | null)?.value;
  const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement).value;
  const selectedCategoryColumn = (document.getElementById("category-selector") as HTMLSelectElement).value;
  const selectedNumeratorColumn = (document.getElementById("numerator-selector") as HTMLSelectElement).value;
  const selectedDenominatorColumn = (document.getElementById("denominator-selector") as HTMLSelectElement).value;
  const selectedSdColumn = (document.getElementById("sd-selector") as HTMLSelectElement | null)?.value;

  if (!selectedWorksheetName || !selectedTableName || !selectedCategoryColumn || !selectedNumeratorColumn) {
    throw new Error("Please select a worksheet, table, category and numerator to preview.");
  }

  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const table = currentWorksheet.tables.getItem(selectedTableName);

    const categoryColumn = table.columns.getItem(selectedCategoryColumn).getDataBodyRange().load("values");
    const numeratorsColumn = table.columns.getItem(selectedNumeratorColumn).getDataBodyRange().load("values");

    const controlChartType = (document.getElementById("controlchart-selector") as HTMLInputElement).value;
    if (controlChartType === "spc") {
      updateSpcInputSettingsFromUi();
    }

    const denomRequired = isDenominatorRequired();
    if (denomRequired && !selectedDenominatorColumn) {
      throw new Error("This chart type requires a Denominator column. Please select a Denominator under Data / Inputs.");
    }
    const denominatorsColumn = (denomRequired && selectedDenominatorColumn)
      ? table.columns.getItem(selectedDenominatorColumn).getDataBodyRange().load("values")
      : null;
    const needsXbarSd = controlChartType === "spc" && spcInputSettings.spc.chart_type === "xbar";
    if (needsXbarSd && !selectedSdColumn) {
      throw new Error("Xbar requires an SD column. Please select an SD column (Xbar) under Data / Inputs.");
    }
    const sdColumnRange = needsXbarSd ? table.columns.getItem(selectedSdColumn!).getDataBodyRange().load("values") : null;

    await context.sync();
    const rawData = categoryColumn.values.flat().map((cat, i) => {
      const row: any = {
        categories: controlChartType === "spc" ? fromExcelDate(cat) : cat,
        numerators: numeratorsColumn.values.flat()[i],
      };
      if (denominatorsColumn) {
        row.denominators = denominatorsColumn.values.flat()[i];
      }
      if (needsXbarSd && sdColumnRange) {
        row.xbar_sds = (sdColumnRange.values.flat() as any[])[i];
      }
      return row;
    });

    const previewHost = document.getElementById("preview-container");
    const containerRect = previewHost.getBoundingClientRect();
    const padding = 8 * 2; // preview container padding
    const width = Math.max(320, Math.floor(containerRect.width - padding));
    const height = Math.max(220, Math.floor(containerRect.height - padding));

    const updateArgs = {
      dataViews: makeUpdateValues(rawData, controlChartType === "spc" ? spcInputSettings : funnelInputSettings, aggregations).dataViews,
      viewport: { width, height },
      type: 2
    } as any;

    const currDiv = controlChartType === "spc" ? spcDiv : funnelDiv;
    const otherDiv = controlChartType === "spc" ? funnelDiv : spcDiv;
    currDiv.removeAttribute("hidden");
    otherDiv.setAttribute("hidden", "true");

    const currVisual = controlChartType === "spc" ? spcVisual : funnelVisual;
    currVisual.update(updateArgs);
    // Remove any mouse handlers that power the tooltip on the root svg (defense in depth)
    (currVisual.svg as any).on("mousemove", null).on("mouseleave", null);
    // Ensure a white background for clarity in dark themes
    currVisual.svg.selectAll('.chart-title').remove();
    currVisual.svg
      .append("rect")
      .attr("width", "100%")
      .attr("height", "100%")
      .attr("fill", "white")
      .lower();

  const titleTextPreview = (document.getElementById('setting-chart-title') as HTMLInputElement)?.value?.trim();
  const titleSizeRawPrev = (document.getElementById('setting-title-size') as HTMLInputElement)?.value || '16';
    const titleSizePrev = Math.min(48, Math.max(10, parseInt(titleSizeRawPrev, 10) || 16));
  const titleColorPrev = (document.getElementById('setting-title-color') as HTMLInputElement)?.value || '#111111';

    if (titleTextPreview) {
      const plotProps: any = currVisual.viewModel?.plotProperties || {};
      const topPad = plotProps.yAxis?.end_padding ?? 0;
      const descenderPrev = 0.2 * titleSizePrev;
      const gapPrev = 2;
      const computedYPrev = topPad > 0 ? (topPad - descenderPrev - gapPrev) : (titleSizePrev + 4);
      const titleYPrev = Math.max(titleSizePrev + 2, computedYPrev);
      currVisual.svg
        .append('text')
        .attr('class', 'chart-title')
        .attr('x', plotProps.xAxis?.start_padding || 20)
        .attr('y', titleYPrev)
        .attr('font-size', titleSizePrev)
        .attr('fill', titleColorPrev)
        .text(titleTextPreview);
    }
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
