/// <reference types="@types/office-js" />

import { makeConstructorArgs, makeUpdateValues } from "../utilities/previewDataUtils";
import { renderSpcDataSettings } from "../utilities/renderSpcDataSettings";
import { patchSpcVisualSettingsForTaskpane } from "../utilities/spcVisualSettingsCompat";
import { attachFieldInfo } from "../utilities/uiFields";
import dateSettingsToFormatOptions from "../PowerBI-SPC/src/Functions/dateSettingsToFormatOptions";
import formatDateParts from "../PowerBI-SPC/src/Functions/formatDateParts";
import { Visual as spcVisualClass } from "../PowerBI-SPC/src/visual";
import { Visual as funnelVisualClass } from "../PowerBI-Funnels/src/visual";
import {
  defaultSettings as spcDefaultSettings,
  type settingsValueType as spcDefaultSettingsType,
} from "../PowerBI-SPC/src/settings";
import {
  defaultSettings as funnelDefaultSettings,
  type settingsValueType as funnelDefaultSettingsType,
} from "../PowerBI-Funnels/src/settings";

const spcDiv = document.createElement("div");
spcDiv.className = "spc-container";
spcDiv.setAttribute("hidden", "true");

const funnelDiv = document.createElement("div");
funnelDiv.className = "funnel-container";
funnelDiv.setAttribute("hidden", "true");

const spcVisual = new spcVisualClass(makeConstructorArgs(spcDiv));
patchSpcVisualSettingsForTaskpane(spcVisual);
const funnelVisual = new funnelVisualClass(makeConstructorArgs(funnelDiv));

const spcInputSettings = structuredClone(spcDefaultSettings) as spcDefaultSettingsType;
const spcBaseCanvasPadding = {
  left: spcInputSettings.canvas.left_padding + 50,
  lower: spcInputSettings.canvas.lower_padding + 50,
  upper: spcInputSettings.canvas.upper_padding,
  right: spcInputSettings.canvas.right_padding,
};
spcInputSettings.canvas.left_padding = spcBaseCanvasPadding.left;
spcInputSettings.canvas.lower_padding = spcBaseCanvasPadding.lower;

const funnelInputSettings = structuredClone(funnelDefaultSettings) as funnelDefaultSettingsType;
funnelInputSettings.labels.show_labels = false;
const funnelBaseCanvasPadding = {
  left: funnelInputSettings.canvas.left_padding + 50,
  lower: funnelInputSettings.canvas.lower_padding + 25,
  upper: funnelInputSettings.canvas.upper_padding,
  right: funnelInputSettings.canvas.right_padding,
};
funnelInputSettings.canvas.left_padding = funnelBaseCanvasPadding.left;
funnelInputSettings.canvas.lower_padding = funnelBaseCanvasPadding.lower;

const aggregations: Record<string, string> = {
  numerators: "sum",
  denominators: "sum",
  xbar_sds: "first",
};

function getChartAggregations(chartFamily: ChartFamily): Record<string, string> {
  if (chartFamily === "funnel" && funnelInputSettings.labels.show_labels) {
    return {
      ...aggregations,
      labels: "first",
    };
  }
  return aggregations;
}

type RawDataRow = {
  categories: string | Date | null;
  numerators: number;
  denominators?: number | undefined;
  xbar_sds?: number | undefined;
  labels?: string | undefined;
};

type ThemeMode = "light" | "dark";
type ChartFamily = "spc" | "funnel";

type SpcDateRangeFilterState = {
  availableDates: Date[];
  startIndex: number;
  endIndex: number;
  enabled: boolean;
};

const themeStorageKey = "excel-control-charts-theme";
const spcDateRangeFilterState: SpcDateRangeFilterState = {
  availableDates: [],
  startIndex: 0,
  endIndex: 0,
  enabled: false,
};

function attachStaticFieldHelpText() {
  const helpItems: Array<[selector: string, helpText: string]> = [
    ["#charttype-label", "Choose whether to build an SPC control chart or a funnel plot."],
    [
      'label[for="worksheet-selector"]',
      "Choose the worksheet that contains the source table for the chart.",
    ],
    ['label[for="table-selector"]', "Choose the Excel table that supplies the chart data."],
    [
      'label[for="category-selector"]',
      "Choose the column used for dates, categories, or x-axis group labels.",
    ],
    [
      'label[for="numerator-selector"]',
      "Choose the main measure that will be plotted on the chart.",
    ],
    [
      'label[for="denominator-selector"]',
      "Choose the denominator or population column when the selected chart type requires one.",
    ],
    ['label[for="sd-selector"]', "Choose the standard deviation column required for Xbar charts."],
    [
      "#spc-date-range-filter-field > label",
      "For SPC date categories, limit the preview and created chart to the selected date window.",
    ],
    [
      ".date-range-filter__row:first-of-type label",
      "Choose the first date to include in the SPC chart range.",
    ],
    [
      ".date-range-filter__row:last-of-type label",
      "Choose the last date to include in the SPC chart range.",
    ],
    ['label[for="setting-chart-title"]', "Optional title shown above the chart."],
    ['label[for="setting-title-size"]', "Set the chart title size in points."],
    ['label[for="setting-title-color"]', "Choose the color used for the chart title text."],
    [
      'label[for="setting-show-date-range"]',
      "Show or hide the formatted SPC date range beneath the chart title.",
    ],
  ];

  helpItems.forEach(([selector, helpText]) => {
    const label = document.querySelector(selector) as HTMLElement | null;
    if (label) {
      attachFieldInfo(label, helpText);
    }
  });
}

function getStoredTheme(): ThemeMode | null {
  try {
    const value = window.localStorage.getItem(themeStorageKey);
    return value === "light" || value === "dark" ? value : null;
  } catch {
    return null;
  }
}

function getPreferredTheme(): ThemeMode {
  const storedTheme = getStoredTheme();
  if (storedTheme) return storedTheme;
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function applyTheme(theme: ThemeMode) {
  const root = document.body;
  const toggle = document.getElementById("theme-toggle") as HTMLButtonElement | null;
  const label = document.getElementById("theme-toggle-label") as HTMLElement | null;
  const isDark = theme === "dark";

  root.dataset.theme = theme;
  toggle?.setAttribute("aria-pressed", String(isDark));
  toggle?.setAttribute("aria-label", isDark ? "Switch to light mode" : "Switch to dark mode");

  if (label) {
    label.textContent = isDark ? "Light mode" : "Dark mode";
  }

  try {
    window.localStorage.setItem(themeStorageKey, theme);
  } catch {
    // Ignore storage failures and keep the in-memory theme.
  }
}

function toggleTheme() {
  applyTheme(document.body.dataset.theme === "dark" ? "light" : "dark");
}

function isValidDateValue(value: unknown): value is Date {
  return value instanceof Date && Number.isFinite(value.getTime());
}

function getChartTitleText(): string {
  return (
    (document.getElementById("setting-chart-title") as HTMLInputElement | null)?.value || ""
  ).trim();
}

function getChartTitleSize(): number {
  const raw = (document.getElementById("setting-title-size") as HTMLInputElement | null)?.value;
  return Math.min(48, Math.max(10, parseInt(raw || "16", 10) || 16));
}

function getChartTitleColor(): string {
  return (
    (document.getElementById("setting-title-color") as HTMLInputElement | null)?.value || "#111111"
  );
}

function shouldShowDateRange(): boolean {
  const dateRangeSel = document.getElementById(
    "setting-show-date-range"
  ) as HTMLSelectElement | null;
  return parseBoolean(dateRangeSel?.value, true);
}

function updateDateRangeSettingVisibility(chartFamily: ChartFamily) {
  const dateRangeField = document
    .getElementById("setting-show-date-range")
    ?.closest(".field") as HTMLElement | null;
  if (!dateRangeField) {
    return;
  }

  const shouldShow = chartFamily === "spc";
  dateRangeField.hidden = !shouldShow;
  dateRangeField.style.display = shouldShow ? "" : "none";
}

function formatDateSliderLabel(date: Date): string {
  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "2-digit",
  });
}

function getSpcDateRangeFilterElements() {
  return {
    field: document.getElementById("spc-date-range-filter-field") as HTMLElement | null,
    summary: document.getElementById("spc-date-range-summary") as HTMLElement | null,
    startInput: document.getElementById("spc-date-range-start") as HTMLInputElement | null,
    endInput: document.getElementById("spc-date-range-end") as HTMLInputElement | null,
    startLabel: document.getElementById("spc-date-range-start-label") as HTMLOutputElement | null,
    endLabel: document.getElementById("spc-date-range-end-label") as HTMLOutputElement | null,
  };
}

function rawDataSupportsDateFiltering(rawData: RawDataRow[]): boolean {
  return rawData.length > 0 && rawData.every((row) => isValidDateValue(row.categories));
}

function getUniqueSortedDates(rawData: RawDataRow[]): Date[] {
  const uniqueDates = new Map<number, Date>();

  rawData.forEach((row) => {
    if (!isValidDateValue(row.categories)) {
      return;
    }

    const timestamp = row.categories.getTime();
    if (!uniqueDates.has(timestamp)) {
      uniqueDates.set(timestamp, row.categories);
    }
  });

  return Array.from(uniqueDates.values()).sort((left, right) => left.getTime() - right.getTime());
}

function updateSpcDateRangeFilterUi() {
  const { field, summary, startInput, endInput, startLabel, endLabel } =
    getSpcDateRangeFilterElements();
  if (!field || !summary || !startInput || !endInput || !startLabel || !endLabel) {
    return;
  }

  const shouldShow =
    getSelectedChartFamily() === "spc" &&
    spcDateRangeFilterState.enabled &&
    spcDateRangeFilterState.availableDates.length > 0;

  field.hidden = !shouldShow;
  field.style.display = shouldShow ? "" : "none";

  if (!shouldShow) {
    summary.textContent = "Select a date-based category column to enable filtering.";
    startInput.disabled = true;
    endInput.disabled = true;
    startInput.min = "0";
    startInput.max = "0";
    endInput.min = "0";
    endInput.max = "0";
    startInput.value = "0";
    endInput.value = "0";
    startLabel.value = "-";
    startLabel.textContent = "-";
    endLabel.value = "-";
    endLabel.textContent = "-";
    return;
  }

  const maxIndex = spcDateRangeFilterState.availableDates.length - 1;
  spcDateRangeFilterState.startIndex = Math.max(
    0,
    Math.min(spcDateRangeFilterState.startIndex, maxIndex)
  );
  spcDateRangeFilterState.endIndex = Math.max(
    spcDateRangeFilterState.startIndex,
    Math.min(spcDateRangeFilterState.endIndex, maxIndex)
  );

  startInput.disabled = maxIndex === 0;
  endInput.disabled = maxIndex === 0;
  startInput.min = "0";
  endInput.min = "0";
  startInput.max = String(maxIndex);
  endInput.max = String(maxIndex);
  startInput.value = String(spcDateRangeFilterState.startIndex);
  endInput.value = String(spcDateRangeFilterState.endIndex);

  const startDate = spcDateRangeFilterState.availableDates[spcDateRangeFilterState.startIndex];
  const endDate = spcDateRangeFilterState.availableDates[spcDateRangeFilterState.endIndex];
  const startText = formatDateSliderLabel(startDate);
  const endText = formatDateSliderLabel(endDate);
  const includedDates = spcDateRangeFilterState.endIndex - spcDateRangeFilterState.startIndex + 1;

  startLabel.value = startText;
  startLabel.textContent = startText;
  endLabel.value = endText;
  endLabel.textContent = endText;
  summary.textContent =
    includedDates === 1
      ? `Including ${startText}`
      : `Including ${startText} to ${endText} (${includedDates} dates)`;
}

function resetSpcDateRangeFilter() {
  spcDateRangeFilterState.availableDates = [];
  spcDateRangeFilterState.startIndex = 0;
  spcDateRangeFilterState.endIndex = 0;
  spcDateRangeFilterState.enabled = false;
  updateSpcDateRangeFilterUi();
}

function refreshSpcDateRangeFilterFromRawData(rawData: RawDataRow[]) {
  if (!rawDataSupportsDateFiltering(rawData)) {
    resetSpcDateRangeFilter();
    return;
  }

  const availableDates = getUniqueSortedDates(rawData);
  if (!availableDates.length) {
    resetSpcDateRangeFilter();
    return;
  }

  spcDateRangeFilterState.availableDates = availableDates;
  spcDateRangeFilterState.startIndex = 0;
  spcDateRangeFilterState.endIndex = availableDates.length - 1;
  spcDateRangeFilterState.enabled = true;
  updateSpcDateRangeFilterUi();
}

function applySpcDateRangeFilter(rawData: RawDataRow[]): RawDataRow[] {
  if (!spcDateRangeFilterState.enabled || !rawDataSupportsDateFiltering(rawData)) {
    return rawData;
  }

  const startDate = spcDateRangeFilterState.availableDates[spcDateRangeFilterState.startIndex];
  const endDate = spcDateRangeFilterState.availableDates[spcDateRangeFilterState.endIndex];
  if (!startDate || !endDate) {
    return rawData;
  }

  const startTime = startDate.getTime();
  const endTime = endDate.getTime();
  return rawData.filter(
    (row) =>
      isValidDateValue(row.categories) &&
      row.categories.getTime() >= startTime &&
      row.categories.getTime() <= endTime
  );
}

function fitTextToWidth(text: string, maxWidthPx: number, fontSizePx: number): string {
  const maxChars = Math.max(8, Math.floor(maxWidthPx / (fontSizePx * 0.56)));
  if (text.length <= maxChars) return text;
  return `${text.slice(0, Math.max(0, maxChars - 3)).trimEnd()}...`;
}

function formatDateForDisplay(date: Date): string {
  const dateSettings = spcInputSettings.dates;
  const formatOptions = dateSettingsToFormatOptions(dateSettings);
  const locale = dateSettings.date_format_locale as "en-GB" | "en-US";
  const dayElement = locale === "en-GB" ? "day" : "month";
  const monthElement = locale === "en-GB" ? "month" : "day";
  const datePartsRecord = formatDateParts(date, locale, formatOptions);
  const datePartStrings = [
    `${datePartsRecord.weekday} ${datePartsRecord[dayElement]}`.trim(),
    datePartsRecord[monthElement],
    datePartsRecord.year,
  ];

  return datePartStrings.filter((part) => String(part).trim()).join(dateSettings.date_format_delim);
}

function formatDateRange(rawData: RawDataRow[]): string | null {
  const dates = rawData
    .map((row) => row.categories)
    .filter(isValidDateValue)
    .slice()
    .sort((a, b) => a.getTime() - b.getTime());

  if (!dates.length) return null;

  const first = formatDateForDisplay(dates[0]);
  const last = formatDateForDisplay(dates[dates.length - 1]);
  return first === last ? first : `${first} to ${last}`;
}

function rawDataSupportsDateFormatting(rawData: RawDataRow[]): boolean {
  const categories = rawData.map((row) => row.categories);
  return (
    categories.some(isValidDateValue) &&
    categories.every((category) => category === null || isValidDateValue(category))
  );
}

function updateHeaderCanvasPadding(controlChartType: string, includeDateRange: boolean) {
  const titleText = getChartTitleText();
  const titleSize = getChartTitleSize();
  const includesDateRange = controlChartType === "spc" && includeDateRange && shouldShowDateRange();
  let headerPadding = 0;

  if (titleText) headerPadding += titleSize + 8;
  if (includesDateRange) headerPadding += 17;
  if (headerPadding > 0) headerPadding += 8;

  if (controlChartType === "spc") {
    spcInputSettings.canvas.upper_padding = spcBaseCanvasPadding.upper + headerPadding;
  } else {
    funnelInputSettings.canvas.upper_padding =
      funnelBaseCanvasPadding.upper + (titleText ? titleSize + 16 : 0);
  }
}

function applyFunnelCategoryLabelStyle() {
  funnelInputSettings.scatter.use_group_text = false;

  if (!funnelInputSettings.labels.show_labels) {
    return;
  }

  funnelInputSettings.labels.label_position = "top" as any;
  funnelInputSettings.labels.label_y_offset = 0;
  funnelInputSettings.labels.label_line_offset = 0;
  funnelInputSettings.labels.label_angle_offset = 90;
  funnelInputSettings.labels.label_line_width = 0;
  funnelInputSettings.labels.label_line_colour = "transparent";
  funnelInputSettings.labels.label_line_max_length = 28;
  funnelInputSettings.labels.label_marker_show = false;
  funnelInputSettings.labels.label_marker_offset = 0;
  funnelInputSettings.labels.label_marker_size = 0;
  funnelInputSettings.labels.label_marker_colour = "transparent";
  funnelInputSettings.labels.label_marker_outline_colour = "transparent";
}

function drawChartFrameAndHeader(currVisual: any, rawData: RawDataRow[], controlChartType: string) {
  const svg = currVisual.svg;
  svg.selectAll(".chart-background").remove();
  svg.selectAll(".chart-title,.chart-subtitle").remove();
  svg
    .append("rect")
    .attr("class", "chart-background")
    .attr("width", "100%")
    .attr("height", "100%")
    .attr("fill", "white")
    .lower();

  const titleText = getChartTitleText();
  const titleSize = getChartTitleSize();
  const titleColor = getChartTitleColor();
  const dateRange =
    controlChartType === "spc" && shouldShowDateRange() ? formatDateRange(rawData) : null;

  if (!titleText && !dateRange) return;

  const plotProps: any = currVisual?.plotProperties || {};
  const svgWidth = Number(svg.attr("width")) || 640;
  const x = plotProps.xAxis?.start_padding || 20;
  const maxTextWidth = Math.max(80, svgWidth - x - 16);
  let nextY = 0;

  if (titleText) {
    nextY = titleSize + 4;
    svg
      .append("text")
      .attr("class", "chart-title")
      .attr("x", x)
      .attr("y", nextY)
      .attr("font-family", "Segoe UI, Arial, sans-serif")
      .attr("font-weight", "700")
      .attr("font-size", titleSize)
      .attr("fill", titleColor)
      .text(fitTextToWidth(titleText, maxTextWidth, titleSize));
  }

  if (dateRange) {
    const subtitleSize = 11;
    nextY = titleText ? nextY + subtitleSize + 6 : subtitleSize + 5;
    svg
      .append("text")
      .attr("class", "chart-subtitle")
      .attr("x", x)
      .attr("y", nextY)
      .attr("font-family", "Segoe UI, Arial, sans-serif")
      .attr("font-size", subtitleSize)
      .attr("fill", "#465169")
      .text(fitTextToWidth(`Date range: ${dateRange}`, maxTextWidth, subtitleSize));
  }
}

function getSelectedSpcChartType(): string {
  const el = document.getElementById("spc-chart-type") as HTMLSelectElement | null;
  return el?.value || "i";
}

function spcChartTypeHasControlLimits(chartType: string = getSelectedSpcChartType()): boolean {
  return chartType !== "run";
}

function canRenderSpcAssuranceIcons(
  altTarget: number | null,
  improvementDirection: string,
  chartType: string = getSelectedSpcChartType()
): boolean {
  return (
    altTarget !== null &&
    improvementDirection !== "neutral" &&
    spcChartTypeHasControlLimits(chartType)
  );
}

function getSelectedChartFamily(): ChartFamily {
  const value = (document.getElementById("controlchart-selector") as HTMLInputElement | null)
    ?.value;
  return value === "funnel" ? "funnel" : "spc";
}

async function updateSpcDateRangeFilterControls() {
  if (getSelectedChartFamily() !== "spc") {
    resetSpcDateRangeFilter();
    return;
  }

  const selectedWorksheetName = (
    document.getElementById("worksheet-selector") as HTMLSelectElement | null
  )?.value;
  const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement | null)
    ?.value;
  const selectedCategoryColumn = (
    document.getElementById("category-selector") as HTMLSelectElement | null
  )?.value;

  if (!selectedWorksheetName || !selectedTableName || !selectedCategoryColumn) {
    resetSpcDateRangeFilter();
    return;
  }

  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const table = worksheet.tables.getItem(selectedTableName);
    const categoryColumn = table.columns
      .getItem(selectedCategoryColumn)
      .getDataBodyRange()
      .load("values");
    await context.sync();

    const rawData: RawDataRow[] = categoryColumn.values.flat().map((value) => ({
      categories: fromExcelDate(value),
      numerators: 0,
    }));
    refreshSpcDateRangeFilterFromRawData(rawData);
  });
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
  const chartFamily = (document.getElementById("controlchart-selector") as HTMLInputElement | null)
    ?.value;
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
  const chartFamily = (document.getElementById("controlchart-selector") as HTMLInputElement | null)
    ?.value;
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

function parseNumber(
  value: string | undefined | null,
  fallback: number,
  opts?: { min?: number; max?: number }
): number {
  const raw = (value ?? "").trim();
  const parsed = raw.length ? Number(raw) : NaN;
  let next = Number.isFinite(parsed) ? parsed : fallback;
  if (opts?.min !== undefined) next = Math.max(opts.min, next);
  if (opts?.max !== undefined) next = Math.min(opts.max, next);
  return next;
}

function parseOptionalNumber(
  value: string | undefined | null,
  opts?: { min?: number; max?: number }
): number | null {
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
  const outliersInLimitsSel = document.getElementById(
    "spc-outliers-in-limits"
  ) as HTMLSelectElement | null;
  const multiplierInput = document.getElementById("spc-multiplier") as HTMLInputElement | null;
  const sigFigsInput = document.getElementById("spc-sig-figs") as HTMLInputElement | null;
  const percLabelsSel = document.getElementById("spc-perc-labels") as HTMLSelectElement | null;
  const splitOnClickSel = document.getElementById("spc-split-on-click") as HTMLSelectElement | null;
  const numPointsSubsetInput = document.getElementById(
    "spc-num-points-subset"
  ) as HTMLInputElement | null;
  const subsetPointsFromSel = document.getElementById(
    "spc-subset-points-from"
  ) as HTMLSelectElement | null;
  const llTruncateInput = document.getElementById("spc-ll-truncate") as HTMLInputElement | null;
  const ulTruncateInput = document.getElementById("spc-ul-truncate") as HTMLInputElement | null;
  const showVariationSel = document.getElementById(
    "spc-show-variation-icons"
  ) as HTMLSelectElement | null;
  const flagLastPointSel = document.getElementById(
    "spc-flag-last-point"
  ) as HTMLSelectElement | null;
  const variationLocationSel = document.getElementById(
    "spc-variation-location"
  ) as HTMLSelectElement | null;
  const variationScalingInput = document.getElementById(
    "spc-variation-scaling"
  ) as HTMLInputElement | null;
  const showAssuranceSel = document.getElementById(
    "spc-show-assurance-icons"
  ) as HTMLSelectElement | null;
  const assuranceLocationSel = document.getElementById(
    "spc-assurance-location"
  ) as HTMLSelectElement | null;
  const assuranceScalingInput = document.getElementById(
    "spc-assurance-scaling"
  ) as HTMLInputElement | null;
  const altTargetInput = document.getElementById("spc-alt-target") as HTMLInputElement | null;
  const altTargetLabelSel = document.getElementById(
    "spc-alt-target-label"
  ) as HTMLSelectElement | null;
  const improvementDirectionSel = document.getElementById(
    "spc-improvement-direction"
  ) as HTMLSelectElement | null;
  const astronomicalPointsSel = document.getElementById(
    "spc-astronomical-points"
  ) as HTMLSelectElement | null;
  const astronomicalLimitSel = document.getElementById(
    "spc-astronomical-limit"
  ) as HTMLSelectElement | null;
  const trendPatternSel = document.getElementById("spc-trend-pattern") as HTMLSelectElement | null;
  const trendPointsInput = document.getElementById("spc-trend-points") as HTMLInputElement | null;
  const twoInThreeSel = document.getElementById("spc-two-in-three") as HTMLSelectElement | null;
  const twoInThreeHighlightSeriesSel = document.getElementById(
    "spc-two-in-three-highlight-series"
  ) as HTMLSelectElement | null;
  const twoInThreeLimitSel = document.getElementById(
    "spc-two-in-three-limit"
  ) as HTMLSelectElement | null;
  const shiftPatternSel = document.getElementById("spc-shift-pattern") as HTMLSelectElement | null;
  const shiftPointsInput = document.getElementById("spc-shift-points") as HTMLInputElement | null;
  const dateFormatDaySel = document.getElementById(
    "spc-date-format-day"
  ) as HTMLSelectElement | null;
  const dateFormatMonthSel = document.getElementById(
    "spc-date-format-month"
  ) as HTMLSelectElement | null;
  const dateFormatYearSel = document.getElementById(
    "spc-date-format-year"
  ) as HTMLSelectElement | null;
  const dateFormatDelimSel = document.getElementById(
    "spc-date-format-delim"
  ) as HTMLSelectElement | null;
  const dateFormatLocaleSel = document.getElementById(
    "spc-date-format-locale"
  ) as HTMLSelectElement | null;

  if (!spcInputSettings?.spc) {
    return;
  }

  if (chartTypeSel) {
    spcInputSettings.spc.chart_type = chartTypeSel.value as any;
  }
  if (outliersInLimitsSel) {
    spcInputSettings.spc.outliers_in_limits = parseBoolean(
      outliersInLimitsSel.value,
      spcInputSettings.spc.outliers_in_limits
    );
  }
  if (multiplierInput) {
    spcInputSettings.spc.multiplier = parseNumber(
      multiplierInput.value,
      spcInputSettings.spc.multiplier,
      { min: 0 }
    );
  }
  if (sigFigsInput) {
    spcInputSettings.spc.sig_figs = parseNumber(sigFigsInput.value, spcInputSettings.spc.sig_figs, {
      min: 0,
      max: 20,
    });
  }
  if (percLabelsSel) {
    spcInputSettings.spc.perc_labels = percLabelsSel.value as any;
  }
  if (splitOnClickSel) {
    spcInputSettings.spc.split_on_click = parseBoolean(
      splitOnClickSel.value,
      spcInputSettings.spc.split_on_click
    );
  }
  if (numPointsSubsetInput) {
    spcInputSettings.spc.num_points_subset = parseOptionalNumber(numPointsSubsetInput.value, {
      min: 1,
    }) as any;
  }
  if (subsetPointsFromSel) {
    spcInputSettings.spc.subset_points_from = subsetPointsFromSel.value as any;
  }
  if (llTruncateInput) {
    spcInputSettings.spc.ll_truncate = parseOptionalNumber(llTruncateInput.value) as any;
  }
  if (ulTruncateInput) {
    spcInputSettings.spc.ul_truncate = parseOptionalNumber(ulTruncateInput.value) as any;
  }
  if (improvementDirectionSel) {
    spcInputSettings.outliers.improvement_direction = improvementDirectionSel.value as any;
  }
  if (shiftPatternSel) {
    spcInputSettings.outliers.shift = parseBoolean(
      shiftPatternSel.value,
      spcInputSettings.outliers.shift
    );
  }
  if (shiftPointsInput) {
    spcInputSettings.outliers.shift_n = parseNumber(
      shiftPointsInput.value,
      spcInputSettings.outliers.shift_n,
      { min: 1 }
    );
  }
  if (showVariationSel) {
    spcInputSettings.nhs_icons.show_variation_icons = parseBoolean(
      showVariationSel.value,
      spcInputSettings.nhs_icons.show_variation_icons
    );
  }
  if (flagLastPointSel) {
    spcInputSettings.nhs_icons.flag_last_point = parseBoolean(
      flagLastPointSel.value,
      spcInputSettings.nhs_icons.flag_last_point
    );
  }
  if (variationLocationSel) {
    spcInputSettings.nhs_icons.variation_icons_locations = variationLocationSel.value as any;
  }
  if (variationScalingInput) {
    spcInputSettings.nhs_icons.variation_icons_scaling = parseNumber(
      variationScalingInput.value,
      spcInputSettings.nhs_icons.variation_icons_scaling,
      { min: 0 }
    );
  }
  if (assuranceLocationSel) {
    spcInputSettings.nhs_icons.assurance_icons_locations = assuranceLocationSel.value as any;
  }
  if (assuranceScalingInput) {
    spcInputSettings.nhs_icons.assurance_icons_scaling = parseNumber(
      assuranceScalingInput.value,
      spcInputSettings.nhs_icons.assurance_icons_scaling,
      { min: 0 }
    );
  }
  const altTarget = parseOptionalNumber(altTargetInput?.value);
  const showAltTargetLabel = altTargetLabelSel?.value === "target_equals";
  spcInputSettings.lines.alt_target = altTarget as any;
  spcInputSettings.lines.show_alt_target = altTarget !== null;
  spcInputSettings.lines.plot_label_show_alt_target = showAltTargetLabel && altTarget !== null;
  spcInputSettings.lines.plot_label_prefix_alt_target = showAltTargetLabel ? "Target = " : "";
  if (showAssuranceSel) {
    const currentImprovementDirection =
      improvementDirectionSel?.value || spcInputSettings.outliers.improvement_direction;
    const canShowAssuranceIcons = canRenderSpcAssuranceIcons(
      altTarget,
      currentImprovementDirection,
      getSelectedSpcChartType()
    );
    if (!canShowAssuranceIcons) {
      showAssuranceSel.value = "false";
    }
    spcInputSettings.nhs_icons.show_assurance_icons =
      parseBoolean(showAssuranceSel.value, spcInputSettings.nhs_icons.show_assurance_icons) &&
      canShowAssuranceIcons;
  }
  if (astronomicalPointsSel) {
    spcInputSettings.outliers.astronomical = parseBoolean(
      astronomicalPointsSel.value,
      spcInputSettings.outliers.astronomical
    );
  }
  if (astronomicalLimitSel) {
    spcInputSettings.outliers.astronomical_limit = astronomicalLimitSel.value as any;
  }
  if (trendPatternSel) {
    spcInputSettings.outliers.trend = parseBoolean(
      trendPatternSel.value,
      spcInputSettings.outliers.trend
    );
  }
  if (trendPointsInput) {
    spcInputSettings.outliers.trend_n = parseNumber(
      trendPointsInput.value,
      spcInputSettings.outliers.trend_n,
      { min: 1 }
    );
  }
  if (twoInThreeSel) {
    spcInputSettings.outliers.two_in_three = parseBoolean(
      twoInThreeSel.value,
      spcInputSettings.outliers.two_in_three
    );
  }
  if (twoInThreeHighlightSeriesSel) {
    spcInputSettings.outliers.two_in_three_highlight_series = parseBoolean(
      twoInThreeHighlightSeriesSel.value,
      spcInputSettings.outliers.two_in_three_highlight_series
    );
  }
  if (twoInThreeLimitSel) {
    spcInputSettings.outliers.two_in_three_limit = twoInThreeLimitSel.value as any;
  }
  if (dateFormatDaySel) {
    spcInputSettings.dates.date_format_day = dateFormatDaySel.value as any;
  }
  if (dateFormatMonthSel) {
    spcInputSettings.dates.date_format_month = dateFormatMonthSel.value as any;
  }
  if (dateFormatYearSel) {
    spcInputSettings.dates.date_format_year = dateFormatYearSel.value as any;
  }
  if (dateFormatDelimSel) {
    spcInputSettings.dates.date_format_delim = dateFormatDelimSel.value as any;
  }
  if (dateFormatLocaleSel) {
    spcInputSettings.dates.date_format_locale = dateFormatLocaleSel.value as any;
  }
}

function updateFunnelInputSettingsFromUi() {
  const chartTypeSel = document.getElementById("spc-chart-type") as HTMLSelectElement | null;
  const odAdjustSel = document.getElementById("funnel-od-adjust") as HTMLSelectElement | null;
  const multiplierInput = document.getElementById("spc-multiplier") as HTMLInputElement | null;
  const sigFigsInput = document.getElementById("spc-sig-figs") as HTMLInputElement | null;
  const percLabelsSel = document.getElementById("spc-perc-labels") as HTMLSelectElement | null;
  const transformationSel = document.getElementById(
    "funnel-transformation"
  ) as HTMLSelectElement | null;
  const llTruncateInput = document.getElementById("spc-ll-truncate") as HTMLInputElement | null;
  const ulTruncateInput = document.getElementById("spc-ul-truncate") as HTMLInputElement | null;
  const processFlagTypeSel = document.getElementById(
    "funnel-process-flag-type"
  ) as HTMLSelectElement | null;
  const improvementDirectionSel = document.getElementById(
    "funnel-improvement-direction"
  ) as HTMLSelectElement | null;
  const categoryPointLabelsSel = document.getElementById(
    "funnel-category-point-labels"
  ) as HTMLSelectElement | null;
  const twoSigmaSel = document.getElementById("funnel-two-sigma") as HTMLSelectElement | null;
  const threeSigmaSel = document.getElementById("funnel-three-sigma") as HTMLSelectElement | null;

  if (!funnelInputSettings?.funnel) {
    return;
  }

  if (chartTypeSel) {
    funnelInputSettings.funnel.chart_type = chartTypeSel.value as any;
  }
  if (odAdjustSel) {
    funnelInputSettings.funnel.od_adjust = odAdjustSel.value as any;
  }
  if (multiplierInput) {
    funnelInputSettings.funnel.multiplier = parseNumber(
      multiplierInput.value,
      funnelInputSettings.funnel.multiplier,
      { min: 0 }
    );
  }
  if (sigFigsInput) {
    funnelInputSettings.funnel.sig_figs = parseNumber(
      sigFigsInput.value,
      funnelInputSettings.funnel.sig_figs,
      { min: 0, max: 20 }
    );
  }
  if (percLabelsSel) {
    funnelInputSettings.funnel.perc_labels = percLabelsSel.value as any;
  }
  if (transformationSel) {
    funnelInputSettings.funnel.transformation = transformationSel.value as any;
  }
  if (categoryPointLabelsSel) {
    funnelInputSettings.labels.show_labels = parseBoolean(
      categoryPointLabelsSel.value,
      funnelInputSettings.labels.show_labels
    );
  }
  if (llTruncateInput) {
    funnelInputSettings.funnel.ll_truncate = parseOptionalNumber(llTruncateInput.value) as any;
  }
  if (ulTruncateInput) {
    funnelInputSettings.funnel.ul_truncate = parseOptionalNumber(ulTruncateInput.value) as any;
  }
  if (processFlagTypeSel) {
    funnelInputSettings.outliers.process_flag_type = processFlagTypeSel.value as any;
  }
  if (improvementDirectionSel) {
    funnelInputSettings.outliers.improvement_direction = improvementDirectionSel.value as any;
  }
  if (twoSigmaSel) {
    funnelInputSettings.outliers.two_sigma = parseBoolean(
      twoSigmaSel.value,
      funnelInputSettings.outliers.two_sigma
    );
  }
  if (threeSigmaSel) {
    funnelInputSettings.outliers.three_sigma = parseBoolean(
      threeSigmaSel.value,
      funnelInputSettings.outliers.three_sigma
    );
  }

  applyFunnelCategoryLabelStyle();
}

function updateCurrentChartInputSettingsFromUi(
  chartFamily: ChartFamily = getSelectedChartFamily()
) {
  if (chartFamily === "spc") {
    updateSpcInputSettingsFromUi();
  } else {
    updateFunnelInputSettingsFromUi();
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    applyTheme(getPreferredTheme());
    attachStaticFieldHelpText();

    // Change the display of sideload message so it is hidden
    document.getElementById("sideload-msg")!.style.display = "none";
    // Show the app body. This is the main form
    document.getElementById("app-body")!.style.display = "flex";

    document.getElementById("create-plot")!.onclick = () => tryCatch(createPlot);
    document.getElementById("preview-plot")!.onclick = () => tryCatch(previewPlot);

    // Move our rendering containers inside the preview area
    const previewHost = document.getElementById("preview-container");
    if (previewHost) {
      previewHost.appendChild(spcDiv);
      previewHost.appendChild(funnelDiv);
      // Ensure containers expand to fit preview area
      (spcDiv as HTMLElement).style.width = "100%";
      (spcDiv as HTMLElement).style.height = "100%";
      (funnelDiv as HTMLElement).style.width = "100%";
      (funnelDiv as HTMLElement).style.height = "100%";
    }

    // Populate worksheet selector when dropdown is clicked; tables/columns depend on worksheet
    document.getElementById("worksheet-selector")!.onclick = () => {
      tryCatch(updateWorksheetSelector);
    };
    // Populate table selector when dropdown is clicked; columns update after table change
    document.getElementById("table-selector")!.onclick = () => {
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
        const nextTable = (document.getElementById("table-selector") as HTMLSelectElement | null)
          ?.value;
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
        const nextTable = (document.getElementById("table-selector") as HTMLSelectElement | null)
          ?.value;
        if (nextTable) {
          await updateColumnSelectors();
        } else {
          clearColumnSelectors();
        }
        updateActionButtonsEnabledState();
      });
    };
    catSel.onchange = () => {
      tryCatch(async () => {
        await updateSpcDateRangeFilterControls();
        updateActionButtonsEnabledState();
        queuePreviewRefresh();
      });
    };
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
    const themeToggle = document.getElementById("theme-toggle") as HTMLButtonElement | null;
    const chartTitleInput = document.getElementById("setting-chart-title") as HTMLInputElement;
    const chartTitleSizeInput = document.getElementById("setting-title-size") as HTMLInputElement;
    const chartTitleColorInput = document.getElementById("setting-title-color") as HTMLInputElement;
    const dateRangeStartInput = document.getElementById(
      "spc-date-range-start"
    ) as HTMLInputElement | null;
    const dateRangeEndInput = document.getElementById(
      "spc-date-range-end"
    ) as HTMLInputElement | null;
    const dynamicSettingsIds = [
      "spc-chart-type",
      "spc-outliers-in-limits",
      "spc-multiplier",
      "spc-sig-figs",
      "spc-perc-labels",
      "spc-split-on-click",
      "spc-num-points-subset",
      "spc-subset-points-from",
      "spc-ll-truncate",
      "spc-ul-truncate",
      "spc-show-variation-icons",
      "spc-flag-last-point",
      "spc-variation-location",
      "spc-variation-scaling",
      "spc-show-assurance-icons",
      "spc-assurance-location",
      "spc-assurance-scaling",
      "spc-alt-target",
      "spc-alt-target-label",
      "spc-improvement-direction",
      "spc-astronomical-points",
      "spc-astronomical-limit",
      "spc-trend-pattern",
      "spc-trend-points",
      "spc-two-in-three",
      "spc-two-in-three-highlight-series",
      "spc-two-in-three-limit",
      "spc-shift-pattern",
      "spc-shift-points",
      "spc-date-format-day",
      "spc-date-format-month",
      "spc-date-format-year",
      "spc-date-format-delim",
      "spc-date-format-locale",
      "funnel-od-adjust",
      "funnel-transformation",
      "funnel-category-point-labels",
      "funnel-process-flag-type",
      "funnel-improvement-direction",
      "funnel-two-sigma",
      "funnel-three-sigma",
    ];

    function activateTab(which: "data" | "settings") {
      const isData = which === "data";
      tabData.classList.toggle("tab--active", isData);
      tabSettings.classList.toggle("tab--active", !isData);
      tabData.setAttribute("aria-selected", String(isData));
      tabSettings.setAttribute("aria-selected", String(!isData));
      tabData.setAttribute("tabindex", isData ? "0" : "-1");
      tabSettings.setAttribute("tabindex", !isData ? "0" : "-1");
      panelData.hidden = !isData;
      panelSettings.hidden = isData;
    }

    tabData?.addEventListener("click", () => activateTab("data"));
    tabSettings?.addEventListener("click", () => activateTab("settings"));
    // Keyboard support: left/right to switch
    [tabData, tabSettings].forEach((tab) => {
      tab?.addEventListener("keydown", (e: KeyboardEvent) => {
        if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
          const isData = document.activeElement === tabData;
          if (e.key === "ArrowRight") {
            (isData ? tabSettings : tabData).focus();
            activateTab(isData ? "settings" : "data");
          } else {
            (isData ? tabSettings : tabData).focus();
            activateTab(isData ? "settings" : "data");
          }
          e.preventDefault();
        }
      });
    });
    // Default to Data tab
    activateTab("data");

    function bindRenderedSettingsControls() {
      dynamicSettingsIds.forEach((id) => {
        const el = document.getElementById(id) as HTMLInputElement | HTMLSelectElement | null;
        el?.addEventListener("input", queuePreviewRefresh);
        el?.addEventListener("change", queuePreviewRefresh);
      });

      const chartTypeSelect = document.getElementById("spc-chart-type") as HTMLSelectElement | null;
      chartTypeSelect?.addEventListener("change", () => {
        updateCurrentChartInputSettingsFromUi();
        updateSdSelectorVisibility();
        updateDenominatorSelectorVisibility();
        updateActionButtonsEnabledState();
      });
    }

    function setChartType(value: ChartFamily) {
      const currentChartFamily = getSelectedChartFamily();
      if (document.getElementById("spc-chart-type")) {
        updateCurrentChartInputSettingsFromUi(currentChartFamily);
      }

      chartTypeHidden.value = value;
      const isSpc = value === "spc";
      toggleSpc.classList.toggle("is-active", isSpc);
      toggleFunnel.classList.toggle("is-active", !isSpc);
      toggleSpc.setAttribute("aria-pressed", String(isSpc));
      toggleFunnel.setAttribute("aria-pressed", String(!isSpc));

      renderSpcDataSettings(value, spcInputSettings, funnelInputSettings);
      updateDateRangeSettingVisibility(value);
      bindRenderedSettingsControls();

      if (value === "spc") {
        tryCatch(updateSpcDateRangeFilterControls);
      } else {
        resetSpcDateRangeFilter();
      }

      updateSdSelectorVisibility();
      updateDenominatorSelectorVisibility();
      updateActionButtonsEnabledState();
      queuePreviewRefresh();
    }
    toggleSpc?.addEventListener("click", () => setChartType("spc"));
    toggleFunnel?.addEventListener("click", () => setChartType("funnel"));
    themeToggle?.addEventListener("click", toggleTheme);

    // Live preview update on title input (debounced)
    let titleDebounce: number | undefined;
    function queuePreviewRefresh() {
      if (titleDebounce) {
        clearTimeout(titleDebounce);
      }
      titleDebounce = window.setTimeout(() => {
        // Only update if preview is active (buttons enabled and maybe already rendered)
        const previewEnabled = !document.getElementById("preview-plot")?.hasAttribute("disabled");
        if (previewEnabled) {
          tryCatch(previewPlot);
        }
      }, 250);
    }
    chartTitleInput?.addEventListener("input", queuePreviewRefresh);
    chartTitleSizeInput?.addEventListener("input", queuePreviewRefresh);
    chartTitleColorInput?.addEventListener("input", queuePreviewRefresh);

    function handleDateRangeSliderInput(boundary: "start" | "end") {
      if (!spcDateRangeFilterState.enabled || !spcDateRangeFilterState.availableDates.length) {
        return;
      }

      const targetInput = boundary === "start" ? dateRangeStartInput : dateRangeEndInput;
      const rawValue = parseInt(targetInput?.value || "0", 10);
      const maxIndex = spcDateRangeFilterState.availableDates.length - 1;
      const nextIndex = Math.max(0, Math.min(Number.isFinite(rawValue) ? rawValue : 0, maxIndex));

      if (boundary === "start") {
        spcDateRangeFilterState.startIndex = Math.min(nextIndex, spcDateRangeFilterState.endIndex);
      } else {
        spcDateRangeFilterState.endIndex = Math.max(nextIndex, spcDateRangeFilterState.startIndex);
      }

      updateSpcDateRangeFilterUi();
      queuePreviewRefresh();
    }

    dateRangeStartInput?.addEventListener("input", () => handleDateRangeSliderInput("start"));
    dateRangeStartInput?.addEventListener("change", () => handleDateRangeSliderInput("start"));
    dateRangeEndInput?.addEventListener("input", () => handleDateRangeSliderInput("end"));
    dateRangeEndInput?.addEventListener("change", () => handleDateRangeSliderInput("end"));

    const dateRangeSel = document.getElementById(
      "setting-show-date-range"
    ) as HTMLSelectElement | null;
    dateRangeSel?.addEventListener("input", queuePreviewRefresh);
    dateRangeSel?.addEventListener("change", queuePreviewRefresh);

    // Initialize settings UI for the default chart family after handlers are ready.
    setChartType("spc");

    // Initial population of worksheet selector, then tables/columns
    tryCatch(async () => {
      await updateWorksheetSelector();
      await updateTableSelector();
      const nextTable = (document.getElementById("table-selector") as HTMLSelectElement | null)
        ?.value;
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
  const numeratorSelector = document.getElementById(
    "numerator-selector"
  ) as HTMLSelectElement | null;
  const denominatorSelector = document.getElementById(
    "denominator-selector"
  ) as HTMLSelectElement | null;
  const sdSelector = document.getElementById("sd-selector") as HTMLSelectElement | null;

  if (categorySelector)
    categorySelector.innerHTML = '<option value="" disabled selected>Select category</option>';
  if (numeratorSelector)
    numeratorSelector.innerHTML = '<option value="" disabled selected>Select numerator</option>';
  if (denominatorSelector)
    denominatorSelector.innerHTML =
      '<option value="" disabled selected>Select denominator</option>';
  if (sdSelector)
    sdSelector.innerHTML = '<option value="" disabled selected>Select SD (Xbar)</option>';
  resetSpcDateRangeFilter();
  updateSdSelectorVisibility();
  updateDenominatorSelectorVisibility();
}

function fromExcelDate(excelValue: unknown): Date | null {
  if (excelValue instanceof Date) {
    return isValidDateValue(excelValue) ? excelValue : null;
  }

  if (typeof excelValue === "number" && Number.isFinite(excelValue)) {
    const parsed = new Date((excelValue - (25567 + 2)) * 86400 * 1000);
    return isValidDateValue(parsed) ? parsed : null;
  }

  if (typeof excelValue === "string") {
    if (!excelValue.trim()) return null;
    const parsed = new Date(excelValue);
    return isValidDateValue(parsed) ? parsed : null;
  }

  return null;
}

async function updateTableSelector() {
  await Excel.run(async (context) => {
    const selectedWorksheetName = (
      document.getElementById("worksheet-selector") as HTMLSelectElement | null
    )?.value;
    if (!selectedWorksheetName) {
      throw new Error("No worksheet selected");
    }
    const worksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const tables = worksheet.tables.load("items/name");
    await context.sync();
    const tableSelector = document.getElementById("table-selector") as HTMLSelectElement;
    tableSelector.innerHTML = '<option value="" disabled selected>Select a table</option>';
    tables.items.forEach((table) => {
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

    worksheets.items.forEach((ws) => {
      const option = document.createElement("option");
      option.value = ws.name;
      option.text = ws.name;
      worksheetSelector.appendChild(option);
    });

    // Default to active worksheet if present
    const activeName = activeWorksheet.name;
    const activeExists = worksheets.items.some((ws) => ws.name === activeName);
    if (activeExists) {
      worksheetSelector.value = activeName;
    } else if (worksheets.items.length > 0) {
      worksheetSelector.value = worksheets.items[0].name;
    }
  });
}

async function updateColumnSelectors() {
  await Excel.run(async (context) => {
    const selectedWorksheetName = (
      document.getElementById("worksheet-selector") as HTMLSelectElement | null
    )?.value;
    if (!selectedWorksheetName) {
      throw new Error("No worksheet selected");
    }
    const worksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement)
      .value;
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
    const denominatorSelector = document.getElementById(
      "denominator-selector"
    ) as HTMLSelectElement;
    const sdSelector = document.getElementById("sd-selector") as HTMLSelectElement | null;
    categorySelector.innerHTML = '<option value="" disabled selected>Select category</option>';
    numeratorSelector.innerHTML = '<option value="" disabled selected>Select numerator</option>';
    denominatorSelector.innerHTML =
      '<option value="" disabled selected>Select denominator</option>';
    if (sdSelector) {
      sdSelector.innerHTML = '<option value="" disabled selected>Select SD (Xbar)</option>';
    }
    columns.items.forEach((column) => {
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
    resetSpcDateRangeFilter();
    // Columns reset, so ensure buttons reflect incomplete selection
    updateSdSelectorVisibility();
    updateActionButtonsEnabledState();
  });
}

function updateActionButtonsEnabledState() {
  const chartFamily = (document.getElementById("controlchart-selector") as HTMLInputElement | null)
    ?.value;
  const isSpc = chartFamily === "spc";
  const isXbar = isSpc && getSelectedSpcChartType() === "xbar";
  const requiredIds = [
    "worksheet-selector",
    "table-selector",
    "category-selector",
    "numerator-selector",
  ];
  if (isDenominatorRequired()) {
    requiredIds.push("denominator-selector");
  }
  if (isXbar) {
    requiredIds.push("sd-selector");
  }
  const allSelected = requiredIds.every((id) => {
    const el = document.getElementById(id) as HTMLSelectElement;
    return el && typeof el.value === "string" && el.value.length > 0;
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
    const selectedWorksheetName = (
      document.getElementById("worksheet-selector") as HTMLSelectElement | null
    )?.value;
    if (!selectedWorksheetName) {
      throw new Error("No worksheet selected");
    }
    const currentWorksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement)
      .value;
    if (!selectedTableName) {
      throw new Error("No table selected");
    }
    const table = currentWorksheet.tables.getItem(selectedTableName);
    const selectedCategoryColumn = (
      document.getElementById("category-selector") as HTMLSelectElement
    ).value;
    const selectedNumeratorColumn = (
      document.getElementById("numerator-selector") as HTMLSelectElement
    ).value;
    const selectedDenominatorColumn = (
      document.getElementById("denominator-selector") as HTMLSelectElement
    ).value;
    const selectedSdColumn = (document.getElementById("sd-selector") as HTMLSelectElement | null)
      ?.value;

    const categoryColumn = table.columns
      .getItem(selectedCategoryColumn)
      .getDataBodyRange()
      .load("values");
    const numeratorsColumn = table.columns
      .getItem(selectedNumeratorColumn)
      .getDataBodyRange()
      .load("values");
    const controlChartType = (document.getElementById("controlchart-selector") as HTMLInputElement)
      .value;
    updateCurrentChartInputSettingsFromUi(controlChartType as ChartFamily);

    const denomRequired = isDenominatorRequired();
    if (denomRequired && !selectedDenominatorColumn) {
      throw new Error(
        "This chart type requires a Denominator column. Please select a Denominator under Data / Inputs."
      );
    }
    const denominatorsColumn =
      denomRequired && selectedDenominatorColumn
        ? table.columns.getItem(selectedDenominatorColumn).getDataBodyRange().load("values")
        : null;

    const needsXbarSd = controlChartType === "spc" && spcInputSettings.spc.chart_type === "xbar";
    if (needsXbarSd && !selectedSdColumn) {
      throw new Error(
        "Xbar requires an SD column. Please select an SD column (Xbar) under Data / Inputs."
      );
    }

    const sdColumnRange = needsXbarSd
      ? table.columns.getItem(selectedSdColumn!).getDataBodyRange().load("values")
      : null;
    await context.sync();
    updateCurrentChartInputSettingsFromUi(controlChartType as ChartFamily);

    const rawData: RawDataRow[] = categoryColumn.values.flat().map((cat, i) => {
      const row: any = {
        categories: controlChartType === "spc" ? fromExcelDate(cat) : cat,
        numerators: numeratorsColumn.values.flat()[i],
      };
      if (controlChartType === "funnel" && funnelInputSettings.labels.show_labels) {
        row.labels = cat == null ? "" : String(cat);
      }
      if (denominatorsColumn) {
        row.denominators = denominatorsColumn.values.flat()[i];
      }
      if (needsXbarSd && sdColumnRange) {
        row.xbar_sds = (sdColumnRange.values.flat() as any[])[i];
      }
      return row;
    });
    const filteredRawData = controlChartType === "spc" ? applySpcDateRangeFilter(rawData) : rawData;
    const useFormattedDates =
      controlChartType === "spc" && rawDataSupportsDateFormatting(filteredRawData);
    updateHeaderCanvasPadding(controlChartType, useFormattedDates);

    var updateArgs = {
      dataViews: makeUpdateValues(
        filteredRawData,
        controlChartType === "spc" ? spcInputSettings : funnelInputSettings,
        getChartAggregations(controlChartType as ChartFamily),
        useFormattedDates
      ).dataViews,
      viewport: { width: 640, height: 480 },
      type: 2, //,
      //headless: true,
      //frontend: true
    };

    var currVisual = controlChartType === "spc" ? spcVisual : funnelVisual;

    currVisual.update(updateArgs as any);
    drawChartFrameAndHeader(currVisual, filteredRawData, controlChartType);

    var image = currentWorksheet.shapes.addImage(
      btoa((currVisual.svg.node() as SVGSVGElement).outerHTML)
    );
    image.name = "Image";
    image.top = 10;
    image.left = 200;

    await context.sync();
  });
}

async function previewPlot() {
  // Render the chart in the side-pane preview area without inserting into main Excel area
  const selectedWorksheetName = (
    document.getElementById("worksheet-selector") as HTMLSelectElement | null
  )?.value;
  const selectedTableName = (document.getElementById("table-selector") as HTMLSelectElement).value;
  const selectedCategoryColumn = (document.getElementById("category-selector") as HTMLSelectElement)
    .value;
  const selectedNumeratorColumn = (
    document.getElementById("numerator-selector") as HTMLSelectElement
  ).value;
  const selectedDenominatorColumn = (
    document.getElementById("denominator-selector") as HTMLSelectElement
  ).value;
  const selectedSdColumn = (document.getElementById("sd-selector") as HTMLSelectElement | null)
    ?.value;

  if (
    !selectedWorksheetName ||
    !selectedTableName ||
    !selectedCategoryColumn ||
    !selectedNumeratorColumn
  ) {
    throw new Error("Please select a worksheet, table, category and numerator to preview.");
  }

  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getItem(selectedWorksheetName);
    const table = currentWorksheet.tables.getItem(selectedTableName);

    const categoryColumn = table.columns
      .getItem(selectedCategoryColumn)
      .getDataBodyRange()
      .load("values");
    const numeratorsColumn = table.columns
      .getItem(selectedNumeratorColumn)
      .getDataBodyRange()
      .load("values");

    const controlChartType = (document.getElementById("controlchart-selector") as HTMLInputElement)
      .value;
    updateCurrentChartInputSettingsFromUi(controlChartType as ChartFamily);

    const denomRequired = isDenominatorRequired();
    if (denomRequired && !selectedDenominatorColumn) {
      throw new Error(
        "This chart type requires a Denominator column. Please select a Denominator under Data / Inputs."
      );
    }
    const denominatorsColumn =
      denomRequired && selectedDenominatorColumn
        ? table.columns.getItem(selectedDenominatorColumn).getDataBodyRange().load("values")
        : null;
    const needsXbarSd = controlChartType === "spc" && spcInputSettings.spc.chart_type === "xbar";
    if (needsXbarSd && !selectedSdColumn) {
      throw new Error(
        "Xbar requires an SD column. Please select an SD column (Xbar) under Data / Inputs."
      );
    }
    const sdColumnRange = needsXbarSd
      ? table.columns.getItem(selectedSdColumn!).getDataBodyRange().load("values")
      : null;

    await context.sync();

    const rawData: RawDataRow[] = categoryColumn.values.flat().map((cat, i) => {
      const row: any = {
        categories: controlChartType === "spc" ? fromExcelDate(cat) : cat,
        numerators: numeratorsColumn.values.flat()[i],
      };
      if (controlChartType === "funnel" && funnelInputSettings.labels.show_labels) {
        row.labels = cat == null ? "" : String(cat);
      }
      if (denominatorsColumn) {
        row.denominators = denominatorsColumn.values.flat()[i];
      }
      if (needsXbarSd && sdColumnRange) {
        row.xbar_sds = (sdColumnRange.values.flat() as any[])[i];
      }
      return row;
    });
    const filteredRawData = controlChartType === "spc" ? applySpcDateRangeFilter(rawData) : rawData;
    const useFormattedDates =
      controlChartType === "spc" && rawDataSupportsDateFormatting(filteredRawData);
    updateHeaderCanvasPadding(controlChartType, useFormattedDates);

    const previewHost = document.getElementById("preview-container") as HTMLElement;
    const containerRect = previewHost.getBoundingClientRect();
    const padding = 8 * 2; // preview container padding
    const width = Math.max(320, Math.floor(containerRect.width - padding));
    const height = Math.max(220, Math.floor(containerRect.height - padding));

    const updateArgs = {
      dataViews: makeUpdateValues(
        filteredRawData,
        controlChartType === "spc" ? spcInputSettings : funnelInputSettings,
        getChartAggregations(controlChartType as ChartFamily),
        useFormattedDates
      ).dataViews,
      viewport: { width, height },
      type: 2,
    } as any;

    const currDiv = controlChartType === "spc" ? spcDiv : funnelDiv;
    const otherDiv = controlChartType === "spc" ? funnelDiv : spcDiv;
    currDiv.removeAttribute("hidden");
    otherDiv.setAttribute("hidden", "true");

    const currVisual = controlChartType === "spc" ? spcVisual : funnelVisual;
    currVisual.update(updateArgs);
    // Remove any mouse handlers that power the tooltip on the root svg (defense in depth)
    (currVisual.svg as any).on("mousemove", null).on("mouseleave", null);
    drawChartFrameAndHeader(currVisual, filteredRawData, controlChartType);
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
