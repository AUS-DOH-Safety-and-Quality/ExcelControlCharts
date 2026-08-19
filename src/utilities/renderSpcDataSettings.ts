import type { settingsValueType as funnelSettingsValueType } from "../PowerBI-Funnels/src/settings";
import type { settingsValueType as spcSettingsValueType } from "../PowerBI-SPC/src/settings";
import { createInputField, createSelectField } from "./uiFields";

type ChartFamily = "spc" | "funnel";

function formatOptionalNumber(value: number | undefined): string | undefined {
  return value === undefined ? undefined : String(value);
}

function setSectionState(sectionId: string, title: string, visible: boolean) {
  const section = document.getElementById(sectionId) as HTMLDetailsElement | null;
  if (!section) return;

  section.hidden = !visible;
  section.style.display = visible ? "" : "none";

  const titleElement = section.querySelector("summary span") as HTMLElement | null;
  if (titleElement) {
    titleElement.textContent = title;
  }
}

function renderSpcDataFields(host: HTMLElement, spcSettings: spcSettingsValueType) {
  host.appendChild(
    createSelectField({
      id: "spc-chart-type",
      label: "Chart type",
      title: "Chart type",
      value: spcSettings.spc.chart_type,
      options: [
        { value: "run", text: "run - Run Chart" },
        { value: "i", text: "i - Individual Measurements" },
        { value: "i_m", text: "i_m - Individual Measurements: Median centerline" },
        {
          value: "i_mm",
          text: "i_mm - Individual Measurements: Median centerline, Median MR Limits",
        },
        { value: "mr", text: "mr - Moving Range of Individual Measurements" },
        { value: "p", text: "p - Proportions" },
        { value: "pp", text: "p prime - Proportions: Large-Sample Corrected" },
        { value: "u", text: "u - Rates" },
        { value: "up", text: "u prime - Rates: Large-Sample Correction" },
        { value: "c", text: "c - Counts" },
        { value: "xbar", text: "xbar - Sample Means" },
        { value: "s", text: "s - Sample SDs" },
        { value: "g", text: "g - Number of Non-Events Between Events" },
        { value: "t", text: "t - Time Between Events" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-outliers-in-limits",
      label: "Keep Outliers in Limit Calcs.",
      title: "Keep Outliers in Limit Calcs.",
      value: String(spcSettings.spc.outliers_in_limits),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-multiplier",
      label: "Multiplier",
      type: "number",
      value: String(spcSettings.spc.multiplier),
      min: "0",
      step: "0.1",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-sig-figs",
      label: "Decimals to Report",
      type: "number",
      value: String(spcSettings.spc.sig_figs),
      min: "0",
      max: "20",
      step: "1",
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-perc-labels",
      label: "Report as percentage",
      title: "Report as percentage",
      value: spcSettings.spc.perc_labels,
      options: [
        { value: "Automatic", text: "Automatic" },
        { value: "Yes", text: "Yes" },
        { value: "No", text: "No" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-split-on-click",
      label: "Split Limits on Click",
      title: "Split Limits on Click",
      value: String(spcSettings.spc.split_on_click),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-num-points-subset",
      label: "Subset Number of Points for Limit Calculations",
      type: "number",
      value: formatOptionalNumber(spcSettings.spc.num_points_subset),
      placeholder: "(optional)",
      min: "1",
      step: "1",
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-subset-points-from",
      label: "Subset Points From",
      title: "Subset Points From",
      value: spcSettings.spc.subset_points_from,
      options: [
        { value: "Start", text: "Start" },
        { value: "End", text: "End" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-ll-truncate",
      label: "Truncate Lower Limits at:",
      type: "number",
      value: formatOptionalNumber(spcSettings.spc.ll_truncate),
      step: "0.1",
      placeholder: "(optional)",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-ul-truncate",
      label: "Truncate Upper Limits at:",
      type: "number",
      value: formatOptionalNumber(spcSettings.spc.ul_truncate),
      step: "0.1",
      placeholder: "(optional)",
    })
  );
}

function renderSpcPatternFields(host: HTMLElement, spcSettings: spcSettingsValueType) {
  host.appendChild(
    createSelectField({
      id: "spc-show-variation-icons",
      label: "Show variation icons",
      title: "Show variation icons",
      value: String(spcSettings.nhs_icons.show_variation_icons),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-flag-last-point",
      label: "Flag only last point",
      title: "Flag only last point",
      value: String(spcSettings.nhs_icons.flag_last_point),
      options: [
        { value: "true", text: "Yes" },
        { value: "false", text: "No" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-variation-location",
      label: "Variation icon location",
      title: "Variation icon location",
      value: spcSettings.nhs_icons.variation_icons_locations,
      options: [
        { value: "Top Right", text: "Top Right" },
        { value: "Bottom Right", text: "Bottom Right" },
        { value: "Top Left", text: "Top Left" },
        { value: "Bottom Left", text: "Bottom Left" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-variation-scaling",
      label: "Variation icon scaling",
      type: "number",
      value: String(spcSettings.nhs_icons.variation_icons_scaling),
      min: "0",
      step: "0.1",
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-show-assurance-icons",
      label: "NHS assurance icon",
      title: "NHS assurance icon",
      value: String(spcSettings.nhs_icons.show_assurance_icons),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-assurance-scaling",
      label: "Assurance icon scaling",
      type: "number",
      value: String(spcSettings.nhs_icons.assurance_icons_scaling),
      min: "0",
      step: "0.1",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-alt-target",
      label: "Additional target value",
      type: "number",
      value: formatOptionalNumber(spcSettings.lines.alt_target),
      step: "0.1",
      placeholder: "(optional)",
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-alt-target-label",
      label: "Alt target line label",
      title: "Alt target line label",
      value: spcSettings.lines.plot_label_show_alt_target ? "target_equals" : "off",
      options: [
        { value: "off", text: "Off" },
        { value: "target_equals", text: "Target =" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-improvement-direction",
      label: "Improvement direction",
      title: "Improvement direction",
      value: spcSettings.outliers.improvement_direction,
      options: [
        { value: "increase", text: "Increase" },
        { value: "decrease", text: "Decrease" },
        { value: "neutral", text: "Neutral" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-assurance-location",
      label: "Assurance icon location",
      title: "Assurance icon location",
      value: spcSettings.nhs_icons.assurance_icons_locations,
      options: [
        { value: "Top Right", text: "Top Right" },
        { value: "Bottom Right", text: "Bottom Right" },
        { value: "Top Left", text: "Top Left" },
        { value: "Bottom Left", text: "Bottom Left" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-astronomical-points",
      label: "Highlight point beyond control limit",
      title: "Highlight point beyond control limit",
      value: String(spcSettings.outliers.astronomical),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-astronomical-limit",
      label: "Control limit threshold",
      title: "Control limit threshold",
      value: spcSettings.outliers.astronomical_limit,
      options: [
        { value: "1 Sigma", text: "1 Sigma" },
        { value: "2 Sigma", text: "2 Sigma" },
        { value: "3 Sigma", text: "3 Sigma" },
        { value: "Specification", text: "Specification" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-trend-pattern",
      label: "Highlight trend pattern",
      title: "Highlight trend pattern",
      value: String(spcSettings.outliers.trend),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-trend-points",
      label: "Trend points",
      type: "number",
      value: String(spcSettings.outliers.trend_n),
      min: "1",
      step: "1",
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-two-in-three",
      label: "Highlight two-in-three",
      title: "Highlight two-in-three",
      value: String(spcSettings.outliers.two_in_three),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-two-in-three-highlight-series",
      label: "Highlight full two-in-three pattern",
      title: "Highlight full two-in-three pattern",
      value: String(spcSettings.outliers.two_in_three_highlight_series),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-two-in-three-limit",
      label: "Two-in-three warning limit",
      title: "Two-in-three warning limit",
      value: spcSettings.outliers.two_in_three_limit,
      options: [
        { value: "1 Sigma", text: "1 Sigma" },
        { value: "2 Sigma", text: "2 Sigma" },
        { value: "3 Sigma", text: "3 Sigma" },
        { value: "Specification", text: "Specification" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-shift-pattern",
      label: "Highlight shift pattern",
      title: "Highlight shift pattern",
      value: String(spcSettings.outliers.shift),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-shift-points",
      label: "Shift points",
      type: "number",
      value: String(spcSettings.outliers.shift_n),
      min: "1",
      step: "1",
    })
  );
}

function renderDateFields(host: HTMLElement, spcSettings: spcSettingsValueType) {
  host.appendChild(
    createSelectField({
      id: "spc-date-format-day",
      label: "Day format",
      title: "Day format",
      value: spcSettings.dates.date_format_day,
      options: [
        { value: "DD", text: "DD" },
        { value: "Thurs DD", text: "Thu DD" },
        { value: "Thursday DD", text: "Thursday DD" },
        { value: "(blank)", text: "(blank)" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-date-format-month",
      label: "Month format",
      title: "Month format",
      value: spcSettings.dates.date_format_month,
      options: [
        { value: "MM", text: "MM" },
        { value: "Mon", text: "Mon" },
        { value: "Month", text: "Month" },
        { value: "(blank)", text: "(blank)" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-date-format-year",
      label: "Year format",
      title: "Year format",
      value: spcSettings.dates.date_format_year,
      options: [
        { value: "YYYY", text: "YYYY" },
        { value: "YY", text: "YY" },
        { value: "(blank)", text: "(blank)" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-date-format-delim",
      label: "Date delimiter",
      title: "Date delimiter",
      value: spcSettings.dates.date_format_delim,
      options: [
        { value: "/", text: "/" },
        { value: "-", text: "-" },
        { value: " ", text: "Space" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-date-format-locale",
      label: "Date locale",
      title: "Date locale",
      value: spcSettings.dates.date_format_locale,
      options: [
        { value: "en-GB", text: "en-GB" },
        { value: "en-US", text: "en-US" },
      ],
    })
  );
}

function renderFunnelDataFields(host: HTMLElement, funnelSettings: funnelSettingsValueType) {
  host.appendChild(
    createSelectField({
      id: "spc-chart-type",
      label: "Chart type",
      title: "Chart type",
      value: funnelSettings.funnel.chart_type,
      options: [
        { value: "SR", text: "Indirectly Standardised (HSMR)" },
        { value: "PR", text: "Proportion" },
        { value: "RC", text: "Rate" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "funnel-od-adjust",
      label: "OD adjustment",
      title: "OD adjustment",
      value: funnelSettings.funnel.od_adjust,
      options: [
        { value: "auto", text: "Automatic" },
        { value: "yes", text: "Yes" },
        { value: "no", text: "No" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-multiplier",
      label: "Multiplier",
      type: "number",
      value: String(funnelSettings.funnel.multiplier),
      min: "0",
      step: "0.1",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-sig-figs",
      label: "Decimals to Report",
      type: "number",
      value: String(funnelSettings.funnel.sig_figs),
      min: "0",
      max: "20",
      step: "1",
    })
  );

  host.appendChild(
    createSelectField({
      id: "spc-perc-labels",
      label: "Report as percentage",
      title: "Report as percentage",
      value: funnelSettings.funnel.perc_labels,
      options: [
        { value: "Automatic", text: "Automatic" },
        { value: "Yes", text: "Yes" },
        { value: "No", text: "No" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "funnel-transformation",
      label: "Transformation",
      title: "Transformation",
      value: funnelSettings.funnel.transformation,
      options: [
        { value: "none", text: "None" },
        { value: "ln", text: "Natural Log (y+1)" },
        { value: "log10", text: "Log10 (y+1)" },
        { value: "sqrt", text: "Square-Root" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "funnel-category-point-labels",
      label: "Category labels beside points",
      title: "Show the selected category next to each funnel point.",
      value: String(funnelSettings.labels.show_labels),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-ll-truncate",
      label: "Truncate Lower Limits at:",
      type: "number",
      value: formatOptionalNumber(funnelSettings.funnel.ll_truncate),
      step: "0.1",
      placeholder: "(optional)",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-ul-truncate",
      label: "Truncate Upper Limits at:",
      type: "number",
      value: formatOptionalNumber(funnelSettings.funnel.ul_truncate),
      step: "0.1",
      placeholder: "(optional)",
    })
  );
}

function renderFunnelOutlierFields(host: HTMLElement, funnelSettings: funnelSettingsValueType) {
  host.appendChild(
    createSelectField({
      id: "funnel-process-flag-type",
      label: "Type of change to flag",
      title: "Type of change to flag",
      value: funnelSettings.outliers.process_flag_type,
      options: [
        { value: "both", text: "Both" },
        { value: "improvement", text: "Improvement" },
        { value: "deterioration", text: "Deterioration" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "funnel-improvement-direction",
      label: "Improvement direction",
      title: "Improvement direction",
      value: funnelSettings.outliers.improvement_direction,
      options: [
        { value: "increase", text: "Increase" },
        { value: "neutral", text: "Neutral" },
        { value: "decrease", text: "Decrease" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "funnel-three-sigma",
      label: "Three sigma outliers",
      title: "Three sigma outliers",
      value: String(funnelSettings.outliers.three_sigma),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );

  host.appendChild(
    createSelectField({
      id: "funnel-two-sigma",
      label: "Two sigma outliers",
      title: "Two sigma outliers",
      value: String(funnelSettings.outliers.two_sigma),
      options: [
        { value: "false", text: "Off" },
        { value: "true", text: "On" },
      ],
    })
  );
}

export function renderSpcDataSettings(
  chartFamily: ChartFamily,
  spcSettings: spcSettingsValueType,
  funnelSettings: funnelSettingsValueType
) {
  const host = document.getElementById("spc-data-settings-grid") as HTMLElement | null;
  const nhsHost = document.getElementById("spc-nhs-settings-grid") as HTMLElement | null;
  const dateHost = document.getElementById("spc-date-settings-grid") as HTMLElement | null;
  if (!host) return;

  host.innerHTML = "";
  if (nhsHost) nhsHost.innerHTML = "";
  if (dateHost) dateHost.innerHTML = "";

  if (chartFamily === "spc") {
    setSectionState("spc-data-settings", "Data Settings", true);
    setSectionState("spc-nhs-settings", "NHS / Pattern Settings", true);
    setSectionState("spc-date-settings", "Date Formatting", true);

    renderSpcDataFields(host, spcSettings);
    if (nhsHost) {
      renderSpcPatternFields(nhsHost, spcSettings);
    }
    if (dateHost) {
      renderDateFields(dateHost, spcSettings);
    }
  } else {
    setSectionState("spc-data-settings", "Data Settings", true);
    setSectionState("spc-nhs-settings", "Outlier Settings", true);
    setSectionState("spc-date-settings", "Date Formatting", false);

    renderFunnelDataFields(host, funnelSettings);
    if (nhsHost) {
      renderFunnelOutlierFields(nhsHost, funnelSettings);
    }
  }
}
