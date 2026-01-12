import { createInputField, createSelectField } from "./uiFields";

export function renderSpcDataSettings() {
  const host = document.getElementById("spc-data-settings-grid") as HTMLElement | null;
  if (!host) return;

  host.innerHTML = "";

  host.appendChild(createSelectField({
    id: "spc-chart-type",
    label: "Chart type",
    title: "Chart type",
    value: "i",
    options: [
      { value: "run", text: "run - Run Chart" },
      { value: "i", text: "i - Individual Measurements" },
      { value: "i_m", text: "i_m - Individual Measurements: Median centerline" },
      { value: "i_mm", text: "i_mm - Individual Measurements: Median centerline, Median MR Limits" },
      { value: "mr", text: "mr - Moving Range of Individual Measurements" },
      { value: "p", text: "p - Proportions" },
      { value: "pp", text: "p prime - Proportions: Large-Sample Corrected" },
      { value: "u", text: "u - Rates" },
      { value: "up", text: "u prime - Rates: Large-Sample Correction" },
      { value: "c", text: "c - Counts" },
      { value: "xbar", text: "xbar - Sample Means" },
      { value: "s", text: "s - Sample SDs" },
      { value: "g", text: "g - Number of Non-Events Between Events" },
      { value: "t", text: "t - Time Between Events" }
    ]
  }));

  host.appendChild(createSelectField({
    id: "spc-outliers-in-limits",
    label: "Keep Outliers in Limit Calcs.",
    title: "Keep Outliers in Limit Calcs.",
    value: "false",
    options: [
      { value: "false", text: "Off" },
      { value: "true", text: "On" }
    ]
  }));

  host.appendChild(createInputField({
    id: "spc-multiplier",
    label: "Multiplier",
    type: "number",
    value: "1",
    min: "0",
    step: "0.1"
  }));

  host.appendChild(createInputField({
    id: "spc-sig-figs",
    label: "Decimals to Report",
    type: "number",
    value: "2",
    min: "0",
    max: "20",
    step: "1"
  }));

  host.appendChild(createSelectField({
    id: "spc-perc-labels",
    label: "Report as percentage",
    title: "Report as percentage",
    value: "Automatic",
    options: [
      { value: "Automatic", text: "Automatic" },
      { value: "Yes", text: "Yes" },
      { value: "No", text: "No" }
    ]
  }));

  host.appendChild(createSelectField({
    id: "spc-split-on-click",
    label: "Split Limits on Click",
    title: "Split Limits on Click",
    value: "false",
    options: [
      { value: "false", text: "Off" },
      { value: "true", text: "On" }
    ]
  }));

  host.appendChild(createInputField({
    id: "spc-num-points-subset",
    label: "Subset Number of Points for Limit Calculations",
    type: "number",
    placeholder: "(optional)",
    min: "1",
    step: "1"
  }));

  host.appendChild(createSelectField({
    id: "spc-subset-points-from",
    label: "Subset Points From",
    title: "Subset Points From",
    value: "Start",
    options: [
      { value: "Start", text: "Start" },
      { value: "End", text: "End" }
    ]
  }));

//   host.appendChild(createSelectField({
//     id: "spc-ttip-show-date",
//     label: "Show Date in Tooltip",
//     title: "Show Date in Tooltip",
//     value: "true",
//     options: [
//       { value: "true", text: "On" },
//       { value: "false", text: "Off" }
//     ]
//   }));

//   host.appendChild(createInputField({
//     id: "spc-ttip-label-date",
//     label: "Date Tooltip label",
//     type: "text",
//     value: "Date"
//   }));

//   host.appendChild(createSelectField({
//     id: "spc-ttip-show-numerator",
//     label: "Show numerator in Tooltip",
//     title: "Show numerator in Tooltip",
//     value: "true",
//     options: [
//       { value: "true", text: "On" },
//       { value: "false", text: "Off" }
//     ]
//   }));

//   host.appendChild(createInputField({
//     id: "spc-ttip-label-numerator",
//     label: "Numerator Tooltip Label",
//     type: "text",
//     value: "Numerator"
//   }));

//   host.appendChild(createSelectField({
//     id: "spc-ttip-show-denominator",
//     label: "Show Denominator in Tooltip",
//     title: "Show Denominator in Tooltip",
//     value: "true",
//     options: [
//       { value: "true", text: "On" },
//       { value: "false", text: "Off" }
//     ]
//   }));

//   host.appendChild(createInputField({
//     id: "spc-ttip-label-denominator",
//     label: "Denominator Tooltip Label",
//     type: "text",
//     value: "Denominator"
//   }));

//   host.appendChild(createSelectField({
//     id: "spc-ttip-show-value",
//     label: "Show Value in Tooltip",
//     title: "Show Value in Tooltip",
//     value: "true",
//     options: [
//       { value: "true", text: "On" },
//       { value: "false", text: "Off" }
//     ]
//   }));

//   host.appendChild(createInputField({
//     id: "spc-ttip-label-value",
//     label: "Value Tooltip Label",
//     type: "text",
//     value: "Automatic"
//   }));

  host.appendChild(createInputField({
    id: "spc-ll-truncate",
    label: "Truncate Lower Limits at:",
    type: "number",
    step: "0.1",
    placeholder: "(optional)"
  }));

  host.appendChild(createInputField({
    id: "spc-ul-truncate",
    label: "Truncate Upper Limits at:",
    type: "number",
    step: "0.1",
    placeholder: "(optional)"
  }));
}
