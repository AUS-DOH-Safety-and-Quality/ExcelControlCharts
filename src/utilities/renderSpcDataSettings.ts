import { createInputField, createSelectField } from "./uiFields";

export function renderSpcDataSettings() {
  const host = document.getElementById("spc-data-settings-grid") as HTMLElement | null;
  const nhsHost = document.getElementById("spc-nhs-settings-grid") as HTMLElement | null;
  const dateHost = document.getElementById("spc-date-settings-grid") as HTMLElement | null;
  if (!host) return;

  host.innerHTML = "";
  if (nhsHost) nhsHost.innerHTML = "";
  if (dateHost) dateHost.innerHTML = "";

  host.appendChild(
    createSelectField({
      id: "spc-chart-type",
      label: "Chart type",
      title: "Chart type",
      value: "i",
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
      value: "false",
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
      value: "1",
      min: "0",
      step: "0.1",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-sig-figs",
      label: "Decimals to Report",
      type: "number",
      value: "2",
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
      value: "Automatic",
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
      value: "false",
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
      value: "Start",
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
      step: "0.1",
      placeholder: "(optional)",
    })
  );

  host.appendChild(
    createInputField({
      id: "spc-ul-truncate",
      label: "Truncate Upper Limits at:",
      type: "number",
      step: "0.1",
      placeholder: "(optional)",
    })
  );

  if (nhsHost) {
    nhsHost.appendChild(
      createSelectField({
        id: "spc-show-variation-icons",
        label: "Show variation icons",
        title: "Show variation icons",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-flag-last-point",
        label: "Flag only last point",
        title: "Flag only last point",
        value: "true",
        options: [
          { value: "true", text: "Yes" },
          { value: "false", text: "No" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-variation-location",
        label: "Variation icon location",
        title: "Variation icon location",
        value: "Top Right",
        options: [
          { value: "Top Right", text: "Top Right" },
          { value: "Bottom Right", text: "Bottom Right" },
          { value: "Top Left", text: "Top Left" },
          { value: "Bottom Left", text: "Bottom Left" },
        ],
      })
    );

    nhsHost.appendChild(
      createInputField({
        id: "spc-variation-scaling",
        label: "Variation icon scaling",
        type: "number",
        value: "1",
        min: "0",
        step: "0.1",
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-show-assurance-icons",
        label: "NHS assurance icon",
        title: "NHS assurance icon",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createInputField({
        id: "spc-assurance-scaling",
        label: "Assurance icon scaling",
        type: "number",
        value: "1",
        min: "0",
        step: "0.1",
      })
    );

    nhsHost.appendChild(
      createInputField({
        id: "spc-alt-target",
        label: "Additional target value",
        type: "number",
        step: "0.1",
        placeholder: "(optional)",
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-improvement-direction",
        label: "Improvement direction",
        title: "Improvement direction",
        value: "increase",
        options: [
          { value: "increase", text: "Increase" },
          { value: "decrease", text: "Decrease" },
          { value: "neutral", text: "Neutral" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-assurance-location",
        label: "Assurance icon location",
        title: "Assurance icon location",
        value: "Bottom Right",
        options: [
          { value: "Top Right", text: "Top Right" },
          { value: "Bottom Right", text: "Bottom Right" },
          { value: "Top Left", text: "Top Left" },
          { value: "Bottom Left", text: "Bottom Left" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-astronomical-points",
        label: "Highlight point beyond control limit",
        title: "Highlight point beyond control limit",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-astronomical-limit",
        label: "Control limit threshold",
        title: "Control limit threshold",
        value: "3 Sigma",
        options: [
          { value: "1 Sigma", text: "1 Sigma" },
          { value: "2 Sigma", text: "2 Sigma" },
          { value: "3 Sigma", text: "3 Sigma" },
          { value: "Specification", text: "Specification" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-trend-pattern",
        label: "Highlight trend pattern",
        title: "Highlight trend pattern",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createInputField({
        id: "spc-trend-points",
        label: "Trend points",
        type: "number",
        value: "5",
        min: "1",
        step: "1",
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-two-in-three",
        label: "Highlight two-in-three",
        title: "Highlight two-in-three",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-two-in-three-highlight-series",
        label: "Highlight full two-in-three pattern",
        title: "Highlight full two-in-three pattern",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-two-in-three-limit",
        label: "Two-in-three warning limit",
        title: "Two-in-three warning limit",
        value: "2 Sigma",
        options: [
          { value: "1 Sigma", text: "1 Sigma" },
          { value: "2 Sigma", text: "2 Sigma" },
          { value: "3 Sigma", text: "3 Sigma" },
          { value: "Specification", text: "Specification" },
        ],
      })
    );

    nhsHost.appendChild(
      createSelectField({
        id: "spc-shift-pattern",
        label: "Highlight shift pattern",
        title: "Highlight shift pattern",
        value: "false",
        options: [
          { value: "false", text: "Off" },
          { value: "true", text: "On" },
        ],
      })
    );

    nhsHost.appendChild(
      createInputField({
        id: "spc-shift-points",
        label: "Shift points",
        type: "number",
        value: "7",
        min: "1",
        step: "1",
      })
    );
  }

  if (dateHost) {
    dateHost.appendChild(
      createSelectField({
        id: "spc-date-format-day",
        label: "Day format",
        title: "Day format",
        value: "DD",
        options: [
          { value: "DD", text: "DD" },
          { value: "Thurs DD", text: "Thu DD" },
          { value: "Thursday DD", text: "Thursday DD" },
          { value: "(blank)", text: "(blank)" },
        ],
      })
    );

    dateHost.appendChild(
      createSelectField({
        id: "spc-date-format-month",
        label: "Month format",
        title: "Month format",
        value: "MM",
        options: [
          { value: "MM", text: "MM" },
          { value: "Mon", text: "Mon" },
          { value: "Month", text: "Month" },
          { value: "(blank)", text: "(blank)" },
        ],
      })
    );

    dateHost.appendChild(
      createSelectField({
        id: "spc-date-format-year",
        label: "Year format",
        title: "Year format",
        value: "YYYY",
        options: [
          { value: "YYYY", text: "YYYY" },
          { value: "YY", text: "YY" },
          { value: "(blank)", text: "(blank)" },
        ],
      })
    );

    dateHost.appendChild(
      createSelectField({
        id: "spc-date-format-delim",
        label: "Date delimiter",
        title: "Date delimiter",
        value: "/",
        options: [
          { value: "/", text: "/" },
          { value: "-", text: "-" },
          { value: " ", text: "Space" },
        ],
      })
    );

    dateHost.appendChild(
      createSelectField({
        id: "spc-date-format-locale",
        label: "Date locale",
        title: "Date locale",
        value: "en-GB",
        options: [
          { value: "en-GB", text: "en-GB" },
          { value: "en-US", text: "en-US" },
        ],
      })
    );
  }
}
