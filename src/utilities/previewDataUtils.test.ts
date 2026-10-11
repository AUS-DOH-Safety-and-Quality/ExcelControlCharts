import { expect, test } from "bun:test";
import SpcViewModel from "powerbi-spc/src/Classes/viewModelClass";
import FunnelViewModel from "powerbi-funnels/src/Classes/viewModelClass";
import { defaultSettings as spcDefaults } from "powerbi-spc/src/settings";
import { defaultSettings as funnelDefaults } from "powerbi-funnels/src/settings";
import { makeConstructorArgs, makeUpdateValues } from "./previewDataUtils";

test("Excel data and settings reach the SPC model without a settings patch", () => {
  const settings = structuredClone(spcDefaults);
  settings.canvas.left_padding = 75;
  settings.lines.show_alt_target = true;
  settings.lines.alt_target = 11;
  const firstDate = new Date(Date.UTC(2024, 0, 1));
  const options = makeUpdateValues(
    [
      { categories: firstDate, numerators: 4 },
      { categories: firstDate, numerators: 6 },
      { categories: new Date(Date.UTC(2024, 1, 1)), numerators: 12 },
      { categories: new Date(Date.UTC(2024, 2, 1)), numerators: 14 },
    ],
    settings,
    { numerators: "sum" },
    true
  );
  options.viewport = { width: 640, height: 480 };
  const host = makeConstructorArgs({} as HTMLElement).host;
  const model = new SpcViewModel();

  expect(model.update(options, host).status).toBe(true);
  expect(model.inputSettings.settings[0]).toEqual(settings);
  expect(model.controlLimits[0].values).toEqual([10, 12, 14]);
  expect(model.plotPoints).toHaveLength(3);
  expect(model.plotPoints[0].table_row.alt_target).toBe(11);
  expect(model.plotPoints[0].table_row.target).toBe(12);
  expect(model.tickLabels[0].label).toBe("01/01/2024");

  const firstKey = model.plotPoints[0].identity.getKey();
  expect(model.plotPoints[1].identity.getKey()).not.toBe(firstKey);
  expect(model.update(options, host).status).toBe(true);
  expect(model.plotPoints[0].identity.getKey()).toBe(firstKey);
});

test("Excel aggregates and category labels reach the Funnel model", () => {
  const settings = structuredClone(funnelDefaults);
  settings.funnel.chart_type = "PR";
  settings.funnel.od_adjust = "no";
  settings.labels.show_labels = true;
  const options = makeUpdateValues(
    [
      { categories: "A", numerators: 2, denominators: 50, labels: "Hospital A" },
      { categories: "A", numerators: 3, denominators: 50, labels: "Hospital A" },
      { categories: "B", numerators: 8, denominators: 80, labels: "Hospital B" },
      { categories: "C", numerators: 6, denominators: 120, labels: "Hospital C" },
    ],
    settings,
    { numerators: "sum", denominators: "sum", labels: "first" }
  );
  options.viewport = { width: 640, height: 480 };
  const host = makeConstructorArgs({} as HTMLElement).host;
  const model = new FunnelViewModel(host);

  expect(model.update(options, host).status).toBe(true);
  expect(model.inputSettings.settings).toEqual(settings);
  expect(model.plotPoints).toHaveLength(3);
  expect(model.plotPoints[0].x).toBe(100);
  expect(model.plotPoints[0].value).toBe(5);
  expect(model.inputSettings.derivedSettings.percentLabels).toBe(true);
  expect(model.plotPoints[0].label.text_value).toBe("Hospital A");

  const firstKey = model.plotPoints[0].identity.getKey();
  expect(model.plotPoints[1].identity.getKey()).not.toBe(firstKey);
  expect(model.update(options, host).status).toBe(true);
  expect(model.plotPoints[0].identity.getKey()).toBe(firstKey);
});
