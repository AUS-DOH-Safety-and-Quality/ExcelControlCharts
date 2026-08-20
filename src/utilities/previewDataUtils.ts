import type powerbi from "powerbi-visuals-api";
type VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
type ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
type IPromise<T> = powerbi.IPromise<T>;
type ISelectionId = powerbi.visuals.ISelectionId;
type IColorInfo = powerbi.IColorInfo;
type IColorPalette = powerbi.extensibility.IColorPalette;
type ModalDialogResult = powerbi.extensibility.visual.ModalDialogResult;
type VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
type DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
type DataViewScopeIdentity = powerbi.DataViewScopeIdentity;
type DataViewValueColumns = powerbi.DataViewValueColumns;
type TooltipShowOptions = powerbi.extensibility.TooltipShowOptions;

import { select } from "d3-selection";
import { type settingsValueType as spcDefaultSettingsType } from "../PowerBI-SPC/src/settings";
import { type settingsValueType as funnelDefaultSettingsType } from "../PowerBI-Funnels/src/settings";

type AggregatableValue = number | string | null | undefined;

/** Matches PowerBI's IPromise well enough for the `.then(...)` calls made on it here. */
function resolved<T>(value: T): IPromise<T> {
  return {
    then: (onFulfilled: (value: T) => unknown) => onFulfilled(value),
  } as unknown as IPromise<T>;
}

const tooltipFontSize = 12;
const tooltipLineHeight = 16;
const tooltipPadding = 6;

// Mirrors controlcharts' interactiveUtils.js, minus crosstalk: local state instead of a group.
function makeTooltipService(element: HTMLElement): powerbi.extensibility.ITooltipService {
  function tooltipGroup() {
    const svg = element.querySelector("svg");
    if (!svg) return null;
    const svgSelection = select(svg as SVGSVGElement);
    const existing = svgSelection.select<SVGGElement>(".chart-tooltip-group");
    return existing.empty()
      ? svgSelection
          .append("g")
          .attr("class", "chart-tooltip-group")
          .style("pointer-events", "none")
      : existing;
  }

  return {
    enabled: () => true,
    move: () => undefined,
    hide: () => {
      tooltipGroup()?.selectAll("*").remove();
    },
    show: (options: TooltipShowOptions) => {
      const svg = element.querySelector("svg") as SVGSVGElement | null;
      const group = tooltipGroup();
      if (!svg || !group) return;

      // Dots report mouseover position in page coordinates; convert to the SVG's own.
      const rect = svg.getBoundingClientRect();
      const svgWidth = Number(svg.getAttribute("width")) || rect.width;
      const svgHeight = Number(svg.getAttribute("height")) || rect.height;
      const x =
        (options.coordinates[0] - window.scrollX - rect.left) *
        (rect.width ? svgWidth / rect.width : 1);
      const y =
        (options.coordinates[1] - window.scrollY - rect.top) *
        (rect.height ? svgHeight / rect.height : 1);

      const textLines = group
        .selectAll<SVGTextElement, powerbi.extensibility.VisualTooltipDataItem>("text")
        .data(options.dataItems)
        .join("text")
        .attr("x", tooltipPadding)
        .attr("y", (_, i) => tooltipPadding + tooltipLineHeight * (i + 0.8))
        .style("font-family", "Segoe UI, Arial, sans-serif")
        .style("font-size", `${tooltipFontSize}px`)
        .style("fill", "#111111")
        .text((d) => (d.displayName ? `${d.displayName}: ${d.value}` : d.value));

      let maxTextLength = 0;
      textLines.each(function () {
        maxTextLength = Math.max(maxTextLength, (this as SVGTextElement).getComputedTextLength());
      });

      const boxWidth = maxTextLength + tooltipPadding * 2;
      const boxHeight = options.dataItems.length * tooltipLineHeight + tooltipPadding;

      group
        .selectAll<SVGRectElement, number>("rect")
        .data([0])
        .join("rect")
        .lower()
        .attr("x", 0)
        .attr("y", 0)
        .attr("width", boxWidth)
        .attr("height", boxHeight)
        .attr("rx", 4)
        .attr("ry", 4)
        .attr("fill", "#ffffff")
        .attr("stroke", "#c9ced8")
        .attr("stroke-width", 1);

      const offsetX = x + 12 + boxWidth > svgWidth ? -(boxWidth + 12) : 12;
      const offsetY = y + boxHeight > svgHeight ? -boxHeight : 0;

      group.attr("transform", `translate(${x + offsetX}, ${y + offsetY})`);
    },
  };
}

function makeSelectionManager(): powerbi.extensibility.ISelectionManager {
  let selectedIds: ISelectionId[] = [];

  return {
    registerOnSelectCallback: () => {},
    getSelectionIds: () => selectedIds,
    hasSelection: () => selectedIds.length > 0,
    showContextMenu: () => resolved({}),
    toggleExpandCollapse: () => resolved({}),
    clear: () => {
      selectedIds = [];
      return resolved({});
    },
    select: (selectionId, multiSelect) => {
      const nextIds = Array.isArray(selectionId) ? selectionId : [selectionId];
      selectedIds = multiSelect ? [...selectedIds, ...nextIds] : nextIds;
      return resolved(selectedIds);
    },
  };
}

function makeConstructorArgs(element: HTMLElement): VisualConstructorOptions {
  return {
    element: element,
    host: {
      createSelectionIdBuilder: () => ({
        withCategory: () => ({
          withCategory: () => ({}) as ISelectionIdBuilder,
          withSeries: () => ({}) as ISelectionIdBuilder,
          withMeasure: () => ({}) as ISelectionIdBuilder,
          withMatrixNode: () => ({}) as ISelectionIdBuilder,
          withTable: () => ({}) as ISelectionIdBuilder,
          createSelectionId: () => ({}) as ISelectionId,
        }),
        withSeries: () => ({}) as ISelectionIdBuilder,
        withMeasure: () => ({}) as ISelectionIdBuilder,
        withMatrixNode: () => ({}) as ISelectionIdBuilder,
        withTable: () => ({}) as ISelectionIdBuilder,
        createSelectionId: () => ({}) as ISelectionId,
      }),
      createSelectionManager: () => makeSelectionManager(),
      colorPalette: {
        isHighContrast: false,
        foreground: { value: "black" },
        foregroundLight: {} as IColorInfo,
        foregroundDark: {} as IColorInfo,
        foregroundNeutralLight: {} as IColorInfo,
        foregroundNeutralDark: {} as IColorInfo,
        foregroundNeutralSecondary: {} as IColorInfo,
        foregroundNeutralSecondaryAlt: {} as IColorInfo,
        foregroundNeutralSecondaryAlt2: {} as IColorInfo,
        foregroundNeutralTertiary: {} as IColorInfo,
        foregroundNeutralTertiaryAlt: {} as IColorInfo,
        foregroundSelected: { value: "black" },
        foregroundButton: {} as IColorInfo,
        background: { value: "white" },
        backgroundLight: {} as IColorInfo,
        backgroundNeutral: {} as IColorInfo,
        backgroundDark: {} as IColorInfo,
        hyperlink: { value: "blue" },
        visitedHyperlink: {} as IColorInfo,
        mapPushpin: {} as IColorInfo,
        shapeStroke: {} as IColorInfo,
        getColor: () => ({}) as IColorInfo,
        reset: () => ({}) as IColorPalette,
      },
      persistProperties: () => {},
      applyJsonFilter: () => {},
      tooltipService: makeTooltipService(element),
      telemetry: {} as powerbi.extensibility.ITelemetryService,
      authenticationService: {} as powerbi.extensibility.IAuthenticationService,
      locale: "",
      hostCapabilities: { allowInteractions: true },
      launchUrl: () => null,
      fetchMoreData: () => false,
      openModalDialog: () => ({}) as IPromise<ModalDialogResult>,
      instanceId: "",
      refreshHostData: () => null,
      createLocalizationManager: () => ({}) as powerbi.extensibility.ILocalizationManager,
      storageService: {} as powerbi.extensibility.ILocalVisualStorageService,
      downloadService: {} as powerbi.extensibility.IDownloadService,
      eventService: {
        renderingStarted: () => {},
        renderingFailed: () => {},
        renderingFinished: () => {},
      },
      switchFocusModeState: () => null,
      hostEnv: {} as powerbi.common.CustomVisualHostEnv,
      displayWarningIcon: () => null,
      licenseManager: {} as powerbi.extensibility.IVisualLicenseManager,
      webAccessService: {} as powerbi.extensibility.IWebAccessService,
      drill: () => null,
      applyCustomSort: () => null,
    },
  };
}

function aggregateColumn(column: AggregatableValue[], aggregation: string): powerbi.PrimitiveValue {
  if (aggregation === "first") {
    return (column[0] ?? null) as powerbi.PrimitiveValue;
  }
  if (aggregation === "last") {
    return (column[column.length - 1] ?? null) as powerbi.PrimitiveValue;
  }

  const numericColumn = column.filter((value): value is number => typeof value === "number");
  if (numericColumn.length === 0) {
    return null;
  }

  if (aggregation === "sum") {
    return numericColumn.reduce((acc: number, val: number) => acc + val, 0);
  }
  if (aggregation === "mean") {
    return numericColumn.reduce((acc: number, val: number) => acc + val, 0) / numericColumn.length;
  }
  if (aggregation === "sd") {
    const mean: number =
      numericColumn.reduce((acc: number, val: number) => acc + val, 0) / numericColumn.length;
    return Math.sqrt(
      numericColumn.reduce((acc: number, val: number) => acc + Math.pow(val - mean, 2), 0) /
        (numericColumn.length - 1)
    );
  }
  if (aggregation === "count") {
    return column.length;
  }
  if (aggregation === "min") {
    return Math.min(...numericColumn);
  }
  if (aggregation === "max") {
    return Math.max(...numericColumn);
  }
  if (aggregation === "median") {
    const sorted = [...numericColumn].sort((a: number, b: number) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }

  throw new Error(`Unsupported aggregation: ${aggregation}`);
}

type rawDataType = Array<{
  categories: string | Date | null;
  numerators: number;
  denominators?: number | undefined;
  xbar_sds?: number | undefined;
  labels?: string | undefined;
  [key: string]: string | number | Date | null | undefined;
}>;

function makeUpdateValues(
  rawData: rawDataType,
  inputSettings: spcDefaultSettingsType | funnelDefaultSettingsType,
  aggregations: Record<string, string>,
  categoryIsTemporal = false
): VisualUpdateOptions {
  const dataGrouped: Array<{ category: string | Date | null; rows: rawDataType }> = [];
  const groupIndexes = new Map<string, number>();

  rawData.forEach((row) => {
    const categoryKey =
      row.categories instanceof Date && Number.isFinite(row.categories.getTime())
        ? `date:${row.categories.getTime()}`
        : String(row.categories ?? "");
    const existingIndex = groupIndexes.get(categoryKey);

    if (existingIndex === undefined) {
      groupIndexes.set(categoryKey, dataGrouped.length);
      dataGrouped.push({ category: row.categories, rows: [row] });
    } else {
      dataGrouped[existingIndex].rows.push(row);
    }
  });

  const categories: DataViewCategoryColumn = {
    source: {
      displayName: "categories",
      roles: { key: true },
      type: categoryIsTemporal
        ? { temporal: {} as powerbi.TemporalTypeDescriptor }
        : { text: true },
    },
    values: [],
    objects: [],
    identity: [],
  };

  const valueNames: string[] = Object.keys(rawData[0]).filter((k) => !["categories"].includes(k));

  var values = valueNames.map((name) => ({
    source: { roles: { [name]: true } },
    values: new Array<powerbi.PrimitiveValue>(),
  }));

  for (const group of dataGrouped) {
    categories.values.push(group.category);
    categories.objects!.push(inputSettings as powerbi.DataViewObjects);
    categories.identity!.push({} as DataViewScopeIdentity);

    for (var i = 0; i < valueNames.length; i++) {
      var name = valueNames[i];
      var aggregatedValue = aggregateColumn(
        group.rows.map((dataRow) => dataRow[name] as AggregatableValue),
        aggregations[name]
      );
      values[i].values.push(aggregatedValue);
    }
  }

  (values as any).grouped = [];

  return {
    dataViews: [
      {
        metadata: {} as powerbi.DataViewMetadata,
        categorical: {
          categories: [categories],
          values: values as DataViewValueColumns,
        },
      },
    ],
    viewport: {} as powerbi.IViewport,
    type: 2,
  };
}

export { makeConstructorArgs, makeUpdateValues };
