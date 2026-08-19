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

import { type settingsValueType as spcDefaultSettingsType } from "../PowerBI-SPC/src/settings";
import { type settingsValueType as funnelDefaultSettingsType } from "../PowerBI-Funnels/src/settings";

type AggregatableValue = number | string | null | undefined;

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
      createSelectionManager: () => ({
        registerOnSelectCallback: () => {},
        getSelectionIds: () => [],
        showContextMenu: () => ({}) as IPromise<{}>,
        clear: () => ({}) as IPromise<{}>,
        toggleExpandCollapse: () => ({}) as IPromise<{}>,
        select: () => ({}) as IPromise<ISelectionId[]>,
        hasSelection: () => false,
      }),
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
      tooltipService: {
        show: () => null,
        hide: () => null,
        enabled: () => true,
        move: () => null,
      },
      telemetry: {} as powerbi.extensibility.ITelemetryService,
      authenticationService: {} as powerbi.extensibility.IAuthenticationService,
      locale: "",
      hostCapabilities: {} as powerbi.extensibility.HostCapabilities,
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
