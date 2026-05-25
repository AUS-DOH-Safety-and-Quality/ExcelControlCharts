import type powerbi from "powerbi-visuals-api";
type VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
type ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
type IPromise<T> = powerbi.IPromise<T>;
type ISelectionId = powerbi.visuals.ISelectionId;
type IColorInfo = powerbi.IColorInfo;
type IColorPalette = powerbi.extensibility.IColorPalette;
type VisualObjectInstancesToPersist = powerbi.VisualObjectInstancesToPersist;
type DialogOpenOptions = powerbi.extensibility.visual.DialogOpenOptions;
type ModalDialogResult = powerbi.extensibility.visual.ModalDialogResult;
type IFilter = powerbi.IFilter;
type CustomVisualApplyCustomSortArgs = powerbi.extensibility.visual.CustomVisualApplyCustomSortArgs;
type VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions
type DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
type DataViewValueColumns = powerbi.DataViewValueColumns;

import { type settingsValueType as spcDefaultSettingsType } from "../PowerBI-SPC/src/settings";
import { type settingsValueType as funnelDefaultSettingsType } from "../PowerBI-Funnels/src/settings";

function makeConstructorArgs(element: HTMLElement): VisualConstructorOptions {
  return {
    element: element,
    host: {
      createSelectionIdBuilder: () => ({
        withCategory: () => ({
          withCategory: () => ({} as ISelectionIdBuilder),
          withSeries: () => ({} as ISelectionIdBuilder),
          withMeasure: () => ({} as ISelectionIdBuilder),
          withMatrixNode: () => ({} as ISelectionIdBuilder),
          withTable: () => ({} as ISelectionIdBuilder),
          createSelectionId: () => ({} as ISelectionId)
        }),
        withSeries: () => ({} as ISelectionIdBuilder),
        withMeasure: () => ({} as ISelectionIdBuilder),
        withMatrixNode: () => ({} as ISelectionIdBuilder),
        withTable: () => ({} as ISelectionIdBuilder),
        createSelectionId: () => ({} as ISelectionId )
      }),
      createSelectionManager: () => ({
        registerOnSelectCallback: () => {},
        getSelectionIds: () => [],
        showContextMenu: () => ({} as IPromise<{}>),
        clear: () => ({} as IPromise<{}>),
        toggleExpandCollapse: () => ({} as IPromise<{}>),
        select: () => ({} as IPromise<ISelectionId[]>),
        hasSelection: () => false
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
        /* background variants*/
        background: { value: "white" },
        backgroundLight: {} as IColorInfo,
        backgroundNeutral: {} as IColorInfo,
        backgroundDark: {} as IColorInfo,
        /* specific purpose colors*/
        hyperlink: { value: "blue" },
        visitedHyperlink: {} as IColorInfo,
        mapPushpin: {} as IColorInfo,
        shapeStroke: {} as IColorInfo,
        getColor: () => ({} as IColorInfo),
        reset: () => ({} as IColorPalette)
      },
      persistProperties: (changes: VisualObjectInstancesToPersist) => {},
      applyJsonFilter: (filter: IFilter[] | IFilter, objectName: string, propertyName: string, action: powerbi.FilterAction) => {},
      tooltipService: {
        show: () => null,
        hide: () => null,
        enabled: () => true,
        move: () => null
      },
      telemetry: {} as powerbi.extensibility.ITelemetryService,
      authenticationService: {} as powerbi.extensibility.IAuthenticationService,
      locale: "",
      hostCapabilities: {} as powerbi.extensibility.HostCapabilities,
      launchUrl: (url: string) => null,
      fetchMoreData: (aggregateSegments?: boolean) => false,
      openModalDialog: (dialogId: string, options?: DialogOpenOptions, initialState?: object) => ({} as IPromise<ModalDialogResult>),
      instanceId: "",
      refreshHostData: () => null,
      createLocalizationManager: () => ({} as powerbi.extensibility.ILocalizationManager),
      storageService: {} as powerbi.extensibility.ILocalVisualStorageService,
      downloadService: {} as powerbi.extensibility.IDownloadService,
      eventService: {
        renderingStarted: () => {},
        renderingFailed: () => {},
        renderingFinished: () => {}
      },
      switchFocusModeState: (on: boolean) => null,
      hostEnv: {} as powerbi.common.CustomVisualHostEnv,
      displayWarningIcon: (hoverText: string, detailedText: string) => null,
      licenseManager: {} as powerbi.extensibility.IVisualLicenseManager,
      webAccessService: {} as powerbi.extensibility.IWebAccessService,
      drill: (args: powerbi.DrillArgs) => null,
      applyCustomSort: (args: CustomVisualApplyCustomSortArgs) => null
    }
  }
}

function aggregateColumn(column: number[], aggregation: string): number {
  if (aggregation === "sum") {
    return column.reduce((acc: number, val: number) => acc + val, 0);
  } else if (aggregation === "mean") {
    return column.reduce((acc: number, val: number) => acc + val, 0) / column.length;
  } else if (aggregation === "sd") {
    const mean: number = column.reduce((acc: number, val: number) => acc + val, 0) / column.length;
    return Math.sqrt(column.reduce((acc: number, val: number) => acc + Math.pow(val - mean, 2), 0) / (column.length - 1));
  } else if (aggregation === "count") {
    return column.length;
  } else if (aggregation === "min") {
    return Math.min(...column);
  } else if (aggregation === "max") {
    return Math.max(...column);
  } else if (aggregation === "median") {
    const sorted = [...column].sort((a: number, b: number) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  } else if (aggregation === "first") {
    return column[0];
  } else if (aggregation === "last") {
    return column[column.length - 1];
  } else {
    throw new Error(`Unsupported aggregation: ${aggregation}`);
  }
}

type rawDataType = Array<{
  categories: string | Date,
  numerators: number,
  denominators?: number | undefined
  xbar_sds?: number | undefined
}>;

  // Custom groupBy implementation to replace Object.groupBy
function groupBy(array: any[], keyFn: (item: any) => any): { [key: string]: any[] } {
  return array.reduce((result, item) => {
    const key = keyFn(item);
    if (!result[key]) result[key] = [];
    result[key].push(item);
    return result;
  }, {});
}

function makeUpdateValues(rawData: rawDataType,
                          inputSettings: spcDefaultSettingsType | funnelDefaultSettingsType,
                          aggregations: Record<string, string>): VisualUpdateOptions {
  const dataGrouped = groupBy(rawData, d => d.categories);
  Object.freeze(dataGrouped);

  const categories: DataViewCategoryColumn = {
    source: {
      displayName: "categories",
      roles: {"key": true},
      type: { temporal: {} as powerbi.TemporalTypeDescriptor }
    },
    values: [],
    objects: []
  };

  const valueNames: string[] = Object.keys(rawData[0]).filter(k => !["categories"].includes(k));

  var values = valueNames.map(name => ({
    source: { roles: {[name]: true} },
    values: new Array<powerbi.PrimitiveValue>()
  }));

  for (var category in dataGrouped) {
    categories.values.push(category);
    categories.objects!.push(inputSettings as powerbi.DataViewObjects);

    for (var i = 0; i < valueNames.length; i++) {
      var name = valueNames[i];
      var aggregatedValue = aggregateColumn(dataGrouped[category].map(dataRow => dataRow[name]), aggregations[name]);
      values[i].values.push(aggregatedValue);
    }
  }

  (values as any).grouped = []

  return {
    dataViews: [{
      metadata: {} as powerbi.DataViewMetadata,
      categorical: {
        categories: [ categories ],
        values: values as DataViewValueColumns
      }
    }],
    viewport: {} as powerbi.IViewport,
    type: 2 // Update type == 'data'
  };
}

export { makeConstructorArgs, makeUpdateValues };
