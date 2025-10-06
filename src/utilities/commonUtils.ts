import type powerbi from "../../PowerBI-SPC/node_modules/powerbi-visuals-api";
type VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;

export default function makeConstructorArgs(element: HTMLElement): VisualConstructorOptions {
  return {
    element: element,
    host: {
      createSelectionIdBuilder: () => ({
        withCategory: () => ({
          withCategory: () => null,
          withSeries: () => null,
          withMeasure: () => null,
          withMatrixNode: () => null,
          withTable: () => null,
          createSelectionId: () => null
        }),
        withSeries: () => null,
        withMeasure: () => null,
        withMatrixNode: () => null,
        withTable: () => null,
        createSelectionId: () => null
      }),
      createSelectionManager: () => ({
        registerOnSelectCallback: () => {},
        getSelectionIds: () => [],
        showContextMenu: () => null,
        clear: () => null,
        toggleExpandCollapse: () => null,
        select: () => null,
        hasSelection: () => false
      }),
      colorPalette: {
        isHighContrast: false,
        foreground: { value: "black" },
        foregroundLight: null,
        foregroundDark: null,
        foregroundNeutralLight: null,
        foregroundNeutralDark: null,
        foregroundNeutralSecondary: null,
        foregroundNeutralSecondaryAlt: null,
        foregroundNeutralSecondaryAlt2: null,
        foregroundNeutralTertiary: null,
        foregroundNeutralTertiaryAlt: null,
        foregroundSelected: { value: "black" },
        foregroundButton: null,
        /* background variants*/
        background: { value: "white" },
        backgroundLight: null,
        backgroundNeutral: null,
        backgroundDark: null,
        /* specific purpose colors*/
        hyperlink: { value: "blue" },
        visitedHyperlink: null,
        mapPushpin: null,
        shapeStroke: null,
        getColor: () => null,
        reset: () => null
      },
      persistProperties: (changes) => null,
      applyJsonFilter: (filter, objectName, propertyName, action) => null,
      tooltipService: {
        show: () => null,
        hide: () => null,
        enabled: () => true,
        move: () => null
      },
      telemetry: null,
      authenticationService: null,
      locale: null,
      hostCapabilities: null,
      launchUrl: (url: string) => null,
      fetchMoreData: (aggregateSegments?: boolean) => null,
      openModalDialog: (dialogId: string, options?, initialState?) => null,
      instanceId: null,
      refreshHostData: () => null,
      createLocalizationManager: () => null,
      storageService: null,
      downloadService: null,
      eventService: {
        renderingStarted: () => {},
        renderingFailed: () => {},
        renderingFinished: () => {}
      },
      switchFocusModeState: (on: boolean) => null,
      hostEnv: null,
      displayWarningIcon: (hoverText: string, detailedText: string) => null,
      licenseManager: null,
      webAccessService: null,
      drill: (args) => null,
      applyCustomSort: (args) => null,
      acquireAADTokenService: null,
      setCanDrill: null,
      storageV2Service: null,
      subSelectionService: null,
      createOpaqueUtils: null,
    }
  }
}

const aggregateColumn = function(column, aggregation) {
  if (aggregation === "sum") {
    return column.reduce((acc, val) => acc + val, 0);
  } else if (aggregation === "mean") {
    return column.reduce((acc, val) => acc + val, 0) / column.length;
  } else if (aggregation === "sd") {
    var mean = column.reduce((acc, val) => acc + val, 0) / column.length;
    return Math.sqrt(column.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / (column.length - 1));
  } else if (aggregation === "count") {
    return column.length;
  } else if (aggregation === "min") {
    return Math.min(...column);
  } else if (aggregation === "max") {
    return Math.max(...column);
  } else if (aggregation === "median") {
    var sorted = [...column].sort((a, b) => a - b);
    var mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  } else if (aggregation === "first") {
    return column[0];
  } else if (aggregation === "last") {
    return column[column.length - 1];
  } else {
    throw new Error(`Unsupported aggregation: ${aggregation}`);
  }
}

function makeUpdateValues(rawData, inputSettings, aggregations, crosstalkFilters) {
  if (crosstalkFilters) {
    rawData = rawData.filter(d => crosstalkFilters.includes(d.crosstalkIdentities));
  }
  // Custom groupBy implementation to replace Object.groupBy
  function groupBy(array, keyFn) {
    return array.reduce((result, item) => {
      const key = keyFn(item);
      if (!result[key]) result[key] = [];
      result[key].push(item);
      return result;
    }, {});
  }

  var dataGrouped = groupBy(rawData, d => d.categories);
  Object.freeze(dataGrouped);
  var identitiesGrouped = [];
  for (var group in dataGrouped) {
    identitiesGrouped.push([group, dataGrouped[group].map(d => d.crosstalkIdentities)]);
  }

  var args = {
    categories: [{
      source: { roles: {"key": true}, type: { temporal: { underlyingType: 519 } } },
      values: [],
      objects: []
    }],
    values: [],
    crosstalkIdentities: Object.fromEntries(identitiesGrouped)
  };

  var valueNames = Object.keys(rawData[0]).filter(k => !["categories", "crosstalkIdentities"].includes(k));

  args.values = valueNames.map(name => ({
    source: { roles: {[name]: true} },
    values: []
  }));

  for (var category in dataGrouped) {
    args.categories[0].values.push(category);
    args.categories[0].objects.push(inputSettings);

    for (var i = 0; i < valueNames.length; i++) {
      var name = valueNames[i];
      var aggregatedValue = aggregateColumn(dataGrouped[category].map(dataRow => dataRow[name]), aggregations[name]);
      args.values[i].values.push(aggregatedValue);
    }
  }

  return {
    dataViews: [{
      categorical: {
        categories: args.categories,
        values: args.values
      }
    }],
    crosstalkIdentities: args.crosstalkIdentities
  };
}
