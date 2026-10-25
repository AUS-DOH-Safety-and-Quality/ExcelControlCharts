const preferenceKey = "excel-control-charts:save-local-data";

const managedStorageKeys = [
  "excel-control-charts:workbook",
  "excel-control-charts:workbook-width",
  "excel-control-charts:panel-width",
  "excel-control-charts:panel-collapsed",
  "excel-control-charts-theme",
];

export function isLocalStorageEnabled(): boolean {
  try {
    return window.localStorage.getItem(preferenceKey) !== "false";
  } catch {
    return false;
  }
}

export function setLocalStorageEnabled(enabled: boolean): void {
  try {
    if (!enabled) {
      managedStorageKeys.forEach((key) => window.localStorage.removeItem(key));
    }
    window.localStorage.setItem(preferenceKey, String(enabled));
  } catch {
    // Storage can be unavailable in private browsing or disabled by browser policy.
  }
}
