const urlDev = "https://localhost:3100/";

const urlDeployed = "https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/";

export const forDeployment = (manifestXml: string) => manifestXml.replaceAll(urlDev, urlDeployed);
