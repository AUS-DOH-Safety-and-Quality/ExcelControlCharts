/**
 * Page integration for the web build of the taskpane. `office-shim` supplies the
 * Excel API; this module moves the chart the taskpane draws into the hosting
 * page's own column and keeps it refreshing as the inputs change.
 *
 * Importing the shim for its side effects also guarantees it installs first.
 */
import { afterReady, resolveHost } from "./office-shim";

/** Matches the padding `previewPlot` subtracts from its container's measured size. */
const previewPadding = 16;
const renderDebounceMs = 120;

/**
 * taskpane.ts already re-previews on its own for the category selector, the chart
 * type toggle and every settings field. These are the inputs it leaves uncovered.
 */
const uncoveredSelectors = ["numerator-selector", "denominator-selector", "sd-selector"];

/** Parks the panel's own preview and action controls off-screen, still measurable. */
const embedStyles = `
  .preview-card, .actions {
    position: absolute !important;
    left: -10000px;
    top: 0;
    width: 1px;
    height: 1px;
    overflow: hidden;
  }
`;

function debounce(callback: () => void, delay: number): () => void {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return () => {
    clearTimeout(timer);
    timer = setTimeout(callback, delay);
  };
}

afterReady(() => {
  const chartHost = resolveHost().getChartHost();
  const previewContainer = document.getElementById("preview-container");
  const previewButton = document.getElementById("preview-plot") as HTMLButtonElement | null;

  if (!previewContainer || !previewButton) {
    return;
  }

  const style = document.createElement("style");
  style.textContent = embedStyles;
  document.head.appendChild(style);

  // The visual's containers were just appended to the panel's preview area; move
  // them into the page so the chart renders at full size next to the spreadsheet.
  for (const child of Array.from(previewContainer.children)) {
    chartHost.appendChild(document.adoptNode(child));
  }

  /**
   * `previewPlot` sizes the chart from the preview container's bounding box, so
   * keep that off-screen box matching the on-page column the chart now lives in.
   */
  const syncSize = () => {
    previewContainer.style.width = `${chartHost.clientWidth + previewPadding}px`;
    previewContainer.style.height = `${chartHost.clientHeight + previewPadding}px`;
  };

  // Clicking the panel's own (now hidden) preview button keeps every code path for
  // building a chart in taskpane.ts, rather than duplicating any of it here.
  const render = () => {
    syncSize();
    if (previewButton.disabled) {
      return;
    }
    previewButton.click();
    chartHost.dataset.rendered = "true";
  };

  const scheduleRender = debounce(render, renderDebounceMs);

  for (const id of uncoveredSelectors) {
    document.getElementById(id)?.addEventListener("change", scheduleRender);
  }

  // Re-render when the column resizes, including when the panel is collapsed.
  const parentWindow = (window.parent as typeof window | undefined) ?? window;
  new parentWindow.ResizeObserver(scheduleRender).observe(chartHost);

  // Lets the page redraw the chart when the spreadsheet itself is edited.
  window.__refreshChart = scheduleRender;
  syncSize();
});
