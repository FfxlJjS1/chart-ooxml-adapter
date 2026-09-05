/**
 * Host-agnostic chart model produced by source adapters.
 *
 * This is the only input the Office Open JSON builder understands. It is not
 * tied to PowerPoint, Excel, or Word packaging.
 */

export const UNIVERSAL_CHART_TYPES = [
  "bar",
  "column",
  "line",
  "area",
  "pie",
  "doughnut",
  "scatter",
  "radar",
  "bubble",
] as const;

export type UniversalChartType = (typeof UNIVERSAL_CHART_TYPES)[number];

export const UNIVERSAL_GROUPINGS = [
  "clustered",
  "standard",
  "stacked",
  "percentStacked",
] as const;

export type UniversalGrouping = (typeof UNIVERSAL_GROUPINGS)[number];

export const UNIVERSAL_LEGEND_POSITIONS = ["right", "left", "top", "bottom"] as const;

export type UniversalLegendPosition = (typeof UNIVERSAL_LEGEND_POSITIONS)[number];

export const CHART_SOURCES = [
  "universal",
  "echarts",
  "highcharts",
  "chartjs",
  "amcharts",
] as const;

export type ChartSource = (typeof CHART_SOURCES)[number];

export interface UniversalSeries {
  /** Series display name (`c:tx`). */
  name: string;
  /** Category-chart values (`c:val`). */
  values?: number[];
  /** Scatter / bubble X values (`c:xVal`). */
  xValues?: number[];
  /** Scatter / bubble Y values (`c:yVal`). */
  yValues?: number[];
  /** Bubble sizes (`c:bubbleSize`). */
  bubbleSizes?: number[];
  /** Optional series color as `#RRGGBB` or `RRGGBB`. */
  color?: string;
}

export interface UniversalAxis {
  title?: string;
  min?: number;
  max?: number;
}

export interface UniversalLegend {
  show?: boolean;
  position?: UniversalLegendPosition;
}

export interface UniversalChartSourceMeta {
  library: ChartSource;
}

export interface UniversalChartModel {
  type: UniversalChartType;
  title?: string;
  categories: string[];
  series: UniversalSeries[];
  grouping?: UniversalGrouping;
  legend?: UniversalLegend;
  xAxis?: UniversalAxis;
  yAxis?: UniversalAxis;
  threeD?: boolean;
  source?: UniversalChartSourceMeta;
}

export const CATEGORY_CHART_TYPES: ReadonlySet<UniversalChartType> = new Set([
  "bar",
  "column",
  "line",
  "area",
  "pie",
  "doughnut",
  "radar",
]);

export const CARTESIAN_CHART_TYPES: ReadonlySet<UniversalChartType> = new Set([
  "bar",
  "column",
  "line",
  "area",
  "scatter",
  "radar",
  "bubble",
]);

export const SCATTER_LIKE_TYPES: ReadonlySet<UniversalChartType> = new Set([
  "scatter",
  "bubble",
]);
