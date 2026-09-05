/**
 * DemoMacro-compatible Office Open JSON chart types.
 *
 * These objects are an intermediate ChartML representation: object trees, not
 * XML strings, and not a PPTX/XLSX/DOCX package. Downstream Office generators
 * consume this structure for PowerPoint, Excel, Word, editors, and AI pipelines.
 */

export const OFFICE_OPEN_CHART_KIND = "chart" as const;

export const OFFICE_PLOT_CHART_KEYS = [
  "barChart",
  "lineChart",
  "pieChart",
  "areaChart",
  "scatterChart",
  "doughnutChart",
  "radarChart",
  "bubbleChart",
  "bar3DChart",
  "line3DChart",
  "pie3DChart",
  "area3DChart",
] as const;

export type OfficePlotChartKey = (typeof OFFICE_PLOT_CHART_KEYS)[number];

export type OfficeBarDirection = "bar" | "col";
export type OfficeGrouping = "clustered" | "stacked" | "percentStacked" | "standard";
export type OfficeLegendPosition = "r" | "l" | "t" | "b";
export type OfficeAxisPosition = "b" | "l" | "r" | "t";
export type OfficeOrientation = "minMax" | "maxMin";
export type OfficeTickLabelPosition = "nextTo" | "high" | "low" | "none";
export type OfficeDisplayBlanksAs = "gap" | "span" | "zero";
export type OfficeScatterStyle = "marker" | "line" | "lineMarker" | "smooth" | "smoothMarker";
export type OfficeRadarStyle = "standard" | "marker" | "filled";

/** CT_Boolean / CT_OnOff encoded as `{ val }`. */
export interface OfficeOnOff {
  val: boolean;
}

/** CT_UnsignedInt encoded as `{ val }`. */
export interface OfficeUnsignedInt {
  val: number;
}

/** CT_* with a string/enum `val` attribute. */
export interface OfficeVal<T extends string | number = string> {
  val: T;
}

export interface OfficeSeries {
  idx: OfficeUnsignedInt;
  order: OfficeUnsignedInt;
  /** Series title (`c:tx`). Simplified DemoMacro form: a string, not XML. */
  tx: string;
  /** Category labels (`c:cat`). */
  cat?: Array<string | number>;
  /** Numeric values (`c:val`). */
  val?: number[];
  /** Scatter / bubble X values (`c:xVal`). */
  xVal?: number[];
  /** Scatter / bubble Y values (`c:yVal`). */
  yVal?: number[];
  /** Bubble sizes (`c:bubbleSize`). */
  bubbleSize?: number[];
  invertIfNegative?: OfficeOnOff;
  explosion?: OfficeUnsignedInt;
  spPr?: OfficeShapeProperties;
}

export interface OfficeShapeProperties {
  solidFill?: {
    srgbClr: OfficeVal<string>;
  };
}

export interface OfficeLayout {
  manualLayout?: Record<string, never>;
}

export interface OfficeDataLabels {
  showLegendKey?: OfficeOnOff;
  showVal?: OfficeOnOff;
  showCatName?: OfficeOnOff;
  showSerName?: OfficeOnOff;
  showPercent?: OfficeOnOff;
  showBubbleSize?: OfficeOnOff;
}

export interface OfficeBarChart {
  barDir: OfficeVal<OfficeBarDirection>;
  grouping: OfficeVal<OfficeGrouping>;
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  gapWidth?: OfficeUnsignedInt;
  overlap?: OfficeVal<number>;
  axId: [OfficeUnsignedInt, OfficeUnsignedInt];
  dLbls?: OfficeDataLabels;
}

export interface OfficeLineChart {
  grouping: OfficeVal<OfficeGrouping>;
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  marker?: OfficeOnOff;
  smooth?: OfficeOnOff;
  axId: [OfficeUnsignedInt, OfficeUnsignedInt];
  dLbls?: OfficeDataLabels;
}

export interface OfficeAreaChart {
  grouping: OfficeVal<OfficeGrouping>;
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  axId: [OfficeUnsignedInt, OfficeUnsignedInt];
  dLbls?: OfficeDataLabels;
}

export interface OfficePieChart {
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  firstSliceAng?: OfficeUnsignedInt;
  dLbls?: OfficeDataLabels;
}

export interface OfficeDoughnutChart extends OfficePieChart {
  holeSize: OfficeUnsignedInt;
}

export interface OfficeScatterChart {
  scatterStyle: OfficeVal<OfficeScatterStyle>;
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  axId: [OfficeUnsignedInt, OfficeUnsignedInt];
  dLbls?: OfficeDataLabels;
}

export interface OfficeRadarChart {
  radarStyle: OfficeVal<OfficeRadarStyle>;
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  axId: [OfficeUnsignedInt, OfficeUnsignedInt];
  dLbls?: OfficeDataLabels;
}

export interface OfficeBubbleChart {
  varyColors: OfficeOnOff;
  ser: OfficeSeries[];
  bubble3D?: OfficeOnOff;
  axId: [OfficeUnsignedInt, OfficeUnsignedInt];
  dLbls?: OfficeDataLabels;
}

export type OfficeChartGroup =
  | OfficeBarChart
  | OfficeLineChart
  | OfficeAreaChart
  | OfficePieChart
  | OfficeDoughnutChart
  | OfficeScatterChart
  | OfficeRadarChart
  | OfficeBubbleChart;

export interface OfficeScaling {
  orientation: OfficeVal<OfficeOrientation>;
  min?: OfficeVal<number>;
  max?: OfficeVal<number>;
}

export interface OfficeAxisTitle {
  tx: string;
}

export interface OfficeCategoryAxis {
  axId: OfficeUnsignedInt;
  scaling: OfficeScaling;
  delete?: OfficeOnOff;
  axPos: OfficeVal<OfficeAxisPosition>;
  majorTickMark?: OfficeVal<"in" | "out" | "cross" | "none">;
  minorTickMark?: OfficeVal<"in" | "out" | "cross" | "none">;
  tickLblPos: OfficeVal<OfficeTickLabelPosition>;
  crossAx: OfficeUnsignedInt;
  title?: OfficeAxisTitle;
}

export interface OfficeValueAxis {
  axId: OfficeUnsignedInt;
  scaling: OfficeScaling;
  delete?: OfficeOnOff;
  axPos: OfficeVal<OfficeAxisPosition>;
  majorGridlines?: Record<string, never>;
  numFmt?: { formatCode: string; sourceLinked: boolean };
  majorTickMark?: OfficeVal<"in" | "out" | "cross" | "none">;
  minorTickMark?: OfficeVal<"in" | "out" | "cross" | "none">;
  tickLblPos: OfficeVal<OfficeTickLabelPosition>;
  crossAx: OfficeUnsignedInt;
  title?: OfficeAxisTitle;
}

export interface OfficePlotAreaBase {
  layout?: OfficeLayout;
  catAx?: OfficeCategoryAxis[];
  valAx?: OfficeValueAxis[];
}

export type OfficePlotArea = OfficePlotAreaBase &
  (
    | { barChart: OfficeBarChart }
    | { bar3DChart: OfficeBarChart }
    | { lineChart: OfficeLineChart }
    | { line3DChart: OfficeLineChart }
    | { pieChart: OfficePieChart }
    | { pie3DChart: OfficePieChart }
    | { areaChart: OfficeAreaChart }
    | { area3DChart: OfficeAreaChart }
    | { scatterChart: OfficeScatterChart }
    | { doughnutChart: OfficeDoughnutChart }
    | { radarChart: OfficeRadarChart }
    | { bubbleChart: OfficeBubbleChart }
  );

export interface OfficeLegend {
  legendPos: OfficeVal<OfficeLegendPosition>;
  overlay?: OfficeOnOff;
}

export interface OfficeChartTitle {
  tx: string;
  overlay?: OfficeOnOff;
}

export interface OfficeChart {
  title?: OfficeChartTitle;
  autoTitleDeleted: OfficeOnOff;
  plotArea: OfficePlotArea;
  legend?: OfficeLegend;
  plotVisOnly: OfficeOnOff;
  dispBlanksAs: OfficeVal<OfficeDisplayBlanksAs>;
  showDLblsOverMax?: OfficeOnOff;
}

export interface OfficeChartSpace {
  date1904: OfficeOnOff;
  lang: OfficeVal<string>;
  roundedCorners: OfficeOnOff;
  style?: OfficeUnsignedInt;
  chart: OfficeChart;
}

/**
 * Root Office Open JSON chart object (DemoMacro-compatible ChartML tree).
 *
 * This is the library output. It is not an XML document and not a ZIP package.
 */
export interface OfficeOpenChart {
  type: typeof OFFICE_OPEN_CHART_KIND;
  chartSpace: OfficeChartSpace;
}

/** Alias kept for callers who want the JSON-document name. */
export type OfficeOpenChartJson = OfficeOpenChart;

export interface PlotChartGroupRef {
  key: OfficePlotChartKey;
  group: OfficeChartGroup;
}

export function getPlotChartGroup(plotArea: OfficePlotArea): PlotChartGroupRef {
  for (const key of OFFICE_PLOT_CHART_KEYS) {
    const group = plotArea[key as keyof OfficePlotArea];
    if (group && typeof group === "object" && "ser" in group) {
      return { key, group: group as OfficeChartGroup };
    }
  }
  throw new Error("plotArea is missing a ChartML chart-type element");
}

export const NO_AXIS_PLOT_KEYS: ReadonlySet<OfficePlotChartKey> = new Set([
  "pieChart",
  "pie3DChart",
  "doughnutChart",
]);
