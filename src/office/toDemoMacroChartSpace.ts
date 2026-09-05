import {
  getPlotChartGroup,
  NO_AXIS_PLOT_KEYS,
  type OfficeOpenChart,
  type OfficePlotChartKey,
} from "./OfficeOpenTypes.js";

/**
 * Subset of DemoMacro `@office-open/core` `ChartSpaceOptions`.
 * Host geometry (x/y/cx/cy) is omitted so the object stays application-agnostic.
 */
export interface DemoMacroChartSpace {
  type: DemoMacroChartType;
  title?: string;
  categories?: string[];
  series: Array<{
    name: string;
    values?: number[];
    xValues?: number[];
    yValues?: number[];
    bubbleSize?: number[];
  }>;
  grouping?: "clustered" | "standard" | "stacked" | "percentStacked";
  showLegend?: boolean;
  legendPosition?: "r" | "l" | "t" | "b";
  threeD?: boolean;
  lang?: string;
  date1904?: boolean;
  style?: number;
}

export type DemoMacroChartType =
  | "column"
  | "bar"
  | "line"
  | "pie"
  | "area"
  | "scatter"
  | "bubble"
  | "doughnut"
  | "radar";

const TYPE_BY_PLOT_KEY: Record<OfficePlotChartKey, DemoMacroChartType> = {
  barChart: "column",
  bar3DChart: "column",
  lineChart: "line",
  line3DChart: "line",
  pieChart: "pie",
  pie3DChart: "pie",
  areaChart: "area",
  area3DChart: "area",
  scatterChart: "scatter",
  doughnutChart: "doughnut",
  radarChart: "radar",
  bubbleChart: "bubble",
};

/**
 * Project an Office Open JSON chart into DemoMacro ChartSpaceOptions.
 * Downstream PPTX/XLSX/DOCX generators can embed this under a host-specific
 * `chart` child together with placement properties they own.
 */
export function toDemoMacroChartSpace(chart: OfficeOpenChart): DemoMacroChartSpace {
  const { key, group } = getPlotChartGroup(chart.chartSpace.chart.plotArea);
  let type = TYPE_BY_PLOT_KEY[key];
  if ((key === "barChart" || key === "bar3DChart") && "barDir" in group && group.barDir.val === "bar") {
    type = "bar";
  }

  const first = group.ser[0];
  const categories = first?.cat?.map((item) => String(item));
  const series = group.ser.map((item) => ({
    name: item.tx,
    values: item.val,
    xValues: item.xVal,
    yValues: item.yVal,
    bubbleSize: item.bubbleSize,
  }));

  return {
    type,
    title: chart.chartSpace.chart.title?.tx,
    categories,
    series,
    grouping: "grouping" in group ? group.grouping.val : undefined,
    showLegend: Boolean(chart.chartSpace.chart.legend),
    legendPosition: chart.chartSpace.chart.legend?.legendPos.val,
    threeD: key.endsWith("3DChart"),
    lang: chart.chartSpace.lang.val,
    date1904: chart.chartSpace.date1904.val,
    style: chart.chartSpace.style?.val,
  };
}

export function isPieLikePlotKey(key: OfficePlotChartKey): boolean {
  return NO_AXIS_PLOT_KEYS.has(key);
}
