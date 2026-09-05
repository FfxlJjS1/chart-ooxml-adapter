import { toUniversalChartModel } from "./adapters/index.js";
import type { ConvertOptions } from "./adapters/types.js";
import { buildOfficeOpenChart } from "./office/OfficeOpenChartBuilder.js";
import { toOfficeOpenChartJson } from "./office/OfficeOpenChartJson.js";
import type { OfficeOpenChart } from "./office/OfficeOpenTypes.js";
import { assertOfficeOpenChart } from "./office/validateOfficeOpenChart.js";

export interface ConvertChartOptions extends ConvertOptions {
  pretty?: boolean;
}

/**
 * Convert a chart library config or UniversalChartModel into a DemoMacro
 * compatible Office Open JSON chart object.
 *
 * This is the public boundary of chart-ooxml-adapter:
 * Input  = Chart.js / ECharts / Highcharts / amCharts / UniversalChartModel
 * Output = Office Open JSON (ChartML object tree)
 *
 * The function does not create ZIP packages, relationships, media, or XML.
 */
export function convertChart(input: unknown, options?: ConvertChartOptions): OfficeOpenChart {
  const model = toUniversalChartModel(input, options);
  const chart = buildOfficeOpenChart(model, options);
  assertOfficeOpenChart(chart);
  return chart;
}

export function convertChartToJson(input: unknown, options?: ConvertChartOptions): string {
  return toOfficeOpenChartJson(convertChart(input, options), { pretty: options?.pretty });
}
