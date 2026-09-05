import { InvalidChartInputError, UnsupportedChartSourceError } from "../errors.js";
import { isUniversalChartModel, parseUniversalChartModel } from "../model/UniversalChartModel.js";
import type { ChartSource, UniversalChartModel } from "../model/types.js";
import { amchartsAdapter } from "./amcharts.js";
import { chartJsAdapter } from "./chartjs.js";
import { echartsAdapter } from "./echarts.js";
import { highchartsAdapter } from "./highcharts.js";
import { isRecord, type ChartSourceAdapter } from "./types.js";

const ADAPTERS: readonly ChartSourceAdapter[] = [
  chartJsAdapter,
  highchartsAdapter,
  amchartsAdapter,
  echartsAdapter,
];

const ADAPTER_BY_SOURCE: Record<Exclude<ChartSource, "universal">, ChartSourceAdapter> = {
  chartjs: chartJsAdapter,
  highcharts: highchartsAdapter,
  amcharts: amchartsAdapter,
  echarts: echartsAdapter,
};

export function detectChartSource(input: unknown): ChartSource {
  if (isUniversalChartModel(input)) {
    return "universal";
  }
  for (const adapter of ADAPTERS) {
    if (adapter.matches(input)) {
      return adapter.source;
    }
  }
  throw new UnsupportedChartSourceError(
    "Could not detect chart source. Pass { source: \"echarts\" | \"highcharts\" | \"chartjs\" | \"amcharts\" } or a UniversalChartModel.",
  );
}

export function toUniversalChartModel(
  input: unknown,
  options?: { source?: ChartSource },
): UniversalChartModel {
  if (input === undefined || input === null) {
    throw new InvalidChartInputError("Chart input is required");
  }
  if (!isRecord(input) && !isUniversalChartModel(input)) {
    throw new InvalidChartInputError("Chart input must be an object");
  }

  const source = options?.source ?? detectChartSource(input);
  if (source === "universal") {
    return parseUniversalChartModel(input);
  }
  const adapter = ADAPTER_BY_SOURCE[source];
  if (!adapter) {
    throw new UnsupportedChartSourceError(`Unknown chart source: ${String(source)}`);
  }
  return parseUniversalChartModel(adapter.toUniversalModel(input));
}

export { chartJsAdapter, echartsAdapter, highchartsAdapter, amchartsAdapter };
