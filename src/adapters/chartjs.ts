import { InvalidChartInputError } from "../errors.js";
import type { UniversalChartModel, UniversalChartType } from "../model/types.js";
import { asFiniteNumber, asString, isRecord, toNumberArray, type ChartSourceAdapter } from "./types.js";

const CHARTJS_TYPE_MAP: Record<string, UniversalChartType> = {
  bar: "column",
  line: "line",
  pie: "pie",
  doughnut: "doughnut",
  scatter: "scatter",
  bubble: "bubble",
  radar: "radar",
  polarArea: "pie",
};

function chartJsType(input: Record<string, unknown>): UniversalChartType {
  const type = asString(input.type)?.toLowerCase() ?? "bar";
  const mapped = CHARTJS_TYPE_MAP[type];
  if (!mapped) {
    throw new InvalidChartInputError(`Unsupported Chart.js type: ${type}`);
  }
  const options = isRecord(input.options) ? input.options : {};
  if (mapped === "column" && options.indexAxis === "y") {
    return "bar";
  }
  return mapped;
}

function groupingFromOptions(input: Record<string, unknown>): UniversalChartModel["grouping"] {
  const options = isRecord(input.options) ? input.options : {};
  const scales = isRecord(options.scales) ? options.scales : {};
  const stacked =
    (isRecord(scales.x) && scales.x.stacked === true) ||
    (isRecord(scales.y) && scales.y.stacked === true);
  if (!stacked) {
    return undefined;
  }
  return "stacked";
}

function legendFromOptions(input: Record<string, unknown>): UniversalChartModel["legend"] {
  const options = isRecord(input.options) ? input.options : {};
  const plugins = isRecord(options.plugins) ? options.plugins : {};
  const legend = isRecord(plugins.legend) ? plugins.legend : {};
  const display = legend.display;
  const position = asString(legend.position);
  const mappedPosition =
    position === "left" || position === "right" || position === "top" || position === "bottom"
      ? position
      : undefined;
  return {
    show: display === undefined ? true : display !== false,
    position: mappedPosition,
  };
}

function titleFromOptions(input: Record<string, unknown>): string | undefined {
  const options = isRecord(input.options) ? input.options : {};
  const plugins = isRecord(options.plugins) ? options.plugins : {};
  const title = isRecord(plugins.title) ? plugins.title : {};
  return asString(title.text);
}

export const chartJsAdapter: ChartSourceAdapter = {
  source: "chartjs",
  matches(input: unknown): boolean {
    return isRecord(input) && isRecord(input.data) && Array.isArray(input.data.datasets);
  },
  toUniversalModel(input: unknown): UniversalChartModel {
    if (!this.matches(input) || !isRecord(input)) {
      throw new InvalidChartInputError("Chart.js config must include data.datasets");
    }
    const data = input.data as Record<string, unknown>;
    const datasets = data.datasets as unknown[];
    const type = chartJsType(input);
    const labels = Array.isArray(data.labels) ? data.labels.map((item) => asString(item) ?? "") : [];

    const series = datasets.map((dataset, index) => {
      if (!isRecord(dataset)) {
        throw new InvalidChartInputError(`Chart.js datasets[${index}] must be an object`);
      }
      const name = asString(dataset.label) ?? `Series ${index + 1}`;
      if (type === "scatter" || type === "bubble") {
        const points = Array.isArray(dataset.data) ? dataset.data : [];
        return {
          name,
          xValues: points.map((point) =>
            isRecord(point) ? (asFiniteNumber(point.x) ?? 0) : 0,
          ),
          yValues: points.map((point) =>
            isRecord(point) ? (asFiniteNumber(point.y) ?? 0) : 0,
          ),
          bubbleSizes:
            type === "bubble"
              ? points.map((point) => (isRecord(point) ? (asFiniteNumber(point.r) ?? 1) : 1))
              : undefined,
          color: asString(dataset.backgroundColor) ?? asString(dataset.borderColor),
        };
      }
      return {
        name,
        values: toNumberArray(Array.isArray(dataset.data) ? dataset.data : []),
        color: asString(dataset.backgroundColor) ?? asString(dataset.borderColor),
      };
    });

    return {
      type,
      title: titleFromOptions(input),
      categories: labels,
      series,
      grouping: groupingFromOptions(input),
      legend: legendFromOptions(input),
      source: { library: "chartjs" },
    };
  },
};
