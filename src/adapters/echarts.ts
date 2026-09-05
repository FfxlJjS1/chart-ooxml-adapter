import { InvalidChartInputError, UnsupportedChartTypeError } from "../errors.js";
import type { UniversalChartModel, UniversalChartType } from "../model/types.js";
import { asFiniteNumber, asString, isRecord, toNumberArray, type ChartSourceAdapter } from "./types.js";

const ECHARTS_TYPE_MAP: Record<string, UniversalChartType> = {
  bar: "column",
  line: "line",
  pie: "pie",
  doughnut: "doughnut",
  scatter: "scatter",
  radar: "radar",
  effectScatter: "scatter",
};

function firstAxis(axis: unknown): Record<string, unknown> | undefined {
  if (Array.isArray(axis)) {
    return isRecord(axis[0]) ? axis[0] : undefined;
  }
  return isRecord(axis) ? axis : undefined;
}

function axisData(axis: Record<string, unknown> | undefined): string[] {
  if (!axis || !Array.isArray(axis.data)) {
    return [];
  }
  return axis.data.map((item) => asString(item) ?? "");
}

function mapSeriesType(type: string | undefined, fallback: UniversalChartType): UniversalChartType {
  if (!type) {
    return fallback;
  }
  const mapped = ECHARTS_TYPE_MAP[type];
  if (!mapped) {
    throw new UnsupportedChartTypeError(`Unsupported ECharts series type: ${type}`);
  }
  return mapped;
}

function pieCategoriesAndValues(seriesItem: Record<string, unknown>): {
  categories: string[];
  values: number[];
} {
  const data = Array.isArray(seriesItem.data) ? seriesItem.data : [];
  const categories: string[] = [];
  const values: number[] = [];
  for (const point of data) {
    if (isRecord(point)) {
      categories.push(asString(point.name) ?? "");
      values.push(asFiniteNumber(point.value) ?? 0);
    } else {
      categories.push("");
      values.push(asFiniteNumber(point) ?? 0);
    }
  }
  return { categories, values };
}

export const echartsAdapter: ChartSourceAdapter = {
  source: "echarts",
  matches(input: unknown): boolean {
    if (!isRecord(input) || !Array.isArray(input.series)) {
      return false;
    }
    if (isRecord(input.data) && Array.isArray(input.data.datasets)) {
      return false;
    }
    if (isRecord(input.chart) && typeof input.chart.type === "string") {
      return false;
    }
    return true;
  },
  toUniversalModel(input: unknown): UniversalChartModel {
    if (!this.matches(input) || !isRecord(input)) {
      throw new InvalidChartInputError("ECharts option must include a series array");
    }
    const seriesItems = input.series as unknown[];
    if (seriesItems.length === 0) {
      throw new InvalidChartInputError("ECharts series must be a non-empty array");
    }

    const xAxis = firstAxis(input.xAxis);
    const yAxis = firstAxis(input.yAxis);
    const xData = axisData(xAxis);
    const yData = axisData(yAxis);
    const horizontal = (xAxis?.type === "value" || xAxis?.type === "log") && yData.length > 0;

    const firstSeries = isRecord(seriesItems[0]) ? seriesItems[0] : {};
    const firstType = asString(firstSeries.type);
    let type = mapSeriesType(firstType, horizontal ? "bar" : "column");
    if (type === "column" && horizontal) {
      type = "bar";
    }
    if (firstType === "pie" && (firstSeries.radius === "50%" || Array.isArray(firstSeries.radius))) {
      const radius = firstSeries.radius;
      if (Array.isArray(radius) && radius.length >= 2) {
        type = "doughnut";
      }
    }

    const title = isRecord(input.title) ? asString(input.title.text) : asString(input.title);
    const legendInput = input.legend;
    const legendShow = legendInput === undefined ? true : legendInput !== false;
    const legendPosition = isRecord(legendInput)
      ? asString(legendInput.orient) === "vertical"
        ? "right"
        : asString(legendInput.top) === "bottom"
          ? "bottom"
          : undefined
      : undefined;

    let categories = horizontal ? yData : xData;
    const grouping = seriesItems.some(
      (item) => isRecord(item) && typeof item.stack === "string" && item.stack.length > 0,
    )
      ? "stacked"
      : undefined;

    const series = seriesItems.map((item, index) => {
      if (!isRecord(item)) {
        throw new InvalidChartInputError(`ECharts series[${index}] must be an object`);
      }
      const name = asString(item.name) ?? `Series ${index + 1}`;
      if (type === "pie" || type === "doughnut") {
        const pie = pieCategoriesAndValues(item);
        if (index === 0) {
          categories = pie.categories;
        }
        return { name, values: pie.values };
      }
      if (type === "scatter") {
        const points = Array.isArray(item.data) ? item.data : [];
        return {
          name,
          xValues: points.map((point) =>
            Array.isArray(point) ? (asFiniteNumber(point[0]) ?? 0) : 0,
          ),
          yValues: points.map((point) =>
            Array.isArray(point) ? (asFiniteNumber(point[1]) ?? 0) : 0,
          ),
        };
      }
      return {
        name,
        values: toNumberArray(Array.isArray(item.data) ? item.data : []),
      };
    });

    return {
      type,
      title,
      categories,
      series,
      grouping,
      legend: { show: legendShow !== false, position: legendPosition },
      source: { library: "echarts" },
    };
  },
};
