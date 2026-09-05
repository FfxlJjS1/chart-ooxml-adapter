import { InvalidChartInputError, UnsupportedChartTypeError } from "../errors.js";
import type { UniversalChartModel, UniversalChartType } from "../model/types.js";
import { asFiniteNumber, asString, isRecord, toNumberArray, type ChartSourceAdapter } from "./types.js";

const HIGHCHARTS_TYPE_MAP: Record<string, UniversalChartType> = {
  bar: "bar",
  column: "column",
  line: "line",
  area: "area",
  pie: "pie",
  scatter: "scatter",
  bubble: "bubble",
  spline: "line",
  areaspline: "area",
  columnrange: "column",
};

function firstAxis(axis: unknown): Record<string, unknown> | undefined {
  if (Array.isArray(axis)) {
    return isRecord(axis[0]) ? axis[0] : undefined;
  }
  return isRecord(axis) ? axis : undefined;
}

function categoriesFromAxis(axis: unknown): string[] {
  const first = firstAxis(axis);
  if (!first || !Array.isArray(first.categories)) {
    return [];
  }
  return first.categories.map((item) => asString(item) ?? "");
}

function mapType(type: string | undefined): UniversalChartType {
  const key = type?.toLowerCase() ?? "column";
  const mapped = HIGHCHARTS_TYPE_MAP[key];
  if (!mapped) {
    throw new UnsupportedChartTypeError(`Unsupported Highcharts type: ${key}`);
  }
  return mapped;
}

function groupingFromPlotOptions(input: Record<string, unknown>): UniversalChartModel["grouping"] {
  const plotOptions = isRecord(input.plotOptions) ? input.plotOptions : {};
  const series = isRecord(plotOptions.series) ? plotOptions.series : {};
  const stacking = asString(series.stacking) ?? asString(
    isRecord(plotOptions.column)
      ? plotOptions.column.stacking
      : isRecord(plotOptions.bar)
        ? plotOptions.bar.stacking
        : undefined,
  );
  if (stacking === "percent") {
    return "percentStacked";
  }
  if (stacking === "normal") {
    return "stacked";
  }
  return undefined;
}

function pieFromSeries(seriesItem: Record<string, unknown>): {
  categories: string[];
  values: number[];
} {
  const data = Array.isArray(seriesItem.data) ? seriesItem.data : [];
  const categories: string[] = [];
  const values: number[] = [];
  for (const point of data) {
    if (Array.isArray(point)) {
      categories.push(asString(point[0]) ?? "");
      values.push(asFiniteNumber(point[1]) ?? 0);
    } else if (isRecord(point)) {
      categories.push(asString(point.name) ?? "");
      values.push(asFiniteNumber(point.y) ?? asFiniteNumber(point.value) ?? 0);
    } else {
      categories.push("");
      values.push(asFiniteNumber(point) ?? 0);
    }
  }
  return { categories, values };
}

export const highchartsAdapter: ChartSourceAdapter = {
  source: "highcharts",
  matches(input: unknown): boolean {
    if (!isRecord(input) || !Array.isArray(input.series)) {
      return false;
    }
    if (isRecord(input.data) && Array.isArray(input.data.datasets)) {
      return false;
    }
    if (isRecord(input.chart) && typeof input.chart.type === "string") {
      return true;
    }
    const xAxis = firstAxis(input.xAxis);
    return Boolean(xAxis && Array.isArray(xAxis.categories));
  },
  toUniversalModel(input: unknown): UniversalChartModel {
    if (!this.matches(input) || !isRecord(input)) {
      throw new InvalidChartInputError("Highcharts config must include series");
    }
    const chart = isRecord(input.chart) ? input.chart : {};
    const seriesItems = input.series as unknown[];
    if (seriesItems.length === 0) {
      throw new InvalidChartInputError("Highcharts series must be a non-empty array");
    }
    const firstSeries = isRecord(seriesItems[0]) ? seriesItems[0] : {};
    const type = mapType(asString(firstSeries.type) ?? asString(chart.type));
    const title = isRecord(input.title) ? asString(input.title.text) : undefined;
    const legend = isRecord(input.legend) ? input.legend : {};
    let categories = categoriesFromAxis(input.xAxis);

    const series = seriesItems.map((item, index) => {
      if (!isRecord(item)) {
        throw new InvalidChartInputError(`Highcharts series[${index}] must be an object`);
      }
      const name = asString(item.name) ?? `Series ${index + 1}`;
      if (type === "pie" || type === "doughnut") {
        const pie = pieFromSeries(item);
        if (index === 0) {
          categories = pie.categories;
        }
        return { name, values: pie.values };
      }
      if (type === "scatter" || type === "bubble") {
        const points = Array.isArray(item.data) ? item.data : [];
        return {
          name,
          xValues: points.map((point) =>
            Array.isArray(point)
              ? (asFiniteNumber(point[0]) ?? 0)
              : isRecord(point)
                ? (asFiniteNumber(point.x) ?? 0)
                : 0,
          ),
          yValues: points.map((point) =>
            Array.isArray(point)
              ? (asFiniteNumber(point[1]) ?? 0)
              : isRecord(point)
                ? (asFiniteNumber(point.y) ?? 0)
                : 0,
          ),
          bubbleSizes:
            type === "bubble"
              ? points.map((point) =>
                  Array.isArray(point)
                    ? (asFiniteNumber(point[2]) ?? 1)
                    : isRecord(point)
                      ? (asFiniteNumber(point.z) ?? 1)
                      : 1,
                )
              : undefined,
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
      grouping: groupingFromPlotOptions(input),
      legend: {
        show: legend.enabled !== false,
        position:
          asString(legend.align) === "left"
            ? "left"
            : asString(legend.verticalAlign) === "top"
              ? "top"
              : asString(legend.verticalAlign) === "bottom"
                ? "bottom"
                : "right",
      },
      source: { library: "highcharts" },
    };
  },
};
