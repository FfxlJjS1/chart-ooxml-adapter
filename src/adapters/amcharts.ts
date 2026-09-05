import { InvalidChartInputError, UnsupportedChartTypeError } from "../errors.js";
import type { UniversalChartModel, UniversalChartType } from "../model/types.js";
import { asFiniteNumber, asString, isRecord, type ChartSourceAdapter } from "./types.js";

const SERIES_TYPE_MAP: Record<string, UniversalChartType> = {
  ColumnSeries: "column",
  BarSeries: "bar",
  LineSeries: "line",
  SmoothedXLineSeries: "line",
  AreaSeries: "area",
  PieSeries: "pie",
  RadarSeries: "radar",
  XYSeries: "scatter",
  column: "column",
  bar: "bar",
  line: "line",
  area: "area",
  pie: "pie",
  scatter: "scatter",
};

function mapType(type: string | undefined): UniversalChartType | undefined {
  if (!type) {
    return undefined;
  }
  return SERIES_TYPE_MAP[type];
}

function chartTypeFromRoot(input: Record<string, unknown>): UniversalChartType | undefined {
  const rootType = asString(input.type);
  if (rootType === "PieChart") {
    return "pie";
  }
  if (rootType === "RadarChart") {
    return "radar";
  }
  if (rootType === "XYChart" || rootType === "SerialChart") {
    return undefined;
  }
  return undefined;
}

function rows(input: Record<string, unknown>): Record<string, unknown>[] {
  const data = input.dataProvider ?? input.data;
  if (!Array.isArray(data)) {
    return [];
  }
  return data.filter(isRecord);
}

function fieldValues(records: Record<string, unknown>[], field: string): number[] {
  return records.map((row) => asFiniteNumber(row[field]) ?? 0);
}

function fieldLabels(records: Record<string, unknown>[], field: string): string[] {
  return records.map((row) => asString(row[field]) ?? "");
}

export const amchartsAdapter: ChartSourceAdapter = {
  source: "amcharts",
  matches(input: unknown): boolean {
    if (!isRecord(input)) {
      return false;
    }
    if (isRecord(input.data) && Array.isArray(input.data.datasets)) {
      return false;
    }
    if (Array.isArray(input.dataProvider)) {
      return true;
    }
    const type = asString(input.type);
    if (type && /Chart$/u.test(type) && (Array.isArray(input.data) || Array.isArray(input.series))) {
      return true;
    }
    return Array.isArray(input.graphs) && Array.isArray(input.dataProvider);
  },
  toUniversalModel(input: unknown): UniversalChartModel {
    if (!this.matches(input) || !isRecord(input)) {
      throw new InvalidChartInputError("amCharts config must include data and series/graphs");
    }
    const records = rows(input);
    const graphs = Array.isArray(input.graphs) ? input.graphs.filter(isRecord) : [];
    const seriesNodes = Array.isArray(input.series) ? input.series.filter(isRecord) : graphs;
    if (seriesNodes.length === 0 && records.length === 0) {
      throw new InvalidChartInputError("amCharts config is missing series data");
    }

    const firstSeries = seriesNodes[0];
    const mappedFromSeries = firstSeries ? mapType(asString(firstSeries.type)) : undefined;
    const type = mappedFromSeries ?? chartTypeFromRoot(input) ?? "column";
    if (!mapType(asString(firstSeries?.type)) && firstSeries?.type && !chartTypeFromRoot(input)) {
      const raw = asString(firstSeries.type);
      if (raw && !SERIES_TYPE_MAP[raw]) {
        throw new UnsupportedChartTypeError(`Unsupported amCharts series type: ${raw}`);
      }
    }

    const categoryField =
      asString(input.categoryField) ??
      asString(firstSeries?.categoryXField) ??
      asString(firstSeries?.categoryField) ??
      "category";

    let categories = fieldLabels(records, categoryField);
    const title =
      asString(input.title) ??
      (isRecord(input.titles) ? undefined : undefined) ??
      (Array.isArray(input.titles) && isRecord(input.titles[0])
        ? asString(input.titles[0].text)
        : undefined);

    const series = (seriesNodes.length > 0 ? seriesNodes : [{ name: "Series 1" }]).map(
      (item, index) => {
        const name = asString(item.name) ?? asString(item.title) ?? `Series ${index + 1}`;
        if (type === "pie") {
          const valueField = asString(item.valueField) ?? asString(input.valueField) ?? "value";
          const categoryPieField =
            asString(item.categoryField) ?? asString(input.categoryField) ?? categoryField;
          categories = fieldLabels(records, categoryPieField);
          return { name, values: fieldValues(records, valueField) };
        }
        if (type === "scatter") {
          const xField = asString(item.valueXField) ?? "x";
          const yField = asString(item.valueYField) ?? "y";
          return {
            name,
            xValues: fieldValues(records, xField),
            yValues: fieldValues(records, yField),
          };
        }
        const valueField =
          asString(item.valueYField) ??
          asString(item.valueField) ??
          asString(item.valueY) ??
          "value";
        return { name, values: fieldValues(records, valueField) };
      },
    );

    return {
      type,
      title,
      categories,
      series,
      legend: { show: input.legend !== false },
      source: { library: "amcharts" },
    };
  },
};
