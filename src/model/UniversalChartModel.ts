import { InvalidUniversalModelError } from "../errors.js";
import {
  CATEGORY_CHART_TYPES,
  SCATTER_LIKE_TYPES,
  UNIVERSAL_CHART_TYPES,
  UNIVERSAL_GROUPINGS,
  UNIVERSAL_LEGEND_POSITIONS,
  type UniversalChartModel,
  type UniversalChartType,
  type UniversalLegendPosition,
  type UniversalSeries,
} from "./types.js";

const UNIVERSAL_TYPE_SET = new Set<string>(UNIVERSAL_CHART_TYPES);
const GROUPING_SET = new Set<string>(UNIVERSAL_GROUPINGS);
const LEGEND_POSITION_SET = new Set<string>(UNIVERSAL_LEGEND_POSITIONS);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function asNumberArray(value: unknown, label: string, errors: string[]): number[] | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (!Array.isArray(value)) {
    errors.push(`${label} must be an array of numbers`);
    return undefined;
  }
  const numbers: number[] = [];
  for (let i = 0; i < value.length; i += 1) {
    const item = value[i];
    if (item === null || item === undefined) {
      numbers.push(0);
      continue;
    }
    if (!isFiniteNumber(item)) {
      errors.push(`${label}[${i}] must be a finite number`);
      continue;
    }
    numbers.push(item);
  }
  return numbers;
}

function parseSeries(value: unknown, index: number, errors: string[]): UniversalSeries | undefined {
  if (!isRecord(value)) {
    errors.push(`series[${index}] must be an object`);
    return undefined;
  }
  const name = value.name;
  if (typeof name !== "string" || name.trim().length === 0) {
    errors.push(`series[${index}].name must be a non-empty string`);
    return undefined;
  }
  const series: UniversalSeries = { name: name.trim() };
  series.values = asNumberArray(value.values, `series[${index}].values`, errors);
  series.xValues = asNumberArray(value.xValues, `series[${index}].xValues`, errors);
  series.yValues = asNumberArray(value.yValues, `series[${index}].yValues`, errors);
  series.bubbleSizes = asNumberArray(value.bubbleSizes, `series[${index}].bubbleSizes`, errors);
  if (value.color !== undefined) {
    if (typeof value.color !== "string") {
      errors.push(`series[${index}].color must be a string`);
    } else {
      series.color = value.color;
    }
  }
  return series;
}

export function isUniversalChartType(value: unknown): value is UniversalChartType {
  return typeof value === "string" && UNIVERSAL_TYPE_SET.has(value);
}

export function isUniversalChartModel(value: unknown): value is UniversalChartModel {
  if (!isRecord(value)) {
    return false;
  }
  if (!isUniversalChartType(value.type)) {
    return false;
  }
  if (!Array.isArray(value.categories) || !Array.isArray(value.series)) {
    return false;
  }
  if ("data" in value && isRecord(value.data) && Array.isArray(value.data.datasets)) {
    return false;
  }
  return true;
}

export function parseUniversalChartModel(input: unknown): UniversalChartModel {
  if (!isRecord(input)) {
    throw new InvalidUniversalModelError("UniversalChartModel must be an object");
  }

  const errors: string[] = [];

  if (!isUniversalChartType(input.type)) {
    errors.push(`type must be one of: ${UNIVERSAL_CHART_TYPES.join(", ")}`);
  }

  const categories = Array.isArray(input.categories)
    ? input.categories.map((item, index) => {
        if (item === null || item === undefined) {
          return "";
        }
        if (typeof item === "string" || typeof item === "number") {
          return String(item);
        }
        errors.push(`categories[${index}] must be a string or number`);
        return "";
      })
    : [];

  if (!Array.isArray(input.categories)) {
    errors.push("categories must be an array");
  }

  if (!Array.isArray(input.series) || input.series.length === 0) {
    errors.push("series must be a non-empty array");
  }

  const series = Array.isArray(input.series)
    ? input.series
        .map((item, index) => parseSeries(item, index, errors))
        .filter((item): item is UniversalSeries => item !== undefined)
    : [];

  const type = isUniversalChartType(input.type) ? input.type : "bar";

  if (CATEGORY_CHART_TYPES.has(type) && categories.length === 0) {
    errors.push(`categories must be a non-empty array for ${type} charts`);
  }

  for (const [index, item] of series.entries()) {
    if (SCATTER_LIKE_TYPES.has(type)) {
      const xValues = item.xValues ?? [];
      const yValues = item.yValues ?? item.values ?? [];
      if (xValues.length === 0 || yValues.length === 0) {
        errors.push(`series[${index}] of type ${type} requires xValues and yValues`);
      } else if (xValues.length !== yValues.length) {
        errors.push(`series[${index}] xValues and yValues must have the same length`);
      }
      if (type === "bubble") {
        const sizes = item.bubbleSizes ?? [];
        if (sizes.length === 0) {
          errors.push(`series[${index}] of type bubble requires bubbleSizes`);
        } else if (sizes.length !== yValues.length) {
          errors.push(`series[${index}] bubbleSizes must match yValues length`);
        }
      }
      continue;
    }

    const values = item.values ?? [];
    if (values.length === 0) {
      errors.push(`series[${index}].values must be a non-empty array`);
    } else if (CATEGORY_CHART_TYPES.has(type) && values.length !== categories.length) {
      errors.push(
        `series[${index}].values length (${values.length}) must match categories length (${categories.length})`,
      );
    }
  }

  let grouping: UniversalChartModel["grouping"];
  if (input.grouping !== undefined) {
    if (typeof input.grouping !== "string" || !GROUPING_SET.has(input.grouping)) {
      errors.push(`grouping must be one of: ${UNIVERSAL_GROUPINGS.join(", ")}`);
    } else {
      grouping = input.grouping as UniversalChartModel["grouping"];
    }
  }

  let legend: UniversalChartModel["legend"];
  if (input.legend !== undefined) {
    if (!isRecord(input.legend)) {
      errors.push("legend must be an object");
    } else {
      legend = {};
      if (input.legend.show !== undefined) {
        if (typeof input.legend.show !== "boolean") {
          errors.push("legend.show must be a boolean");
        } else {
          legend.show = input.legend.show;
        }
      }
      if (input.legend.position !== undefined) {
        if (
          typeof input.legend.position !== "string" ||
          !LEGEND_POSITION_SET.has(input.legend.position)
        ) {
          errors.push(`legend.position must be one of: ${UNIVERSAL_LEGEND_POSITIONS.join(", ")}`);
        } else {
          legend.position = input.legend.position as UniversalLegendPosition;
        }
      }
    }
  }

  const parseAxis = (value: unknown, label: string): UniversalChartModel["xAxis"] => {
    if (value === undefined) {
      return undefined;
    }
    if (!isRecord(value)) {
      errors.push(`${label} must be an object`);
      return undefined;
    }
    const axis: NonNullable<UniversalChartModel["xAxis"]> = {};
    if (value.title !== undefined) {
      if (typeof value.title !== "string") {
        errors.push(`${label}.title must be a string`);
      } else {
        axis.title = value.title;
      }
    }
    if (value.min !== undefined) {
      if (!isFiniteNumber(value.min)) {
        errors.push(`${label}.min must be a finite number`);
      } else {
        axis.min = value.min;
      }
    }
    if (value.max !== undefined) {
      if (!isFiniteNumber(value.max)) {
        errors.push(`${label}.max must be a finite number`);
      } else {
        axis.max = value.max;
      }
    }
    return axis;
  };

  const model: UniversalChartModel = {
    type,
    categories,
    series,
    grouping,
    legend,
    xAxis: parseAxis(input.xAxis, "xAxis"),
    yAxis: parseAxis(input.yAxis, "yAxis"),
  };

  if (input.title !== undefined) {
    if (typeof input.title !== "string") {
      errors.push("title must be a string");
    } else if (input.title.trim().length > 0) {
      model.title = input.title.trim();
    }
  }

  if (input.threeD !== undefined) {
    if (typeof input.threeD !== "boolean") {
      errors.push("threeD must be a boolean");
    } else {
      model.threeD = input.threeD;
    }
  }

  if (isRecord(input.source) && typeof input.source.library === "string") {
    model.source = { library: input.source.library as NonNullable<UniversalChartModel["source"]>["library"] };
  }

  if (errors.length > 0) {
    throw new InvalidUniversalModelError("UniversalChartModel is invalid", errors);
  }

  return model;
}
