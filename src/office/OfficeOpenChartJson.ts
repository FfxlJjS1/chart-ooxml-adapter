import { OfficeOpenValidationError } from "../errors.js";
import { parseUniversalChartModel } from "../model/UniversalChartModel.js";
import type { UniversalChartModel } from "../model/types.js";
import { buildOfficeOpenChart } from "./OfficeOpenChartBuilder.js";
import type { OfficeOpenChart, OfficeOpenChartJson } from "./OfficeOpenTypes.js";
import { assertOfficeOpenChart, isOfficeOpenChart } from "./validateOfficeOpenChart.js";

export interface OfficeOpenJsonOptions {
  pretty?: boolean;
}

function isUniversalLike(value: unknown): value is UniversalChartModel {
  return (
    typeof value === "object" &&
    value !== null &&
    "type" in value &&
    (value as { type: unknown }).type !== "chart" &&
    "series" in value &&
    "categories" in value
  );
}

/**
 * Convert a UniversalChartModel or an OfficeOpenChart object into JSON text.
 * The payload is an object tree, never an XML document.
 */
export function toOfficeOpenChartJson(
  input: UniversalChartModel | OfficeOpenChart,
  options: OfficeOpenJsonOptions = {},
): string {
  const chart = isUniversalLike(input) ? buildOfficeOpenChart(parseUniversalChartModel(input)) : input;
  assertOfficeOpenChart(chart);
  return JSON.stringify(chart, null, options.pretty ? 2 : undefined);
}

export function parseOfficeOpenChartJson(json: string): OfficeOpenChartJson {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch (error) {
    throw new OfficeOpenValidationError("Office Open JSON is not valid JSON", [
      error instanceof Error ? error.message : String(error),
    ]);
  }
  assertOfficeOpenChart(parsed);
  return parsed;
}

/** Structured clone of an Office Open chart as a plain JSON object. */
export function toOfficeOpenChartObject(
  input: UniversalChartModel | OfficeOpenChart,
): OfficeOpenChartJson {
  return parseOfficeOpenChartJson(toOfficeOpenChartJson(input));
}

export { isOfficeOpenChart };
