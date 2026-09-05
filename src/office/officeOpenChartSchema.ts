import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const schemaUrl = fileURLToPath(import.meta.url);

export const OFFICE_OPEN_CHART_SCHEMA_ID =
  "https://github.com/FfxlJjS1/chart-ooxml-adapter/schemas/office-open-chart.schema.json";

/**
 * Canonical JSON Schema for the Office Open chart object.
 * Loaded from `schemas/office-open-chart.schema.json`.
 */
export function loadOfficeOpenChartSchema(): Record<string, unknown> {
  const schemaPath = join(dirname(schemaUrl), "../../schemas/office-open-chart.schema.json");
  return JSON.parse(readFileSync(schemaPath, "utf8")) as Record<string, unknown>;
}

let cached: Record<string, unknown> | undefined;

export function getOfficeOpenChartSchema(): Record<string, unknown> {
  cached ??= loadOfficeOpenChartSchema();
  return cached;
}
