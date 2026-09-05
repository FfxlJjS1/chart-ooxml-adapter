import type { ChartSource, UniversalChartModel } from "../model/types.js";

export interface ConvertOptions {
  /** Force a source adapter instead of auto-detection. */
  source?: ChartSource;
  /** Chart language written to `c:lang` (default `en-US`). */
  lang?: string;
  /** Override legend visibility. */
  showLegend?: boolean;
  /** OOXML chart style index (`c:style`, default 2). */
  style?: number;
}

export interface ChartSourceAdapter {
  readonly source: ChartSource;
  matches(input: unknown): boolean;
  toUniversalModel(input: unknown): UniversalChartModel;
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function asString(value: unknown): string | undefined {
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(value);
  }
  return undefined;
}

export function asFiniteNumber(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) {
    return Number(value);
  }
  return undefined;
}

export function toNumberArray(values: readonly unknown[]): number[] {
  return values.map((value) => {
    const number = asFiniteNumber(value);
    return number === undefined ? 0 : number;
  });
}

export function firstDefined<T>(...values: Array<T | undefined>): T | undefined {
  return values.find((value) => value !== undefined);
}
