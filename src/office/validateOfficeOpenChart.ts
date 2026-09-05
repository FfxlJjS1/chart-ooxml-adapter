import { OfficeOpenValidationError } from "../errors.js";
import {
  getPlotChartGroup,
  NO_AXIS_PLOT_KEYS,
  OFFICE_PLOT_CHART_KEYS,
  type OfficeChartGroup,
  type OfficeOpenChart,
  type OfficePlotArea,
  type OfficeSeries,
} from "./OfficeOpenTypes.js";

export interface OfficeOpenValidationIssue {
  path: string;
  message: string;
}

export interface OfficeOpenValidationResult {
  valid: boolean;
  errors: OfficeOpenValidationIssue[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isOnOff(value: unknown): boolean {
  return isRecord(value) && typeof value.val === "boolean";
}

function isUnsignedInt(value: unknown): boolean {
  return isRecord(value) && typeof value.val === "number" && Number.isInteger(value.val) && value.val >= 0;
}

function isStringVal(value: unknown): boolean {
  return isRecord(value) && typeof value.val === "string" && value.val.length > 0;
}

function looksLikeXml(value: unknown): boolean {
  return typeof value === "string" && /<\/?[a-zA-Z][^>]*>/u.test(value);
}

function push(errors: OfficeOpenValidationIssue[], path: string, message: string): void {
  errors.push({ path, message });
}

function requireOnOff(errors: OfficeOpenValidationIssue[], value: unknown, path: string): void {
  if (!isOnOff(value)) {
    push(errors, path, "must be { val: boolean }");
  }
}

function requireUnsignedInt(errors: OfficeOpenValidationIssue[], value: unknown, path: string): void {
  if (!isUnsignedInt(value)) {
    push(errors, path, "must be { val: integer >= 0 }");
  }
}

function collectXmlStrings(value: unknown, path: string, errors: OfficeOpenValidationIssue[]): void {
  if (typeof value === "string" && looksLikeXml(value) && value.includes("<c:")) {
    push(errors, path, "must be an object tree, not an OOXML string");
  } else if (Array.isArray(value)) {
    value.forEach((item, index) => collectXmlStrings(item, `${path}[${index}]`, errors));
  } else if (isRecord(value)) {
    for (const [key, child] of Object.entries(value)) {
      collectXmlStrings(child, `${path}.${key}`, errors);
    }
  }
}

function validateSeries(series: unknown, path: string, errors: OfficeOpenValidationIssue[]): void {
  if (!isRecord(series)) {
    push(errors, path, "series item must be an object");
    return;
  }
  if (!isUnsignedInt(series.idx)) {
    push(errors, `${path}.idx`, "required ChartML entity c:idx is missing");
  }
  if (!isUnsignedInt(series.order)) {
    push(errors, `${path}.order`, "required ChartML entity c:order is missing");
  }
  if (typeof series.tx !== "string" || series.tx.length === 0) {
    push(errors, `${path}.tx`, "required ChartML entity c:tx must be a non-empty string");
  }
  const hasCatVal = Array.isArray(series.cat) && Array.isArray(series.val);
  const hasXy = Array.isArray(series.xVal) && Array.isArray(series.yVal);
  if (!hasCatVal && !hasXy) {
    push(errors, path, "series must include cat+val or xVal+yVal");
  }
  if (hasCatVal && (series.cat as unknown[]).length !== (series.val as unknown[]).length) {
    push(errors, path, "cat and val must have the same length");
  }
  if (hasXy && (series.xVal as unknown[]).length !== (series.yVal as unknown[]).length) {
    push(errors, path, "xVal and yVal must have the same length");
  }
  if (Array.isArray(series.val) && series.val.some((item) => typeof item !== "number")) {
    push(errors, `${path}.val`, "val items must be numbers");
  }
}

function axisIds(axes: unknown): number[] {
  if (!Array.isArray(axes)) {
    return [];
  }
  return axes
    .filter(isRecord)
    .map((axis) => (isRecord(axis.axId) && typeof axis.axId.val === "number" ? axis.axId.val : undefined))
    .filter((id): id is number => id !== undefined);
}

function validateAxes(
  plotArea: OfficePlotArea,
  group: OfficeChartGroup,
  key: string,
  errors: OfficeOpenValidationIssue[],
): void {
  const plot = plotArea as Record<string, unknown>;
  if (NO_AXIS_PLOT_KEYS.has(key as (typeof NO_AXIS_PLOT_KEYS extends ReadonlySet<infer T> ? T : never))) {
    return;
  }
  if (!("axId" in group) || !Array.isArray(group.axId) || group.axId.length !== 2) {
    push(errors, `chartSpace.chart.plotArea.${key}.axId`, "cartesian charts require two c:axId values");
    return;
  }
  const referenced = group.axId.map((item) => item.val);
  const defined = [...axisIds(plot.catAx), ...axisIds(plot.valAx)];
  for (const id of referenced) {
    if (!defined.includes(id)) {
      push(
        errors,
        `chartSpace.chart.plotArea.${key}.axId`,
        `axId ${id} is not present on catAx/valAx (required ChartML axis)`,
      );
    }
  }
  if (key === "scatterChart" || key === "bubbleChart") {
    if (!Array.isArray(plot.valAx) || plot.valAx.length < 2) {
      push(errors, "chartSpace.chart.plotArea.valAx", "scatter/bubble charts require two valAx entities");
    }
  } else if (key !== "radarChart") {
    if (!Array.isArray(plot.catAx) || plot.catAx.length < 1) {
      push(errors, "chartSpace.chart.plotArea.catAx", "required ChartML entity c:catAx is missing");
    }
    if (!Array.isArray(plot.valAx) || plot.valAx.length < 1) {
      push(errors, "chartSpace.chart.plotArea.valAx", "required ChartML entity c:valAx is missing");
    }
  } else if (!Array.isArray(plot.catAx) || plot.catAx.length < 1 || !Array.isArray(plot.valAx)) {
    if (!Array.isArray(plot.catAx) || plot.catAx.length < 1) {
      push(errors, "chartSpace.chart.plotArea.catAx", "required ChartML entity c:catAx is missing");
    }
    if (!Array.isArray(plot.valAx) || plot.valAx.length < 1) {
      push(errors, "chartSpace.chart.plotArea.valAx", "required ChartML entity c:valAx is missing");
    }
  }
}

function plotKeysPresent(plotArea: Record<string, unknown>): string[] {
  return OFFICE_PLOT_CHART_KEYS.filter((key) => key in plotArea && plotArea[key] !== undefined);
}

export function validateOfficeOpenChart(input: unknown): OfficeOpenValidationResult {
  const errors: OfficeOpenValidationIssue[] = [];
  if (!isRecord(input)) {
    return { valid: false, errors: [{ path: "", message: "Office Open chart must be an object" }] };
  }

  collectXmlStrings(input, "$", errors);

  if (input.type !== "chart") {
    push(errors, "type", 'must be "chart"');
  }
  if (!isRecord(input.chartSpace)) {
    push(errors, "chartSpace", "required ChartML entity c:chartSpace is missing");
    return { valid: false, errors };
  }

  const chartSpace = input.chartSpace;
  requireOnOff(errors, chartSpace.date1904, "chartSpace.date1904");
  if (!isStringVal(chartSpace.lang)) {
    push(errors, "chartSpace.lang", "required ChartML entity c:lang is missing");
  }
  requireOnOff(errors, chartSpace.roundedCorners, "chartSpace.roundedCorners");
  if (!isRecord(chartSpace.chart)) {
    push(errors, "chartSpace.chart", "required ChartML entity c:chart is missing");
    return { valid: false, errors };
  }

  const chart = chartSpace.chart;
  requireOnOff(errors, chart.autoTitleDeleted, "chartSpace.chart.autoTitleDeleted");
  requireOnOff(errors, chart.plotVisOnly, "chartSpace.chart.plotVisOnly");
  if (!isRecord(chart.dispBlanksAs) || typeof chart.dispBlanksAs.val !== "string") {
    push(errors, "chartSpace.chart.dispBlanksAs", "required ChartML entity c:dispBlanksAs is missing");
  }
  if (!isRecord(chart.plotArea)) {
    push(errors, "chartSpace.chart.plotArea", "required ChartML entity c:plotArea is missing");
    return { valid: false, errors };
  }

  const present = plotKeysPresent(chart.plotArea);
  if (present.length !== 1) {
    push(
      errors,
      "chartSpace.chart.plotArea",
      "plotArea must contain exactly one ChartML chart-type element (barChart, lineChart, pieChart, ...)",
    );
    return { valid: false, errors };
  }

  try {
    const { key, group } = getPlotChartGroup(chart.plotArea as OfficePlotArea);
    if (!Array.isArray(group.ser) || group.ser.length === 0) {
      push(errors, `chartSpace.chart.plotArea.${key}.ser`, "required ChartML entity c:ser is missing");
    } else {
      group.ser.forEach((series: OfficeSeries, index: number) => {
        validateSeries(series, `chartSpace.chart.plotArea.${key}.ser[${index}]`, errors);
      });
    }
    if (!("varyColors" in group) || !isOnOff(group.varyColors)) {
      push(errors, `chartSpace.chart.plotArea.${key}.varyColors`, "required ChartML entity c:varyColors is missing");
    }
    validateAxes(chart.plotArea as OfficePlotArea, group, key, errors);
  } catch (error) {
    push(errors, "chartSpace.chart.plotArea", error instanceof Error ? error.message : "invalid plotArea");
  }

  return { valid: errors.length === 0, errors };
}

export function isOfficeOpenChart(input: unknown): input is OfficeOpenChart {
  return validateOfficeOpenChart(input).valid;
}

export function assertOfficeOpenChart(input: unknown): asserts input is OfficeOpenChart {
  const result = validateOfficeOpenChart(input);
  if (!result.valid) {
    throw new OfficeOpenValidationError(
      "Office Open JSON chart failed validation",
      result.errors.map((issue) => `${issue.path}: ${issue.message}`),
    );
  }
}
