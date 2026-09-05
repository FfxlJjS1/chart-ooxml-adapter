import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import Ajv2020 from "ajv/dist/2020.js";
import { describe, expect, it } from "vitest";
import { buildOfficeOpenChart } from "../../src/office/OfficeOpenChartBuilder.js";
import { parseOfficeOpenChartJson, toOfficeOpenChartJson } from "../../src/office/OfficeOpenChartJson.js";
import { getOfficeOpenChartSchema } from "../../src/office/officeOpenChartSchema.js";
import { validateOfficeOpenChart } from "../../src/office/validateOfficeOpenChart.js";
import type { UniversalChartModel } from "../../src/model/types.js";

const schemaDir = join(dirname(fileURLToPath(import.meta.url)), "../../schemas");
const schemaFromFile = JSON.parse(
  readFileSync(join(schemaDir, "office-open-chart.schema.json"), "utf8"),
) as Record<string, unknown>;

const ajv = new Ajv2020({ allErrors: true, strict: false });
const validateSchema = ajv.compile(schemaFromFile);

const MODELS: UniversalChartModel[] = [
  {
    type: "bar",
    categories: ["A", "B", "C"],
    series: [{ name: "Sales", values: [10, 20, 30] }],
  },
  {
    type: "column",
    title: "Columns",
    categories: ["Q1", "Q2"],
    series: [
      { name: "A", values: [1, 2] },
      { name: "B", values: [3, 4] },
    ],
  },
  {
    type: "line",
    categories: ["Jan", "Feb"],
    series: [{ name: "Trend", values: [5, 7] }],
  },
  {
    type: "area",
    categories: ["Jan", "Feb"],
    series: [{ name: "Fill", values: [5, 7] }],
    grouping: "stacked",
  },
  {
    type: "pie",
    categories: ["Yes", "No"],
    series: [{ name: "Votes", values: [70, 30] }],
  },
  {
    type: "doughnut",
    categories: ["Yes", "No"],
    series: [{ name: "Votes", values: [70, 30] }],
  },
  {
    type: "scatter",
    categories: [],
    series: [{ name: "Pts", xValues: [1, 2, 3], yValues: [3, 1, 4] }],
  },
  {
    type: "bubble",
    categories: [],
    series: [{ name: "Bubbles", xValues: [1, 2], yValues: [3, 4], bubbleSizes: [10, 20] }],
  },
  {
    type: "radar",
    categories: ["A", "B", "C"],
    series: [{ name: "Skill", values: [2, 4, 3] }],
  },
];

describe("Office Open JSON schema", () => {
  it("loads the canonical schema file", () => {
    expect(getOfficeOpenChartSchema().$id).toBe(schemaFromFile.$id);
    expect(schemaFromFile.title).toBe("OfficeOpenChart");
  });

  it("accepts every built chart type", () => {
    for (const model of MODELS) {
      const chart = buildOfficeOpenChart(model);
      const schemaOk = validateSchema(chart);
      const runtime = validateOfficeOpenChart(chart);
      expect(runtime.errors, `${model.type} runtime: ${JSON.stringify(runtime.errors)}`).toEqual([]);
      expect(runtime.valid, model.type).toBe(true);
      expect(schemaOk, `${model.type} schema: ${ajv.errorsText(validateSchema.errors)}`).toBe(true);
    }
  });

  it("round-trips through JSON text", () => {
    const chart = buildOfficeOpenChart(MODELS[0]!);
    const json = toOfficeOpenChartJson(chart, { pretty: true });
    const parsed = parseOfficeOpenChartJson(json);
    expect(parsed).toEqual(chart);
    expect(validateSchema(parsed)).toBe(true);
  });

  it("rejects an object that is missing plotArea", () => {
    const invalid = {
      type: "chart",
      chartSpace: {
        date1904: { val: false },
        lang: { val: "en-US" },
        roundedCorners: { val: false },
        chart: {
          autoTitleDeleted: { val: true },
          plotVisOnly: { val: true },
          dispBlanksAs: { val: "gap" },
        },
      },
    };
    expect(validateOfficeOpenChart(invalid).valid).toBe(false);
    expect(validateSchema(invalid)).toBe(false);
  });

  it("rejects series that omit cat/val", () => {
    const chart = buildOfficeOpenChart(MODELS[0]!);
    const ser = chart.chartSpace.chart.plotArea.barChart?.ser[0];
    if (ser) {
      delete ser.cat;
      delete ser.val;
    }
    const result = validateOfficeOpenChart(chart);
    expect(result.valid).toBe(false);
    expect(result.errors.some((issue) => issue.message.includes("cat+val"))).toBe(true);
  });

  it("rejects XML string payloads", () => {
    const result = validateOfficeOpenChart({
      type: "chart",
      chartSpace: "<c:chartSpace xmlns:c='...'></c:chartSpace>",
    });
    expect(result.valid).toBe(false);
    expect(result.errors.some((issue) => issue.message.includes("object tree"))).toBe(true);
  });
});
