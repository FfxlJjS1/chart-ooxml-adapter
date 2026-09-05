import { describe, expect, it } from "vitest";
import { buildOfficeOpenChart, OfficeOpenChartBuilder } from "../../src/office/OfficeOpenChartBuilder.js";
import { getPlotChartGroup } from "../../src/office/OfficeOpenTypes.js";
import { toDemoMacroChartSpace } from "../../src/office/toDemoMacroChartSpace.js";
import type { UniversalChartModel } from "../../src/model/types.js";

const SALES_BAR: UniversalChartModel = {
  type: "bar",
  categories: ["A", "B", "C"],
  series: [{ name: "Sales", values: [10, 20, 30] }],
};

describe("OfficeOpenChartBuilder", () => {
  it("converts UniversalChartModel bar data into Office Open JSON ChartML", () => {
    const chart = buildOfficeOpenChart(SALES_BAR);

    expect(chart.type).toBe("chart");
    expect(chart.chartSpace.chart.plotArea.barChart).toBeDefined();
    expect(chart.chartSpace.chart.plotArea.barChart?.ser).toEqual([
      expect.objectContaining({
        tx: "Sales",
        cat: ["A", "B", "C"],
        val: [10, 20, 30],
      }),
    ]);
  });

  it("uses ChartML element names instead of a generic chartType string", () => {
    const chart = buildOfficeOpenChart(SALES_BAR);
    const plotArea = chart.chartSpace.chart.plotArea as Record<string, unknown>;
    expect(plotArea.chartType).toBeUndefined();
    expect(plotArea.barChart).toEqual(
      expect.objectContaining({
        barDir: { val: "bar" },
        grouping: { val: "clustered" },
        ser: expect.any(Array),
        axId: [{ val: 1 }, { val: 2 }],
      }),
    );
  });

  it("maps column charts to barChart with barDir col", () => {
    const chart = buildOfficeOpenChart({
      ...SALES_BAR,
      type: "column",
    });
    expect(chart.chartSpace.chart.plotArea.barChart?.barDir).toEqual({ val: "col" });
  });

  it("emits required ChartML preamble on chartSpace and chart", () => {
    const chart = buildOfficeOpenChart(SALES_BAR);
    expect(chart.chartSpace.date1904).toEqual({ val: false });
    expect(chart.chartSpace.lang).toEqual({ val: "en-US" });
    expect(chart.chartSpace.roundedCorners).toEqual({ val: false });
    expect(chart.chartSpace.chart.autoTitleDeleted).toEqual({ val: true });
    expect(chart.chartSpace.chart.plotVisOnly).toEqual({ val: true });
    expect(chart.chartSpace.chart.dispBlanksAs).toEqual({ val: "gap" });
    expect(chart.chartSpace.chart.legend?.legendPos).toEqual({ val: "r" });
    expect(chart.chartSpace.chart.plotArea.catAx?.[0]?.axId).toEqual({ val: 1 });
    expect(chart.chartSpace.chart.plotArea.valAx?.[0]?.axId).toEqual({ val: 2 });
  });

  it("builds pie charts without axes", () => {
    const chart = buildOfficeOpenChart({
      type: "pie",
      categories: ["A", "B"],
      series: [{ name: "Share", values: [60, 40] }],
    });
    const { key, group } = getPlotChartGroup(chart.chartSpace.chart.plotArea);
    expect(key).toBe("pieChart");
    expect("axId" in group).toBe(false);
    expect(chart.chartSpace.chart.plotArea.catAx).toBeUndefined();
    expect(group.varyColors).toEqual({ val: true });
  });

  it("builds scatter series with xVal/yVal", () => {
    const chart = buildOfficeOpenChart({
      type: "scatter",
      categories: [],
      series: [{ name: "Points", xValues: [1, 2], yValues: [3, 4] }],
    });
    const ser = chart.chartSpace.chart.plotArea.scatterChart?.ser[0];
    expect(ser?.xVal).toEqual([1, 2]);
    expect(ser?.yVal).toEqual([3, 4]);
    expect(chart.chartSpace.chart.plotArea.valAx).toHaveLength(2);
  });

  it("supports the fluent builder", () => {
    const chart = OfficeOpenChartBuilder.fromUniversalModel(SALES_BAR)
      .withLang("fr-FR")
      .withLegend(false)
      .withStyle(10)
      .build();
    expect(chart.chartSpace.lang).toEqual({ val: "fr-FR" });
    expect(chart.chartSpace.chart.legend).toBeUndefined();
    expect(chart.chartSpace.style).toEqual({ val: 10 });
  });

  it("projects to DemoMacro ChartSpaceOptions without host geometry", () => {
    const chart = buildOfficeOpenChart({
      ...SALES_BAR,
      title: "Revenue",
    });
    const demo = toDemoMacroChartSpace(chart);
    expect(demo).toMatchObject({
      type: "bar",
      title: "Revenue",
      categories: ["A", "B", "C"],
      series: [{ name: "Sales", values: [10, 20, 30] }],
    });
    expect(demo).not.toHaveProperty("x");
    expect(demo).not.toHaveProperty("slides");
    expect(demo).not.toHaveProperty("worksheets");
  });

  it("does not emit XML strings", () => {
    const chart = buildOfficeOpenChart(SALES_BAR);
    const json = JSON.stringify(chart);
    expect(json).not.toMatch(/<c:(chartSpace|barChart|ser)/u);
  });
});
