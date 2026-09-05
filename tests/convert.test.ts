import { describe, expect, it } from "vitest";
import { convertChart, convertChartToJson } from "../src/convert.js";
import { detectChartSource } from "../src/adapters/index.js";
import { validateOfficeOpenChart } from "../src/office/validateOfficeOpenChart.js";

describe("convertChart pipeline", () => {
  it("converts a UniversalChartModel into Office Open JSON", () => {
    const chart = convertChart({
      type: "bar",
      categories: ["A", "B", "C"],
      series: [{ name: "Sales", values: [10, 20, 30] }],
    });
    expect(chart.chartSpace.chart.plotArea.barChart?.ser[0]).toMatchObject({
      tx: "Sales",
      cat: ["A", "B", "C"],
      val: [10, 20, 30],
    });
    expect(validateOfficeOpenChart(chart).valid).toBe(true);
  });

  it("auto-detects Chart.js", () => {
    const input = {
      type: "bar",
      data: {
        labels: ["A", "B"],
        datasets: [{ label: "Sales", data: [1, 2] }],
      },
    };
    expect(detectChartSource(input)).toBe("chartjs");
    const chart = convertChart(input);
    expect(chart.chartSpace.chart.plotArea.barChart?.barDir).toEqual({ val: "col" });
    expect(chart.chartSpace.chart.plotArea.barChart?.ser[0]?.tx).toBe("Sales");
  });

  it("maps Chart.js indexAxis y to a horizontal bar chart", () => {
    const chart = convertChart({
      type: "bar",
      data: {
        labels: ["A", "B"],
        datasets: [{ label: "Sales", data: [1, 2] }],
      },
      options: { indexAxis: "y" },
    });
    expect(chart.chartSpace.chart.plotArea.barChart?.barDir).toEqual({ val: "bar" });
  });

  it("auto-detects ECharts", () => {
    const input = {
      title: { text: "Sales" },
      xAxis: { type: "category", data: ["A", "B", "C"] },
      yAxis: { type: "value" },
      series: [{ type: "bar", name: "Sales", data: [10, 20, 30] }],
    };
    expect(detectChartSource(input)).toBe("echarts");
    const chart = convertChart(input);
    expect(chart.chartSpace.chart.title?.tx).toBe("Sales");
    expect(chart.chartSpace.chart.plotArea.barChart?.ser[0]?.val).toEqual([10, 20, 30]);
  });

  it("auto-detects Highcharts", () => {
    const input = {
      chart: { type: "column" },
      title: { text: "Quarter" },
      xAxis: { categories: ["Q1", "Q2"] },
      series: [{ name: "A", data: [1, 2] }],
    };
    expect(detectChartSource(input)).toBe("highcharts");
    const chart = convertChart(input);
    expect(chart.chartSpace.chart.plotArea.barChart?.barDir).toEqual({ val: "col" });
  });

  it("auto-detects amCharts", () => {
    const input = {
      type: "XYChart",
      data: [
        { category: "A", value: 10 },
        { category: "B", value: 20 },
      ],
      series: [{ type: "ColumnSeries", name: "Sales", valueYField: "value", categoryXField: "category" }],
    };
    expect(detectChartSource(input)).toBe("amcharts");
    const chart = convertChart(input);
    expect(chart.chartSpace.chart.plotArea.barChart?.ser[0]?.cat).toEqual(["A", "B"]);
    expect(chart.chartSpace.chart.plotArea.barChart?.ser[0]?.val).toEqual([10, 20]);
  });

  it("serializes convertChartToJson without XML", () => {
    const json = convertChartToJson(
      {
        type: "bar",
        categories: ["A", "B", "C"],
        series: [{ name: "Sales", values: [10, 20, 30] }],
      },
      { pretty: true },
    );
    expect(json).toContain('"type": "chart"');
    expect(json).toContain('"barChart"');
    expect(json).not.toContain("<c:chartSpace");
    expect(() => JSON.parse(json)).not.toThrow();
  });
});
