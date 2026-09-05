# Converting charts

Use `convertChart` when you have either a chart-library option object or a `UniversalChartModel`.

## From UniversalChartModel

```ts
import { convertChart, convertChartToJson } from "chart-ooxml-adapter";

const model = {
  type: "bar",
  categories: ["A", "B", "C"],
  series: [{ name: "Sales", values: [10, 20, 30] }],
};

const chart = convertChart(model);
const json = convertChartToJson(model, { pretty: true });
```

The `bar` model above becomes:

```json
{
  "type": "chart",
  "chartSpace": {
    "chart": {
      "plotArea": {
        "barChart": {
          "ser": [
            {
              "tx": "Sales",
              "cat": ["A", "B", "C"],
              "val": [10, 20, 30]
            }
          ]
        }
      }
    }
  }
}
```

The produced object also includes required ChartML entities (`idx`, `order`, `barDir`, `axId`, axes, `plotVisOnly`, `lang`, and so on). Those fields are omitted from the sketch above only for readability.

## From a chart library

```ts
const chart = convertChart(echartsOption);
const chartFromHighcharts = convertChart(highchartsConfig, { source: "highcharts" });
```

Supported automatic sources: `chartjs`, `echarts`, `highcharts`, `amcharts`, and `universal`.

## Passing the result downstream

Hand the object to an Office package generator. For DemoMacro `@office-open/core` consumers, project it with `toDemoMacroChartSpace(chart)` and let the host library add placement (`x`, `y`, `width`, `height`, worksheet anchor). Do not expect this adapter to write a `.pptx` file.
