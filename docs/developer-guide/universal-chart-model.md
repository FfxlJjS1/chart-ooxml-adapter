# UniversalChartModel

`UniversalChartModel` is the host-neutral chart model produced by source adapters. The Office Open JSON builder accepts only this model.

## Required fields

| Field | Role |
| --- | --- |
| `type` | `bar`, `column`, `line`, `area`, `pie`, `doughnut`, `scatter`, `radar`, `bubble` |
| `categories` | Category labels. Required and non-empty for category charts. Empty for scatter/bubble. |
| `series` | Non-empty array. Each series has a `name` plus `values` or `xValues`/`yValues`. |

`bar` is a horizontal bar chart (`c:barChart` + `barDir=bar`). `column` is a vertical column chart (`c:barChart` + `barDir=col`). Chart.js `type: "bar"` maps to `column` unless `options.indexAxis` is `"y"`.

## Optional fields

- `title`, `grouping`, `legend`, `xAxis`, `yAxis`, `threeD`, `source`

Parse unknown JSON with `parseUniversalChartModel`. The function fails when series lengths do not match categories.
