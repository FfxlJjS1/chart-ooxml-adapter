# Office Open JSON (DemoMacro)

The output layer lives in `src/office/` under the name **OfficeOpenChartJson**. It is an object representation of OOXML ChartML, compatible with DemoMacro office-open consumers, and independent of any ZIP/PPTX generator.

## Files

| File | Responsibility |
| --- | --- |
| `OfficeOpenTypes.ts` | ChartML-shaped TypeScript types (`OfficeOpenChart`) |
| `OfficeOpenChartBuilder.ts` | `UniversalChartModel` → `OfficeOpenChart` |
| `OfficeOpenChartJson.ts` | JSON stringify/parse of the object tree |
| `../schemas/office-open-chart.schema.json` | Canonical JSON Schema |

## Object shape

Root type is always `"chart"`. The ChartML type is a plot-area key (`barChart`, `lineChart`, `pieChart`, …), not a sibling `chartType` string.

Series use the simplified DemoMacro form:

- `tx`: string
- `cat`: string array
- `val`: number array

not XML fragments such as `<c:tx>`.

Required ChartML entities the builder always writes:

- `chartSpace` preamble: `date1904`, `lang`, `roundedCorners`
- `chart`: `autoTitleDeleted`, `plotVisOnly`, `dispBlanksAs`
- plot group: `ser`, `varyColors`, and for cartesian charts `axId` plus `catAx`/`valAx`
- each series: `idx`, `order`, `tx`, and `cat`+`val` or `xVal`+`yVal`

## Validation

1. Build `UniversalChartModel` → `OfficeOpenChart`
2. Validate with `validateOfficeOpenChart` (required ChartML entities)
3. Validate with the JSON Schema (`getOfficeOpenChartSchema` / `chart-ooxml-adapter/schema`)

`assertOfficeOpenChart` throws `OfficeOpenValidationError` when either structural or ChartML-entity checks fail.

## DemoMacro projection

`toDemoMacroChartSpace(chart)` returns a `ChartSpaceOptions`-like object (`type`, `categories`, `series`, `title`, …) without PowerPoint/Excel/Word placement. Host generators add geometry.
