import type { ConvertOptions } from "../adapters/types.js";
import {
  CARTESIAN_CHART_TYPES,
  SCATTER_LIKE_TYPES,
  type UniversalChartModel,
  type UniversalChartType,
  type UniversalGrouping,
  type UniversalLegendPosition,
  type UniversalSeries,
} from "../model/types.js";
import { parseUniversalChartModel } from "../model/UniversalChartModel.js";
import {
  type OfficeBarChart,
  type OfficeBarDirection,
  type OfficeChartGroup,
  type OfficeChartSpace,
  type OfficeGrouping,
  type OfficeLegendPosition,
  type OfficeOpenChart,
  type OfficePlotArea,
  type OfficePlotChartKey,
  type OfficeSeries,
  type OfficeUnsignedInt,
} from "./OfficeOpenTypes.js";

const CAT_AX_ID = 1;
const VAL_AX_ID = 2;

const PLOT_KEY_BY_TYPE: Record<UniversalChartType, OfficePlotChartKey> = {
  bar: "barChart",
  column: "barChart",
  line: "lineChart",
  area: "areaChart",
  pie: "pieChart",
  doughnut: "doughnutChart",
  scatter: "scatterChart",
  radar: "radarChart",
  bubble: "bubbleChart",
};

const PLOT_KEY_BY_TYPE_3D: Partial<Record<UniversalChartType, OfficePlotChartKey>> = {
  bar: "bar3DChart",
  column: "bar3DChart",
  line: "line3DChart",
  area: "area3DChart",
  pie: "pie3DChart",
};

function onOff(val: boolean): { val: boolean } {
  return { val };
}

function uInt(val: number): OfficeUnsignedInt {
  return { val };
}

function val<T extends string | number>(value: T): { val: T } {
  return { val: value };
}

function defaultGrouping(type: UniversalChartType, grouping?: UniversalGrouping): OfficeGrouping {
  if (grouping) {
    return grouping;
  }
  if (type === "bar" || type === "column") {
    return "clustered";
  }
  return "standard";
}

function legendPos(position: UniversalLegendPosition | undefined): OfficeLegendPosition {
  switch (position) {
    case "left":
      return "l";
    case "top":
      return "t";
    case "bottom":
      return "b";
    default:
      return "r";
  }
}

function hexColor(color: string | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  const match = /^#?([0-9a-fA-F]{6})$/u.exec(color.trim());
  return match?.[1]?.toUpperCase();
}

function toOfficeSeries(
  model: UniversalChartModel,
  series: UniversalSeries,
  index: number,
): OfficeSeries {
  const officeSeries: OfficeSeries = {
    idx: uInt(index),
    order: uInt(index),
    tx: series.name,
  };

  if (SCATTER_LIKE_TYPES.has(model.type)) {
    officeSeries.xVal = [...(series.xValues ?? [])];
    officeSeries.yVal = [...(series.yValues ?? series.values ?? [])];
    if (model.type === "bubble") {
      officeSeries.bubbleSize = [...(series.bubbleSizes ?? [])];
    }
  } else {
    officeSeries.cat = [...model.categories];
    officeSeries.val = [...(series.values ?? [])];
  }

  const color = hexColor(series.color);
  if (color) {
    officeSeries.spPr = { solidFill: { srgbClr: val(color) } };
  }

  if (model.type === "bar" || model.type === "column") {
    officeSeries.invertIfNegative = onOff(false);
  }

  return officeSeries;
}

function axPair(): [OfficeUnsignedInt, OfficeUnsignedInt] {
  return [uInt(CAT_AX_ID), uInt(VAL_AX_ID)];
}

function scatterAxPair(): [OfficeUnsignedInt, OfficeUnsignedInt] {
  return [uInt(CAT_AX_ID), uInt(VAL_AX_ID)];
}

function buildChartGroup(model: UniversalChartModel, ser: OfficeSeries[]): OfficeChartGroup {
  const grouping = val(defaultGrouping(model.type, model.grouping));
  switch (model.type) {
    case "bar":
    case "column": {
      const barDir: OfficeBarDirection = model.type === "bar" ? "bar" : "col";
      const barChart: OfficeBarChart = {
        barDir: val(barDir),
        grouping,
        varyColors: onOff(false),
        ser,
        gapWidth: uInt(150),
        overlap: val(0),
        axId: axPair(),
      };
      return barChart;
    }
    case "line":
      return {
        grouping,
        varyColors: onOff(false),
        ser,
        marker: onOff(true),
        smooth: onOff(false),
        axId: axPair(),
      };
    case "area":
      return {
        grouping,
        varyColors: onOff(false),
        ser,
        axId: axPair(),
      };
    case "pie":
      return {
        varyColors: onOff(true),
        ser,
        firstSliceAng: uInt(0),
      };
    case "doughnut":
      return {
        varyColors: onOff(true),
        ser,
        firstSliceAng: uInt(0),
        holeSize: uInt(50),
      };
    case "scatter":
      return {
        scatterStyle: val("marker"),
        varyColors: onOff(false),
        ser,
        axId: scatterAxPair(),
      };
    case "radar":
      return {
        radarStyle: val("marker"),
        varyColors: onOff(false),
        ser,
        axId: axPair(),
      };
    case "bubble":
      return {
        varyColors: onOff(false),
        ser,
        bubble3D: onOff(false),
        axId: scatterAxPair(),
      };
  }
}

function plotKey(model: UniversalChartModel): OfficePlotChartKey {
  if (model.threeD) {
    return PLOT_KEY_BY_TYPE_3D[model.type] ?? PLOT_KEY_BY_TYPE[model.type];
  }
  return PLOT_KEY_BY_TYPE[model.type];
}

function buildAxes(model: UniversalChartModel): Pick<OfficePlotArea, "catAx" | "valAx"> {
  if (!CARTESIAN_CHART_TYPES.has(model.type)) {
    return {};
  }

  const isBar = model.type === "bar";
  const isScatter = SCATTER_LIKE_TYPES.has(model.type);

  if (isScatter) {
    return {
      valAx: [
        {
          axId: uInt(CAT_AX_ID),
          scaling: {
            orientation: val("minMax"),
            min: model.xAxis?.min !== undefined ? val(model.xAxis.min) : undefined,
            max: model.xAxis?.max !== undefined ? val(model.xAxis.max) : undefined,
          },
          axPos: val("b"),
          majorTickMark: val("out"),
          minorTickMark: val("none"),
          tickLblPos: val("nextTo"),
          crossAx: uInt(VAL_AX_ID),
          title: model.xAxis?.title ? { tx: model.xAxis.title } : undefined,
        },
        {
          axId: uInt(VAL_AX_ID),
          scaling: {
            orientation: val("minMax"),
            min: model.yAxis?.min !== undefined ? val(model.yAxis.min) : undefined,
            max: model.yAxis?.max !== undefined ? val(model.yAxis.max) : undefined,
          },
          axPos: val("l"),
          majorGridlines: {},
          numFmt: { formatCode: "General", sourceLinked: true },
          majorTickMark: val("out"),
          minorTickMark: val("none"),
          tickLblPos: val("nextTo"),
          crossAx: uInt(CAT_AX_ID),
          title: model.yAxis?.title ? { tx: model.yAxis.title } : undefined,
        },
      ],
    };
  }

  return {
    catAx: [
      {
        axId: uInt(CAT_AX_ID),
        scaling: { orientation: val("minMax") },
        axPos: val(isBar ? "l" : "b"),
        majorTickMark: val("out"),
        minorTickMark: val("none"),
        tickLblPos: val("nextTo"),
        crossAx: uInt(VAL_AX_ID),
        title: model.xAxis?.title ? { tx: model.xAxis.title } : undefined,
      },
    ],
    valAx: [
      {
        axId: uInt(VAL_AX_ID),
        scaling: {
          orientation: val("minMax"),
          min: model.yAxis?.min !== undefined ? val(model.yAxis.min) : undefined,
          max: model.yAxis?.max !== undefined ? val(model.yAxis.max) : undefined,
        },
        axPos: val(isBar ? "b" : "l"),
        majorGridlines: {},
        numFmt: { formatCode: "General", sourceLinked: true },
        majorTickMark: val("out"),
        minorTickMark: val("none"),
        tickLblPos: val("nextTo"),
        crossAx: uInt(CAT_AX_ID),
        title: model.yAxis?.title ? { tx: model.yAxis.title } : undefined,
      },
    ],
  };
}

function buildPlotArea(model: UniversalChartModel): OfficePlotArea {
  const ser = model.series.map((item, index) => toOfficeSeries(model, item, index));
  const group = buildChartGroup(model, ser);
  const key = plotKey(model);
  const axes = buildAxes(model);
  return {
    layout: {},
    [key]: group,
    ...axes,
  } as OfficePlotArea;
}

function buildChartSpace(model: UniversalChartModel, options?: ConvertOptions): OfficeChartSpace {
  const showLegend = options?.showLegend ?? model.legend?.show ?? true;
  const hasTitle = Boolean(model.title);
  return {
    date1904: onOff(false),
    lang: val(options?.lang ?? "en-US"),
    roundedCorners: onOff(false),
    style: uInt(options?.style ?? 2),
    chart: {
      title: hasTitle ? { tx: model.title as string, overlay: onOff(false) } : undefined,
      autoTitleDeleted: onOff(!hasTitle),
      plotArea: buildPlotArea(model),
      legend: showLegend
        ? {
            legendPos: val(legendPos(model.legend?.position)),
            overlay: onOff(false),
          }
        : undefined,
      plotVisOnly: onOff(true),
      dispBlanksAs: val("gap"),
      showDLblsOverMax: onOff(false),
    },
  };
}

/**
 * Builds a DemoMacro-compatible Office Open JSON chart object from a
 * UniversalChartModel. Does not create packages, relationships, media, or XML.
 */
export function buildOfficeOpenChart(
  model: UniversalChartModel,
  options?: ConvertOptions,
): OfficeOpenChart {
  const normalized = parseUniversalChartModel(model);
  return {
    type: "chart",
    chartSpace: buildChartSpace(normalized, options),
  };
}

export class OfficeOpenChartBuilder {
  private model: UniversalChartModel;
  private options: ConvertOptions;

  constructor(model: UniversalChartModel, options: ConvertOptions = {}) {
    this.model = parseUniversalChartModel(model);
    this.options = { ...options };
  }

  static fromUniversalModel(
    model: UniversalChartModel,
    options?: ConvertOptions,
  ): OfficeOpenChartBuilder {
    return new OfficeOpenChartBuilder(model, options);
  }

  withLang(lang: string): this {
    this.options.lang = lang;
    return this;
  }

  withLegend(show: boolean): this {
    this.options.showLegend = show;
    return this;
  }

  withStyle(style: number): this {
    this.options.style = style;
    return this;
  }

  build(): OfficeOpenChart {
    return buildOfficeOpenChart(this.model, this.options);
  }
}
