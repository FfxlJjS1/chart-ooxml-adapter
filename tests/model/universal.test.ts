import { describe, expect, it } from "vitest";
import { parseUniversalChartModel } from "../../src/model/UniversalChartModel.js";
import { InvalidUniversalModelError } from "../../src/errors.js";

describe("UniversalChartModel", () => {
  it("accepts the documented bar model", () => {
    const model = parseUniversalChartModel({
      type: "bar",
      categories: ["A", "B", "C"],
      series: [{ name: "Sales", values: [10, 20, 30] }],
    });
    expect(model.series[0]?.values).toEqual([10, 20, 30]);
  });

  it("rejects a values/categories length mismatch", () => {
    expect(() =>
      parseUniversalChartModel({
        type: "bar",
        categories: ["A", "B"],
        series: [{ name: "Sales", values: [10] }],
      }),
    ).toThrow(InvalidUniversalModelError);
  });
});
