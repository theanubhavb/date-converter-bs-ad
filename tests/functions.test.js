const { adToBs, adToBsLong, bsToAd, bsToAdLong, bsDateDif } = require("../src/functions/functions");

describe("BS/AD custom functions", () => {
  test("converts the anchor date in both directions", () => {
    expect(adToBs("1943-04-14")).toBe("2000.01.01");
    expect(bsToAd("2000.01.01")).toBe("1943-04-14");
  });

  test("converts the supplied VBA example", () => {
    expect(adToBs("2014-01-01")).toBe("2070.09.17");
    expect(bsToAd("2072.04.01")).toBe("2015-07-17");
  });

  test("supports long values, accepted BS delimiters, and differences", () => {
    expect(adToBsLong("2024-04-13")).toBe("01 Baisakh 2081");
    expect(bsToAdLong("2081/01/01")).toBe("April 13, 2024");
    expect(bsDateDif("2081.01.02", "2081.01.01")).toBe(1);
  });

  test("returns a worksheet-friendly message for invalid input", () => {
    expect(bsToAd("2081.13.01")).toMatch(/not a valid day/);
    expect(adToBs("1900-01-01")).toMatch(/outside the supported range/);
  });
});
