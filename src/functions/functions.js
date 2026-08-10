/* global CustomFunctions */

"use strict";

// Year followed by Baisakh through Chaitra month lengths. Data is from the supplied VBA add-in.
  const BS_MONTHS = [
    [2000,30,32,31,32,31,30,30,30,29,30,29,31],[2001,31,31,32,31,31,31,30,29,30,29,30,30],[2002,31,31,32,32,31,30,30,29,30,29,30,30],[2003,31,32,31,32,31,30,30,30,29,29,30,31],[2004,30,32,31,32,31,30,30,30,29,30,29,31],[2005,31,31,32,31,31,31,30,29,30,29,30,30],[2006,31,31,32,32,31,30,30,29,30,29,30,30],[2007,31,32,31,32,31,30,30,30,29,29,30,31],[2008,31,31,31,32,31,31,29,30,30,29,29,31],[2009,31,31,32,31,31,31,30,29,30,29,30,30],
    [2010,31,31,32,32,31,30,30,29,30,29,30,30],[2011,31,32,31,32,31,30,30,30,29,29,30,31],[2012,31,31,31,32,31,31,29,30,30,29,30,30],[2013,31,31,32,31,31,31,30,29,30,29,30,30],[2014,31,31,32,32,31,30,30,29,30,29,30,30],[2015,31,32,31,32,31,30,30,30,29,29,30,31],[2016,31,31,31,32,31,31,29,30,30,29,30,30],[2017,31,31,32,31,31,31,30,29,30,29,30,30],[2018,31,32,31,32,31,30,30,29,30,29,30,30],[2019,31,32,31,32,31,30,30,30,29,30,29,31],
    [2020,31,31,31,32,31,31,30,29,30,29,30,30],[2021,31,31,32,31,31,31,30,29,30,29,30,30],[2022,31,32,31,32,31,30,30,30,29,29,30,30],[2023,31,32,31,32,31,30,30,30,29,30,29,31],[2024,31,31,31,32,31,31,30,29,30,29,30,30],[2025,31,31,32,31,31,31,30,29,30,29,30,30],[2026,31,32,31,32,31,30,30,30,29,29,30,31],[2027,30,32,31,32,31,30,30,30,29,30,29,31],[2028,31,31,32,31,31,31,30,29,30,29,30,30],[2029,31,31,32,31,32,30,30,29,30,29,30,30],
    [2030,31,32,31,32,31,30,30,30,29,29,30,31],[2031,30,32,31,32,31,30,30,30,29,30,29,31],[2032,31,31,32,31,31,31,30,29,30,29,30,30],[2033,31,31,32,32,31,30,30,29,30,29,30,30],[2034,31,32,31,32,31,30,30,30,29,29,30,31],[2035,30,32,31,32,31,31,29,30,30,29,29,31],[2036,31,31,32,31,31,31,30,29,30,29,30,30],[2037,31,31,32,32,31,30,30,29,30,29,30,30],[2038,31,32,31,32,31,30,30,30,29,29,30,31],[2039,31,31,31,32,31,31,29,30,30,29,30,30],
    [2040,31,31,32,31,31,31,30,29,30,29,30,30],[2041,31,31,32,32,31,30,30,29,30,29,30,30],[2042,31,32,31,32,31,30,30,30,29,29,30,31],[2043,31,31,31,32,31,31,29,30,30,29,30,30],[2044,31,31,32,31,31,31,30,29,30,29,30,30],[2045,31,32,31,32,31,30,30,29,30,29,30,30],[2046,31,32,31,32,31,30,30,30,29,29,30,31],[2047,31,31,31,32,31,31,30,29,30,29,30,30],[2048,31,31,32,31,31,31,30,29,30,29,30,30],[2049,31,32,31,32,31,30,30,30,29,29,30,30],
    [2050,31,32,31,32,31,30,30,30,29,30,29,31],[2051,31,31,31,32,31,31,30,29,30,29,30,30],[2052,31,31,32,31,31,31,30,29,30,29,30,30],[2053,31,32,31,32,31,30,30,30,29,29,30,30],[2054,31,32,31,32,31,30,30,30,29,30,29,31],[2055,31,31,32,31,31,31,30,29,30,29,30,30],[2056,31,31,32,31,32,30,30,29,30,29,30,30],[2057,31,32,31,32,31,30,30,30,29,29,30,31],[2058,30,32,31,32,31,30,30,30,29,30,29,31],[2059,31,31,32,31,31,31,30,29,30,29,30,30],
    [2060,31,31,32,32,31,30,30,29,30,29,30,30],[2061,31,32,31,32,31,30,30,30,29,29,30,31],[2062,30,32,31,32,31,31,29,30,29,30,29,31],[2063,31,31,32,31,31,31,30,29,30,29,30,30],[2064,31,31,32,32,31,30,30,29,30,29,30,30],[2065,31,32,31,32,31,30,30,30,29,29,30,31],[2066,31,31,31,32,31,31,29,30,30,29,29,31],[2067,31,31,32,31,31,31,30,29,30,29,30,30],[2068,31,31,32,32,31,30,30,29,30,29,30,30],[2069,31,32,31,32,31,30,30,30,29,29,30,31],
    [2070,31,31,31,32,31,31,29,30,30,29,30,30],[2071,31,31,32,31,31,31,30,29,30,29,30,30],[2072,31,32,31,32,31,30,30,29,30,29,30,30],[2073,31,32,31,32,31,30,30,30,29,29,30,31],[2074,31,31,31,32,31,31,30,29,30,29,30,30],[2075,31,31,32,31,31,31,30,29,30,29,30,30],[2076,31,32,31,32,31,30,30,30,29,29,30,30],[2077,31,32,31,32,31,30,30,30,29,30,29,31],[2078,31,31,31,32,31,31,30,29,30,29,30,30],[2079,31,31,32,31,31,31,30,29,30,29,30,30],
    [2080,31,32,31,32,31,30,30,30,29,29,30,30],[2081,31,32,31,32,31,30,30,30,29,30,29,31],[2082,31,31,32,31,31,31,30,29,30,29,30,30],[2083,31,31,32,31,31,31,30,29,30,29,30,30],[2084,31,31,32,31,31,30,30,30,29,30,30,30],[2085,31,32,31,32,30,31,30,30,29,30,30,30],[2086,30,32,31,32,31,30,30,30,29,30,30,30],[2087,31,31,32,31,31,31,30,30,29,30,30,30],[2088,30,31,32,32,30,31,30,30,29,30,30,30],[2089,30,32,31,32,31,30,30,30,29,30,30,30],
    [2090,30,32,31,32,31,30,30,30,29,30,30,30]
  ];

  const MONTH_NAMES = ["Baisakh", "Jestha", "Ashad", "Shrawan", "Bhadra", "Ashwin", "Kartik", "Mangsir", "Paush", "Magh", "Falgun", "Chaitra"];
  const MS_PER_DAY = 86400000;
  const ANCHOR_UTC = Date.UTC(1943, 3, 14); // 2000.01.01 BS
  const yearOffsets = [];
  let cumulative = 0;
  BS_MONTHS.forEach((year) => { yearOffsets.push(cumulative); cumulative += year.slice(1).reduce((sum, days) => sum + days, 0); });

  function pad(value) { return String(value).padStart(2, "0"); }
  function formatBs(year, month, day) { return `${year}.${pad(month)}.${pad(day)}`; }
  function validBs(year, month, day) { const entry = BS_MONTHS[year - 2000]; return entry && month >= 1 && month <= 12 && day >= 1 && day <= entry[month]; }
  function parseBs(value) {
    const match = String(value).trim().match(/^(\d{4})[.\/-]?(\d{1,2})[.\/-]?(\d{1,2})$/);
    if (!match) throw new Error("Enter a BS date as YYYYMMDD, YYYY.MM.DD, YYYY-MM-DD, or YYYY/MM/DD.");
    const [, year, month, day] = match.map(Number);
    if (!validBs(year, month, day)) throw new Error("The BS date is outside the supported range or is not a valid day.");
    return { year, month, day };
  }
  function parseAd(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return Date.UTC(value.getFullYear(), value.getMonth(), value.getDate());
    if (typeof value === "number") return Date.UTC(1899, 11, 30) + Math.floor(value) * MS_PER_DAY;
    const match = String(value).trim().match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
    if (!match) throw new Error("Enter an AD date as YYYY-MM-DD (or supply an Excel date cell).");
    const [, year, month, day] = match.map(Number);
    const utc = Date.UTC(year, month - 1, day);
    const check = new Date(utc);
    if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) throw new Error("The AD date is not valid.");
    return utc;
  }
  function adToBsParts(adDate) {
    const offset = Math.floor((parseAd(adDate) - ANCHOR_UTC) / MS_PER_DAY);
    if (offset < 0 || offset >= cumulative) throw new Error("The AD date is outside the supported range.");
    const index = BS_MONTHS.findIndex((_, i) => offset < yearOffsets[i] + BS_MONTHS[i].slice(1).reduce((sum, days) => sum + days, 0));
    let remaining = offset - yearOffsets[index];
    const lengths = BS_MONTHS[index];
    let month = 1;
    while (remaining >= lengths[month]) { remaining -= lengths[month]; month += 1; }
    return { year: lengths[0], month, day: remaining + 1 };
  }
  function bsToAdUtc(bsDate) {
    const { year, month, day } = parseBs(bsDate);
    const index = year - 2000;
    const offset = yearOffsets[index] + BS_MONTHS[index].slice(1, month).reduce((sum, days) => sum + days, 0) + day - 1;
    return ANCHOR_UTC + offset * MS_PER_DAY;
  }
  function formatAd(utc) { const date = new Date(utc); return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`; }
  function formatAdLong(utc) { return new Intl.DateTimeFormat("en-US", { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" }).format(new Date(utc)); }
  function excel(fn) { try { return fn(); } catch (error) { return error.message; } }

  /**
   * Converts an AD/Gregorian date to a BS date in YYYY.MM.DD format.
   * @customfunction ADTOBS ADTOBS
   * @param {any} adDate An Excel date or YYYY-MM-DD date.
   * @returns {string} The equivalent BS date.
   */
  function adToBs(adDate) { return excel(() => { const { year, month, day } = adToBsParts(adDate); return formatBs(year, month, day); }); }
  /**
   * Converts an AD/Gregorian date to a long BS date.
   * @customfunction ADTOBSLONG ADTOBSLONG
   * @param {any} adDate An Excel date or YYYY-MM-DD date.
   * @returns {string} The equivalent BS date, such as 01 Baisakh 2081.
   */
  function adToBsLong(adDate) { return excel(() => { const { year, month, day } = adToBsParts(adDate); return `${pad(day)} ${MONTH_NAMES[month - 1]} ${year}`; }); }
  /**
   * Converts a BS date to an AD/Gregorian date in YYYY-MM-DD format.
   * @customfunction BSTOAD BSTOAD
   * @param {string} bsDate A BS date in YYYYMMDD, YYYY.MM.DD, YYYY-MM-DD, or YYYY/MM/DD.
   * @returns {string} The equivalent AD date.
   */
  function bsToAd(bsDate) { return excel(() => formatAd(bsToAdUtc(bsDate))); }
  /**
   * Converts a BS date to a long AD/Gregorian date.
   * @customfunction BSTOADLONG BSTOADLONG
   * @param {string} bsDate A BS date in YYYYMMDD, YYYY.MM.DD, YYYY-MM-DD, or YYYY/MM/DD.
   * @returns {string} The equivalent AD date, such as April 13, 2024.
   */
  function bsToAdLong(bsDate) { return excel(() => formatAdLong(bsToAdUtc(bsDate))); }
  /**
   * Returns the signed whole-day difference between two BS dates (first minus second).
   * @customfunction BSDATEDIF BSDATEDIF
   * @param {string} bsDate1 First BS date.
   * @param {string} bsDate2 Second BS date.
   * @returns {number} Days between the dates.
   */
  function bsDateDif(bsDate1, bsDate2) { return excel(() => Math.round((bsToAdUtc(bsDate1) - bsToAdUtc(bsDate2)) / MS_PER_DAY)); }

if (typeof CustomFunctions !== "undefined") {
  CustomFunctions.associate("ADTOBS", adToBs);
  CustomFunctions.associate("ADTOBSLONG", adToBsLong);
  CustomFunctions.associate("BSTOAD", bsToAd);
  CustomFunctions.associate("BSTOADLONG", bsToAdLong);
  CustomFunctions.associate("BSDATEDIF", bsDateDif);
}
if (typeof module !== "undefined") module.exports = { adToBs, adToBsLong, bsToAd, bsToAdLong, bsDateDif };
