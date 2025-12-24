/**
 * Formatting Utility Functions
 * 
 * Functions for formatting dates, numbers, and text.
 */

/**
 * Converts a number to its ordinal form (1st, 2nd, 3rd, etc.)
 * @param {number} n - The number to convert
 * @returns {string} The ordinal representation
 */
export function toOrdinal(n) {
    n = parseInt(n, 10);
    if (isNaN(n)) return "";
    const s = ["th", "st", "nd", "rd"];
    const v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

/**
 * Formats an ISO date string into a human-readable format
 * @param {string} isoDate - ISO date string (YYYY-MM-DD)
 * @returns {string} Formatted date (e.g., "1st Jan 2025")
 */
export function formatDateForDoc(isoDate) {
    if (!isoDate) return "";
    const d = new Date(isoDate + "T00:00:00");
    if (isNaN(d.getTime())) return "";
    const day = d.getDate();
    const monthNames = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ];
    const month = monthNames[d.getMonth()];
    const year = d.getFullYear();
    return `${toOrdinal(day)} ${month} ${year}`;
}

/**
 * Converts a number to Indian English words
 * Supports crore, lakh, thousand, hundred system
 * @param {number} num - The number to convert
 * @returns {string} Number in words (e.g., "One lakh twenty-three thousand")
 */
export function numberToIndianWords(num) {
    if (num === 0) return "zero";

    const ones = [
        "",
        "one",
        "two",
        "three",
        "four",
        "five",
        "six",
        "seven",
        "eight",
        "nine",
        "ten",
        "eleven",
        "twelve",
        "thirteen",
        "fourteen",
        "fifteen",
        "sixteen",
        "seventeen",
        "eighteen",
        "nineteen",
    ];

    const tens = [
        "",
        "",
        "twenty",
        "thirty",
        "forty",
        "fifty",
        "sixty",
        "seventy",
        "eighty",
        "ninety",
    ];

    /**
     * Converts a two-digit number to words
     * @param {number} n - Number between 0-99
     * @returns {string}
     */
    function twoDigits(n) {
        if (n < 20) return ones[n];
        const t = Math.floor(n / 10);
        const o = n % 10;
        return tens[t] + (o ? " " + ones[o] : "");
    }

    /**
     * Converts a three-digit number to words
     * @param {number} n - Number between 0-999
     * @returns {string}
     */
    function threeDigits(n) {
        const h = Math.floor(n / 100);
        const r = n % 100;
        let str = "";
        if (h) {
            str += ones[h] + " hundred";
            if (r) str += " ";
        }
        if (r) str += twoDigits(r);
        return str;
    }

    let result = "";

    const crore = Math.floor(num / 10000000);
    num = num % 10000000;
    const lakh = Math.floor(num / 100000);
    num = num % 100000;
    const thousand = Math.floor(num / 1000);
    num = num % 1000;
    const hundred = num;

    if (crore) result += threeDigits(crore) + " crore";
    if (lakh) {
        if (result) result += " ";
        result += threeDigits(lakh) + " lakh";
    }
    if (thousand) {
        if (result) result += " ";
        result += threeDigits(thousand) + " thousand";
    }
    if (hundred) {
        if (result) result += " ";
        result += threeDigits(hundred);
    }

    return result.charAt(0).toUpperCase() + result.slice(1);
}

const DEFAULT_CURRENCY_SYMBOL = "\u20B9";

/**
 * Normalizes month identifiers into a YYYY-MM key when possible.
 * @param {string|Date|number} value
 * @param {{ lowercase?: boolean }} options
 * @returns {string}
 */
export function normalizeMonthKey(value, options = {}) {
    const { lowercase = false } = options;

    if (!value) return "";

    if (value instanceof Date && !Number.isNaN(value)) {
        const month = `${value.getUTCMonth() + 1}`.padStart(2, "0");
        const result = `${value.getUTCFullYear()}-${month}`;
        return lowercase ? result.toLowerCase() : result;
    }

    const str = value.toString().trim();
    if (!str) return "";

    // Excel / Sheets serial numbers (roughly covers 1990-2150 ranges)
    const numericValue = Number(str);
    if (!Number.isNaN(numericValue) && /^\d+(?:\.0+)?$/.test(str) && numericValue > 30000) {
        const excelEpoch = Date.UTC(1899, 11, 30);
        const utcDate = new Date(excelEpoch + numericValue * 24 * 60 * 60 * 1000);
        if (!Number.isNaN(utcDate)) {
            const month = `${utcDate.getUTCMonth() + 1}`.padStart(2, "0");
            const result = `${utcDate.getUTCFullYear()}-${month}`;
            return lowercase ? result.toLowerCase() : result;
        }
    }

    const compact = str.match(/^(\d{4})(\d{2})$/);
    if (compact) {
        const result = `${compact[1]}-${compact[2]}`;
        return lowercase ? result.toLowerCase() : result;
    }

    const ymd = str.match(/^(\d{4})[-/.](\d{1,2})/);
    if (ymd) {
        const month = `${ymd[2]}`.padStart(2, "0");
        const result = `${ymd[1]}-${month}`;
        return lowercase ? result.toLowerCase() : result;
    }

    const my = str.match(/^(\d{1,2})[-/.](\d{4})$/);
    if (my) {
        const month = `${my[1]}`.padStart(2, "0");
        const result = `${my[2]}-${month}`;
        return lowercase ? result.toLowerCase() : result;
    }

    const mdy = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (mdy) {
        const rawYear = mdy[3].length === 2 ? `20${mdy[3]}` : mdy[3];
        const first = parseInt(mdy[1], 10);
        const second = parseInt(mdy[2], 10);
        const monthNum = first > 12 && second <= 12 ? second : first;
        const month = `${monthNum}`.padStart(2, "0");
        const result = `${rawYear}-${month}`;
        return lowercase ? result.toLowerCase() : result;
    }

    const monthName = str.match(/^([A-Za-z]{3,})\s+(\d{2,4})$/);
    if (monthName) {
        const monthIdx = new Date(`${monthName[1]} 1, 2000`).getMonth();
        if (!Number.isNaN(monthIdx)) {
            const month = `${monthIdx + 1}`.padStart(2, "0");
            const year = monthName[2].length === 2 ? `20${monthName[2]}` : monthName[2];
            const result = `${year}-${month}`;
            return lowercase ? result.toLowerCase() : result;
        }
    }

    const parsedDate = new Date(str);
    if (!Number.isNaN(parsedDate)) {
        const month = `${parsedDate.getUTCMonth() + 1}`.padStart(2, "0");
        const result = `${parsedDate.getUTCFullYear()}-${month}`;
        return lowercase ? result.toLowerCase() : result;
    }

    return lowercase ? str.toLowerCase() : str;
}

/**
 * Formats a number as INR currency text.
 * @param {number|string} value
 * @param {{ currencySymbol?: string, emptyValue?: string, invalidValue?: string, rawOnInvalid?: boolean, parseMode?: "float"|"int", minimumFractionDigits?: number, maximumFractionDigits?: number, useGrouping?: boolean, roundTo?: number, coerceEmptyToZero?: boolean }} options
 * @returns {string}
 */
export function formatCurrency(value, options = {}) {
    const {
        currencySymbol = DEFAULT_CURRENCY_SYMBOL,
        emptyValue = "",
        invalidValue = "",
        rawOnInvalid = false,
        parseMode = "float",
        minimumFractionDigits = 2,
        maximumFractionDigits = 2,
        useGrouping = true,
        roundTo = null,
        coerceEmptyToZero = false,
    } = options;

    if (value === null || value === undefined || value === "") {
        if (coerceEmptyToZero) {
            value = 0;
        } else {
            return emptyValue;
        }
    }

    const numeric =
        parseMode === "int" ? parseInt(value, 10) : Number(value);

    if (Number.isNaN(numeric)) {
        if (rawOnInvalid) return `${currencySymbol}${String(value ?? "")}`;
        if (invalidValue) return invalidValue;
        return emptyValue;
    }

    let normalized = numeric;
    if (typeof roundTo === "number") {
        const factor = Math.pow(10, roundTo);
        normalized = Math.round(normalized * factor) / factor;
    }

    const formatted = useGrouping
        ? normalized.toLocaleString("en-IN", {
              minimumFractionDigits,
              maximumFractionDigits,
          })
        : maximumFractionDigits > 0
        ? normalized.toFixed(maximumFractionDigits)
        : `${Math.round(normalized)}`;

    return `${currencySymbol}${formatted}`;
}

/**
 * Builds a label for a unit using wing and unit number.
 * @param {object} unit
 * @returns {string}
 */
export function buildUnitLabel(unit) {
    if (!unit) return "";
    const wing = unit.wing || "";
    const unitNumber = unit.unit_number || unit.unitNumber || "";
    const label = [wing, unitNumber].filter(Boolean).join(" - ");
    if (label) return label;
    return unit.unit_id || unit.unitId || "";
}
