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
