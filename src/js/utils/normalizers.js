import { normalizeMonthKey as normalizeMonthKeyBase } from "./formatters.js";

export function normalizeKey(value, options = {}) {
    const { lowercase = true } = options;
    const str = (value ?? "").toString().trim();
    return lowercase ? str.toLowerCase() : str;
}

export function normalizeWing(value) {
    return normalizeKey(value);
}

export function normalizeMonthKey(value) {
    return normalizeMonthKeyBase(value, { lowercase: true });
}

