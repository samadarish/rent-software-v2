import { fetchLandlordsFromSheet, fetchUnitsFromSheet } from "../api/sheets.js";

let unitCache = [];
let landlordCache = [];
let unitsLoaded = false;
let landlordsLoaded = false;

function normalizeOccupiedFlag(raw) {
    if (raw === true || raw === false) return raw;
    if (typeof raw === "string") return raw.toLowerCase() === "true";
    return !!raw;
}

function normalizeUnit(raw) {
    if (!raw || typeof raw !== "object") return null;
    return {
        ...raw,
        is_occupied: normalizeOccupiedFlag(raw.is_occupied),
    };
}

export function getUnitCache() {
    return unitCache.slice();
}

export function getLandlordCache() {
    return landlordCache.slice();
}

export async function refreshUnits(force = false) {
    if (unitsLoaded && !force) return unitCache;
    const data = await fetchUnitsFromSheet(force);
    unitCache = Array.isArray(data.units)
        ? data.units.map(normalizeUnit).filter(Boolean)
        : [];
    unitsLoaded = true;
    document.dispatchEvent(new CustomEvent("units:updated", { detail: unitCache }));
    return unitCache;
}

export async function refreshLandlords(force = false) {
    if (landlordsLoaded && !force) return landlordCache;
    const data = await fetchLandlordsFromSheet(force);
    landlordCache = Array.isArray(data.landlords)
        ? data.landlords.map((l) => ({ ...l }))
        : [];
    landlordsLoaded = true;
    document.dispatchEvent(new CustomEvent("landlords:updated", { detail: landlordCache }));
    return landlordCache;
}

export async function ensureUnitsLoaded() {
    if (unitsLoaded) return unitCache;
    return refreshUnits();
}

export async function ensureLandlordsLoaded() {
    if (landlordsLoaded) return landlordCache;
    return refreshLandlords();
}

