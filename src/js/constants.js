/**
 * Application Constants
 * 
 * This file contains all constant values used throughout the application.
 * These values do not change during runtime.
 */

/**
 * Local storage keys for persisting user preferences
 */
export const STORAGE_KEYS = {
    APP_SCRIPT_URL: "tenantApp.appscriptUrl",
    LANDLORD_DEFAULTS: "tenantApp.landlordDefaults",
    LAST_FULL_SYNC_AT: "tenantApp.lastFullSyncAt",
};

/**
 * Clause sections structure
 * Each section contains a label and items array that will be populated from Google Sheets
 */
export const clauseSections = {
    tenant: {
        label: "Tenant Responsibilities",
        items: [],
    },
    landlord: {
        label: "Landlord Responsibilities",
        items: [],
    },
    penalties: {
        label: "Penalties",
        items: [],
    },
    misc: {
        label: "Miscellaneous",
        items: [],
    },
};

/**
 * Available floor options for the premises
 */
export const floorOptions = [
    "Ground Floor",
    "1st Floor",
    "2nd Floor",
    "3rd Floor",
    "4th Floor",
    "5th Floor",
    "6th Floor",
    "7th Floor",
    "8th Floor",
    "9th Floor",
    "10th Floor",
];
