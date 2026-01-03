/**
 * Application State
 * 
 * This file contains all global state variables that change during runtime.
 * Import this file when you need to access or modify application state.
 */

/**
 * Current application flow mode
 * @type {"agreement" | "createTenantNew" | "viewTenants" | "generateBill" | "payments" | "exportData"}
 */
export let currentFlow = "agreement";

/**
 * Updates the current flow mode
 * @param {string} mode - The new flow mode
 */
export function setCurrentFlow(mode) {
    currentFlow = mode;
}

/**
 * ID of the last clause that was moved (for animation purposes)
 * @type {string | null}
 */
export let lastMovedClauseId = null;

/**
 * Direction of the last clause movement
 * @type {"up" | "down" | null}
 */
export let lastMoveDirection = null;

/**
 * Updates the last moved clause tracking
 * @param {string} id - Clause ID
 * @param {"up" | "down"} direction - Movement direction
 */
export function setLastMovedClause(id, direction) {
    lastMovedClauseId = id;
    lastMoveDirection = direction;
}

/**
 * Counter for generating unique clause IDs
 * @type {number}
 */
let clauseIdCounter = 0;

/**
 * Increments and returns a new clause ID counter value
 * @returns {number}
 */
export function incrementClauseIdCounter() {
    return ++clauseIdCounter;
}
