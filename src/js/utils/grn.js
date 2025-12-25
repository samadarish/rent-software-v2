/**
 * Generates a NoGRN value with a 5-digit, zero-padded suffix.
 * @param {() => number} randomSource - Optional random source (0..1).
 * @returns {string}
 */
export function generateNoGrnValue(randomSource = Math.random) {
    const value = Math.floor(randomSource() * 100000);
    return `NoGRN${String(value).padStart(5, "0")}`;
}
