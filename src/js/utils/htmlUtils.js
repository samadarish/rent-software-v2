/**
 * HTML Utility Functions
 * 
 * Functions for manipulating HTML content and text formatting.
 */

/**
 * Strips HTML tags from a string and returns plain text
 * @param {string} html - HTML string to strip
 * @returns {string} Plain text without HTML tags
 */
export function stripHtml(html) {
    const tmp = document.createElement("div");
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || "";
}

/**
 * Escapes HTML special characters for safe string interpolation.
 * @param {string} value
 * @returns {string}
 */
export function escapeHtml(value) {
    const raw = value === null || value === undefined ? "" : String(value);
    return raw
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

/**
 * Applies bold formatting to the currently selected text within a container
 * @param {HTMLElement} container - The container element to apply bold formatting within
 */
export function applyBoldToSelection(container) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;

    const range = sel.getRangeAt(0);
    if (!container.contains(range.commonAncestorContainer)) return;
    if (range.collapsed) return; // nothing selected

    const strong = document.createElement("strong");
    strong.appendChild(range.extractContents());
    range.insertNode(strong);

    // Move caret after the bolded part
    sel.removeAllRanges();
    const newRange = document.createRange();
    newRange.setStartAfter(strong);
    newRange.setEndAfter(strong);
    sel.addRange(newRange);
}

/**
 * Converts contenteditable HTML into markdown-style **bold** markers
 * Replaces <strong> and <b> tags with **text** format
 * @param {string} html - HTML string with bold tags
 * @returns {string} Text with markdown-style bold markers
 */
export function htmlToMarkedText(html) {
    if (!html) return "";

    const tmp = document.createElement("div");
    tmp.innerHTML = html;

    tmp.querySelectorAll("strong, b").forEach((el) => {
        const fullText = el.textContent || "";
        if (!fullText) {
            el.remove();
            return;
        }

        const leadingSpaces = fullText.match(/^\s*/)[0];
        const trailingSpaces = fullText.match(/\s*$/)[0];
        const core = fullText.trim();

        const frag = document.createDocumentFragment();

        if (leadingSpaces) {
            frag.appendChild(document.createTextNode(leadingSpaces));
        }

        frag.appendChild(document.createTextNode(`**${core}**`));

        if (trailingSpaces) {
            frag.appendChild(document.createTextNode(trailingSpaces));
        }

        el.replaceWith(frag);
    });

    return tmp.textContent || tmp.innerText || "";
}
