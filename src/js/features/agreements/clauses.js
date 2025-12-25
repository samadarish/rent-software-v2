/**
 * Clauses Feature Module
 * 
 * Handles all logic related to clause management:
 * - Rendering the clause UI
 * - Normalizing and sorting clauses
 * - Managing clause state (dirty tracking, animations)
 * - Generating unique clause IDs
 */

import { clauseSections } from "../../constants.js";
import {
    setClausesDirtyState,
    lastMovedClauseId,
    lastMoveDirection,
    incrementClauseIdCounter,
    setLastMovedClause
} from "../../state.js";
import { stripHtml, applyBoldToSelection } from "../../utils/htmlUtils.js";

/**
 * Generates a unique ID for a new clause
 * @returns {string} Unique clause ID
 */
function generateClauseId() {
    return `clause-${Date.now()}-${incrementClauseIdCounter()}`;
}

/**
 * Updates the clauses dirty state and shows/hides the sync button
 * @param {boolean} dirty - Whether clauses have been modified
 */
export function setClausesDirty(dirty = true) {
    setClausesDirtyState(dirty);

    const syncBtn = document.getElementById("saveClausesBtn");
    const dirtyMsg = document.getElementById("clausesDirtyMessage");

    if (syncBtn) {
        if (dirty) {
            syncBtn.classList.remove("hidden");
        } else {
            syncBtn.classList.add("hidden");
        }
    }

    if (dirtyMsg) {
        if (dirty) {
            dirtyMsg.classList.remove("hidden");
        } else {
            dirtyMsg.classList.add("hidden");
        }
    }
}

/**
 * Normalizes clause sections by ensuring consistent data structure
 * Sets default values for missing properties and sorts clauses by sortOrder
 */
export function normalizeClauseSections() {
    Object.values(clauseSections).forEach((section) => {
        section.items = section.items.map((item, index) => {
            const html = item.html || "";

            const text =
                item.text && item.text.trim()
                    ? item.text
                    : stripHtml(html || "");

            let sortOrder = Number(
                typeof item.sortOrder === "undefined"
                    ? item.sort_order
                    : item.sortOrder
            );
            if (!Number.isFinite(sortOrder) || sortOrder <= 0) {
                sortOrder = index + 1;
            }

            const enabled =
                typeof item.enabled === "boolean"
                    ? item.enabled
                    : String(item.enabled).toLowerCase() !== "false";

            const id = item.id || generateClauseId();

            return {
                id,
                text,
                html,
                enabled,
                sortOrder,
            };
        });

        section.items.sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0));
    });
}

/**
 * Renders the clause UI with all sections and items
 * Creates editable clause lists with controls for reordering, toggling, and editing
 */
export function renderClausesUI() {
    const container = document.getElementById("clausesContainer");
    if (!container) return;
    container.innerHTML = "";

    Object.entries(clauseSections).forEach(([, section]) => {
        const box = document.createElement("div");
        box.className =
            "border border-slate-200 rounded-lg p-2 space-y-2 bg-white/80";

        const header = document.createElement("div");
        header.className = "flex items-center justify-between";
        header.innerHTML = `<span class="text-xs font-semibold text-slate-700">${section.label}</span>`;
        box.appendChild(header);

        const list = document.createElement("ul");
        list.className = "space-y-1";

        section.items
            .sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0))
            .forEach((item) => {
                const li = document.createElement("li");
                li.className =
                    "flex items-start gap-2 text-[11px] rounded-md px-1 py-1 transition-colors";

                // Apply animation if this was the last moved clause
                if (item.id && item.id === lastMovedClauseId) {
                    li.classList.add(
                        lastMoveDirection === "up" ? "clause-anim-up" : "clause-anim-down"
                    );
                }

                const arr = section.items;

                // Up button - move clause up in the list
                const upBtn = document.createElement("button");
                upBtn.type = "button";
                upBtn.textContent = "Up";
                upBtn.title = "Move up";
                upBtn.className =
                    "mt-0.5 text-[10px] px-1.5 py-0.5 rounded border border-slate-300 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed";
                upBtn.addEventListener("click", () => {
                    const index = arr.indexOf(item);
                    if (index <= 0) return;
                    const tmp = arr[index - 1];
                    arr[index - 1] = arr[index];
                    arr[index] = tmp;

                    arr.forEach((it, i) => {
                        it.sortOrder = i + 1;
                    });

                    setLastMovedClause(item.id, "up");
                    setClausesDirty(true);
                    renderClausesUI();
                });
                upBtn.disabled = arr.indexOf(item) <= 0;
                li.appendChild(upBtn);

                // Down button - move clause down in the list
                const downBtn = document.createElement("button");
                downBtn.type = "button";
                downBtn.textContent = "Down";
                downBtn.title = "Move down";
                downBtn.className =
                    "mt-0.5 text-[10px] px-1.5 py-0.5 rounded border border-slate-300 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed";
                downBtn.addEventListener("click", () => {
                    const index = arr.indexOf(item);
                    if (index === -1 || index >= arr.length - 1) return;
                    const tmp = arr[index + 1];
                    arr[index + 1] = arr[index];
                    arr[index] = tmp;

                    arr.forEach((it, i) => {
                        it.sortOrder = i + 1;
                    });

                    setLastMovedClause(item.id, "down");
                    setClausesDirty(true);
                    renderClausesUI();
                });
                downBtn.disabled = arr.indexOf(item) === arr.length - 1;
                li.appendChild(downBtn);

                // Enabled checkbox - toggles whether clause is included
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.className = "mt-1";
                cb.checked = item.enabled;
                cb.addEventListener("change", () => {
                    item.enabled = cb.checked;
                    setClausesDirty(true);
                });
                li.appendChild(cb);

                // Editable text - contenteditable div for clause text
                const textDiv = document.createElement("div");
                textDiv.contentEditable = "true";
                textDiv.className =
                    "flex-1 px-2 py-1 border border-slate-200 rounded min-h-[28px] focus:outline-none focus:ring-1 focus:ring-slate-400 bg-white/70";
                if (item.html) {
                    textDiv.innerHTML = item.html;
                } else {
                    textDiv.innerText = item.text || "";
                }
                textDiv.addEventListener("input", () => {
                    item.html = textDiv.innerHTML;
                    item.text = textDiv.innerText;
                    setClausesDirty(true);
                });
                li.appendChild(textDiv);

                // Bold button - applies bold formatting to selected text
                const boldBtn = document.createElement("button");
                boldBtn.type = "button";
                boldBtn.textContent = "B";
                boldBtn.className =
                    "mt-0.5 text-[10px] px-1.5 py-0.5 rounded border border-slate-400 font-semibold";
                boldBtn.addEventListener("mousedown", (e) => {
                    e.preventDefault();
                    textDiv.focus();
                    applyBoldToSelection(textDiv);
                    item.html = textDiv.innerHTML;
                    item.text = textDiv.innerText;
                    setClausesDirty(true);
                });
                li.appendChild(boldBtn);

                // Delete button - removes the clause
                const deleteBtn = document.createElement("button");
                deleteBtn.type = "button";
                deleteBtn.textContent = "×";
                deleteBtn.title = "Remove clause";
                deleteBtn.className =
                    "mt-0.5 text-[10px] px-1.5 py-0.5 rounded border border-red-400 text-red-600 hover:bg-red-50";
                deleteBtn.addEventListener("click", () => {
                    const index = arr.indexOf(item);
                    if (index === -1) return;
                    arr.splice(index, 1);
                    arr.forEach((it, i) => {
                        it.sortOrder = i + 1;
                    });
                    setClausesDirty(true);
                    renderClausesUI();
                });
                li.appendChild(deleteBtn);

                list.appendChild(li);
            });

        box.appendChild(list);

        // Add new clause input and button
        const addWrap = document.createElement("div");
        addWrap.className = "flex gap-1 mt-1";

        const addInput = document.createElement("input");
        addInput.type = "text";
        addInput.placeholder = "Add custom clause and press +";
        addInput.className =
            "flex-1 text-[11px] border border-slate-300 rounded px-2 py-1 bg-white/80";
        addWrap.appendChild(addInput);

        const addBtn = document.createElement("button");
        addBtn.type = "button";
        addBtn.textContent = "+";
        addBtn.className =
            "text-xs px-2 py-1 rounded bg-slate-900 text-white hover:bg-slate-800";
        addBtn.addEventListener("click", () => {
            const text = addInput.value.trim();
            if (!text) return;
            section.items.push({
                id: generateClauseId(),
                text,
                html: "",
                enabled: true,
                sortOrder: section.items.length + 1,
            });
            addInput.value = "";
            setClausesDirty(true);
            renderClausesUI();
        });
        addWrap.appendChild(addBtn);

        box.appendChild(addWrap);
        container.appendChild(box);
    });
}

/**
 * Gets all enabled clauses, sorted by section and sortOrder
 * @returns {Object} Object with clause sections containing enabled clauses only
 */
export function getSelectedClauses() {
    const selected = {};
    Object.entries(clauseSections).forEach(([key, section]) => {
        const sorted = [...section.items].sort(
            (a, b) => (a.sortOrder || 0) - (b.sortOrder || 0)
        );

        selected[key] = sorted
            .filter((item) => item.enabled && (item.text || item.html))
            .map((item) => ({
                text: item.text || stripHtml(item.html || ""),
                html: item.html || "",
                sortOrder: item.sortOrder,
            }));
    });
    return selected;
}
