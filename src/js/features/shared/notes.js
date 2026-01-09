import { fetchNotesFromSheet, saveNotesToSheet } from "../../api/sheets.js";
import { LOCAL_KEYS, getLocalList, setLocalData } from "../../api/localStore.js";

const LEGACY_COLOR_MAP = {
    amber: "#fef3c7",
    emerald: "#d1fae5",
    sky: "#e0f2fe",
    rose: "#ffe4e6",
    violet: "#ede9fe",
    lime: "#ecfccb",
};

const DEFAULT_NOTE_BG = "#f8fafc";
const NOTE_TEXT_DARK = "#0f172a";
const NOTE_TEXT_LIGHT = "#f8fafc";

const notesState = {
    notes: [],
    currentIndex: 0,
    saveTimer: null,
    widgets: [],
    bound: false,
    dirty: false,
};

function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
}

function isHexColor(value) {
    return /^#?[0-9a-fA-F]{6}$/.test(value) || /^#?[0-9a-fA-F]{3}$/.test(value);
}

function normalizeHexColor(value) {
    const raw = value.replace("#", "");
    if (raw.length === 3) {
        return (
            "#" +
            raw
                .split("")
                .map((ch) => ch + ch)
                .join("")
                .toLowerCase()
        );
    }
    return `#${raw.toLowerCase()}`;
}

function hexToRgb(hex) {
    const normalized = normalizeHexColor(hex).replace("#", "");
    return {
        r: parseInt(normalized.slice(0, 2), 16),
        g: parseInt(normalized.slice(2, 4), 16),
        b: parseInt(normalized.slice(4, 6), 16),
    };
}

function rgbToHex(r, g, b) {
    const toHex = (value) => clamp(Math.round(value), 0, 255).toString(16).padStart(2, "0");
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function hslToHex(hue, saturation, lightness) {
    const s = saturation / 100;
    const l = lightness / 100;
    const c = (1 - Math.abs(2 * l - 1)) * s;
    const hPrime = hue / 60;
    const x = c * (1 - Math.abs((hPrime % 2) - 1));
    let r = 0;
    let g = 0;
    let b = 0;

    if (hPrime >= 0 && hPrime < 1) {
        r = c;
        g = x;
    } else if (hPrime >= 1 && hPrime < 2) {
        r = x;
        g = c;
    } else if (hPrime >= 2 && hPrime < 3) {
        g = c;
        b = x;
    } else if (hPrime >= 3 && hPrime < 4) {
        g = x;
        b = c;
    } else if (hPrime >= 4 && hPrime < 5) {
        r = x;
        b = c;
    } else if (hPrime >= 5 && hPrime < 6) {
        r = c;
        b = x;
    }

    const m = l - c / 2;
    return rgbToHex((r + m) * 255, (g + m) * 255, (b + m) * 255);
}

function randomNoteColor() {
    const hue = Math.floor(Math.random() * 360);
    const saturation = 55 + Math.random() * 20;
    const lightness = 80 + Math.random() * 10;
    return hslToHex(hue, saturation, lightness);
}

function normalizeNoteColor(value) {
    const raw = (value || "").toString().trim();
    if (LEGACY_COLOR_MAP[raw]) return LEGACY_COLOR_MAP[raw];
    if (raw && isHexColor(raw)) return normalizeHexColor(raw);
    return randomNoteColor();
}

function getLuminance({ r, g, b }) {
    const toLinear = (channel) => {
        const normalized = channel / 255;
        return normalized <= 0.03928
            ? normalized / 12.92
            : Math.pow((normalized + 0.055) / 1.055, 2.4);
    };
    return 0.2126 * toLinear(r) + 0.7152 * toLinear(g) + 0.0722 * toLinear(b);
}

function contrastRatio(colorA, colorB) {
    const lumA = getLuminance(hexToRgb(colorA));
    const lumB = getLuminance(hexToRgb(colorB));
    const lighter = Math.max(lumA, lumB);
    const darker = Math.min(lumA, lumB);
    return (lighter + 0.05) / (darker + 0.05);
}

function pickTextColor(bgColor) {
    const darkRatio = contrastRatio(bgColor, NOTE_TEXT_DARK);
    const lightRatio = contrastRatio(bgColor, NOTE_TEXT_LIGHT);
    return darkRatio >= lightRatio ? NOTE_TEXT_DARK : NOTE_TEXT_LIGHT;
}

function adjustColor(hex, amount) {
    const { r, g, b } = hexToRgb(hex);
    return rgbToHex(r + amount, g + amount, b + amount);
}

function toRgba(hex, alpha) {
    const { r, g, b } = hexToRgb(hex);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function buildNoteStyles(bgColor) {
    const textColor = pickTextColor(bgColor);
    const textIsDark = textColor === NOTE_TEXT_DARK;
    const borderColor = adjustColor(bgColor, textIsDark ? -28 : 30);
    const placeholderColor = toRgba(textColor, textIsDark ? 0.45 : 0.7);
    const shadowBase = adjustColor(bgColor, textIsDark ? -55 : 45);
    return {
        bg: bgColor,
        text: textColor,
        border: borderColor,
        placeholder: placeholderColor,
        shadow: `inset 0 10px 24px ${toRgba(shadowBase, 0.35)}, inset 0 -6px 10px ${toRgba(shadowBase, 0.22)}`,
        shadowFocus: `inset 0 12px 28px ${toRgba(shadowBase, 0.45)}, inset 0 -8px 14px ${toRgba(shadowBase, 0.3)}`,
    };
}

function createNoteId() {
    return `note-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeNote(note = {}, index = 0) {
    const now = new Date();
    const color = normalizeNoteColor(note.color || note.bg_color || note.bgColor);
    return {
        note_id: note.note_id || note.noteId || note.id || createNoteId(),
        content: (note.content || note.text || note.note || "").toString(),
        color,
        created_at: note.created_at || note.createdAt || now.toISOString().slice(0, 10),
        updated_at: note.updated_at || note.updatedAt || now.toISOString(),
        order: typeof note.order === "number" ? note.order : index,
    };
}

function normalizeNotes(list = []) {
    return list.map((note, idx) => normalizeNote(note, idx));
}

function getCurrentNote() {
    if (!notesState.notes.length) return null;
    const idx = Math.min(Math.max(notesState.currentIndex, 0), notesState.notes.length - 1);
    return notesState.notes[idx] || null;
}

function filterPersistableNotes(list = []) {
    const notes = Array.isArray(list) ? list : [];
    return notes.filter((note) => (note?.content || "").trim());
}

function updateLocalNotes(notes) {
    const filtered = filterPersistableNotes(notes);
    setLocalData(LOCAL_KEYS.notes, filtered);
    document.dispatchEvent(new CustomEvent("notes:updated", { detail: filtered }));
}

function scheduleSyncNotes() {
    if (notesState.saveTimer) clearTimeout(notesState.saveTimer);
    notesState.saveTimer = setTimeout(() => {
        saveNotesToSheet(filterPersistableNotes(notesState.notes));
    }, 700);
}

function formatCreatedDate(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleDateString();
}

function formatUpdatedDate(value) {
    if (!value) return "";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleString();
}

function applyThemeToWidget(widget, note) {
    if (!widget) return;
    const input = widget.querySelector("[data-note-input]");
    if (!input) return;
    const bgColor = note ? normalizeNoteColor(note.color) : DEFAULT_NOTE_BG;
    if (note && note.color !== bgColor) {
        note.color = bgColor;
    }
    const styles = buildNoteStyles(bgColor);
    input.style.setProperty("--note-bg", styles.bg);
    input.style.setProperty("--note-text", styles.text);
    input.style.setProperty("--note-border", styles.border);
    input.style.setProperty("--note-placeholder", styles.placeholder);
    input.style.setProperty("--note-shadow", styles.shadow);
    input.style.setProperty("--note-shadow-focus", styles.shadowFocus);
    input.style.caretColor = styles.text;
}

function renderWidget(widget) {
    const note = getCurrentNote();
    const total = notesState.notes.length;
    const index = total ? notesState.currentIndex + 1 : 0;
    const hasContent = !!note && note.content.trim().length > 0;
    const canDiscardEmptyDraft =
        !!note && !hasContent && total > 1 && notesState.currentIndex === total - 1;
    const position = widget.querySelector("[data-note-position]");
    const created = widget.querySelector("[data-note-created]");
    const updated = widget.querySelector("[data-note-updated]");
    const empty = widget.querySelector("[data-note-empty]");
    const input = widget.querySelector("[data-note-input]");
    const prevBtn = widget.querySelector('[data-note-action="prev"]');
    const nextBtn = widget.querySelector('[data-note-action="next"]');
    const addBtn = widget.querySelector('[data-note-action="add"]');
    const deleteBtn = widget.querySelector('[data-note-action="delete"]');
    const copyBtn = widget.querySelector('[data-note-action="copy"]');
    const cancelBtn = widget.querySelector('[data-note-action="cancel"]');

    if (position) position.textContent = total ? `Note ${index} of ${total}` : "Note 0 of 0";
    if (prevBtn) {
        prevBtn.disabled = !hasContent || total <= 1 || notesState.currentIndex <= 0 || canDiscardEmptyDraft;
        prevBtn.classList.toggle("hidden", canDiscardEmptyDraft);
    }
    if (nextBtn) {
        const hasNext = total > 1 && notesState.currentIndex < total - 1;
        nextBtn.disabled = !hasNext || canDiscardEmptyDraft;
        nextBtn.classList.toggle("hidden", canDiscardEmptyDraft);
    }
    if (addBtn) addBtn.disabled = !hasContent;
    if (deleteBtn) {
        deleteBtn.disabled = !hasContent;
        deleteBtn.classList.toggle("invisible", !hasContent);
        deleteBtn.classList.toggle("pointer-events-none", !hasContent);
    }
    if (copyBtn) copyBtn.disabled = !hasContent;
    if (cancelBtn) {
        cancelBtn.disabled = !canDiscardEmptyDraft;
        cancelBtn.classList.toggle("hidden", !canDiscardEmptyDraft);
    }

    if (!note) {
        if (empty) empty.classList.remove("hidden");
        if (input) {
            input.value = "";
            input.disabled = false;
        }
        if (created) created.textContent = "-";
        if (updated) updated.textContent = "-";
        applyThemeToWidget(widget, null);
        return;
    }

    if (empty) empty.classList.add("hidden");
    if (input) {
        input.disabled = false;
        if (document.activeElement !== input) {
            input.value = note.content;
        }
    }
    if (created) created.textContent = formatCreatedDate(note.created_at);
    if (updated) updated.textContent = formatUpdatedDate(note.updated_at);
    applyThemeToWidget(widget, note);
}

function renderAllWidgets() {
    notesState.widgets.forEach((widget) => renderWidget(widget));
}

function setNotes(list) {
    notesState.notes = normalizeNotes(list || []);
    if (notesState.currentIndex >= notesState.notes.length) {
        notesState.currentIndex = Math.max(0, notesState.notes.length - 1);
    }
    notesState.dirty = false;
    renderAllWidgets();
}

function addNote() {
    const current = getCurrentNote();
    if (!current || !current.content.trim()) return;
    const next = normalizeNotes(notesState.notes);
    const note = normalizeNote(
        {
            content: "",
            color: randomNoteColor(),
        },
        next.length
    );
    next.push(note);
    notesState.currentIndex = next.length - 1;
    setNotes(next);
}

function deleteNote() {
    if (!notesState.notes.length) return;
    const next = notesState.notes.slice();
    next.splice(notesState.currentIndex, 1);
    if (notesState.currentIndex >= next.length) {
        notesState.currentIndex = Math.max(0, next.length - 1);
    }
    setNotes(next);
    updateLocalNotes(next);
    scheduleSyncNotes();
}

function updateCurrentNoteDraft(value) {
    const note = getCurrentNote();
    if (!note) {
        const now = new Date().toISOString();
        const createdAt = now.slice(0, 10);
        const newNote = normalizeNote(
            {
            content: value,
            color: randomNoteColor(),
            created_at: createdAt,
            updated_at: now,
            },
            notesState.notes.length
        );
        notesState.notes = notesState.notes.concat(newNote);
        notesState.currentIndex = notesState.notes.length - 1;
        notesState.dirty = true;
        renderAllWidgets();
        return;
    }
    if (note.content === value) return;
    note.content = value;
    notesState.dirty = true;
    renderAllWidgets();
}

function commitCurrentNote() {
    if (!notesState.dirty) return;
    const note = getCurrentNote();
    if (!note) return;
    if (!note.content.trim()) {
        notesState.dirty = false;
        renderAllWidgets();
        return;
    }
    note.updated_at = new Date().toISOString();
    notesState.dirty = false;
    updateLocalNotes(notesState.notes);
    scheduleSyncNotes();
}

function moveNote(direction) {
    if (!notesState.notes.length) return;
    const nextIndex = notesState.currentIndex + direction;
    if (nextIndex < 0 || nextIndex >= notesState.notes.length) return;
    notesState.currentIndex = nextIndex;
    renderAllWidgets();
}

function cancelDraftNote() {
    const note = getCurrentNote();
    if (!note || note.content.trim()) return;
    const next = notesState.notes.slice();
    next.splice(notesState.currentIndex, 1);
    if (notesState.currentIndex >= next.length) {
        notesState.currentIndex = Math.max(0, next.length - 1);
    }
    setNotes(next);
    updateLocalNotes(next);
}

async function copyNoteContent(widget) {
    const input = widget.querySelector("[data-note-input]");
    const text = input ? input.value : (getCurrentNote() && getCurrentNote().content) || "";
    if (!text) return;
    try {
        await navigator.clipboard.writeText(text);
    } catch (err) {
        if (!input) return;
        const selectionStart = input.selectionStart;
        const selectionEnd = input.selectionEnd;
        input.focus();
        input.select();
        document.execCommand("copy");
        if (typeof selectionStart === "number" && typeof selectionEnd === "number") {
            input.setSelectionRange(selectionStart, selectionEnd);
        }
    }
}

function bindWidget(widget) {
    const input = widget.querySelector("[data-note-input]");
    const prevBtn = widget.querySelector('[data-note-action="prev"]');
    const nextBtn = widget.querySelector('[data-note-action="next"]');
    const addBtn = widget.querySelector('[data-note-action="add"]');
    const deleteBtn = widget.querySelector('[data-note-action="delete"]');
    const copyBtn = widget.querySelector('[data-note-action="copy"]');
    const cancelBtn = widget.querySelector('[data-note-action="cancel"]');

    if (input && !input.dataset.bound) {
        input.dataset.bound = "true";
        input.addEventListener("input", (event) => {
            updateCurrentNoteDraft(event.target.value);
        });
        input.addEventListener("blur", () => {
            commitCurrentNote();
        });
    }

    if (prevBtn && !prevBtn.dataset.bound) {
        prevBtn.dataset.bound = "true";
        prevBtn.addEventListener("click", () => moveNote(-1));
    }

    if (nextBtn && !nextBtn.dataset.bound) {
        nextBtn.dataset.bound = "true";
        nextBtn.addEventListener("click", () => moveNote(1));
    }

    if (addBtn && !addBtn.dataset.bound) {
        addBtn.dataset.bound = "true";
        addBtn.addEventListener("click", () => addNote());
    }

    if (deleteBtn && !deleteBtn.dataset.bound) {
        deleteBtn.dataset.bound = "true";
        deleteBtn.addEventListener("click", () => deleteNote());
    }

    if (copyBtn && !copyBtn.dataset.bound) {
        copyBtn.dataset.bound = "true";
        copyBtn.addEventListener("click", () => copyNoteContent(widget));
    }

    if (cancelBtn && !cancelBtn.dataset.bound) {
        cancelBtn.dataset.bound = "true";
        cancelBtn.addEventListener("click", () => cancelDraftNote());
    }
}

async function loadInitialNotes() {
    const localNotes = await getLocalList(LOCAL_KEYS.notes, []);
    if (localNotes.length) {
        setNotes(localNotes);
    }

    const remote = await fetchNotesFromSheet(false);
    if (Array.isArray(remote?.notes)) {
        setNotes(remote.notes);
    }
}

export function initNotesFeature() {
    if (notesState.bound) return;
    notesState.bound = true;
    notesState.widgets = Array.from(document.querySelectorAll("[data-notes-widget]"));
    notesState.widgets.forEach((widget) => bindWidget(widget));

    document.addEventListener("notes:updated", (event) => {
        if (Array.isArray(event.detail)) {
            setNotes(event.detail);
        }
    });

    loadInitialNotes();
}

export async function refreshNotes(force = false) {
    const res = await fetchNotesFromSheet(force);
    if (Array.isArray(res?.notes)) {
        setNotes(res.notes);
    }
}
