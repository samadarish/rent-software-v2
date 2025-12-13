/**
 * DOCX Export Feature Module
 * 
 * Handles exporting agreements as DOCX files using a template.
 */

import { collectFormDataForTemplate } from "../tenants/form.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

let lastDocxDownloadUrl = "";
let lastDocxFileName = "";
let lastDocxBlob = null;
let lastDocxFilePath = "";
let docxModalWired = false;
let fsHelperPromise = null;

function revokeLastDocxUrl() {
    if (lastDocxDownloadUrl) {
        URL.revokeObjectURL(lastDocxDownloadUrl);
        lastDocxDownloadUrl = "";
    }

    lastDocxFilePath = "";
}

function syncDocxExportModal(fileName, objectUrl) {
    const modal = document.getElementById("docxExportModal");
    const fileNameLabel = document.getElementById("docxExportFileName");
    const locationLabel = document.getElementById("docxExportLocation");
    const openLink = document.getElementById("docxExportOpen");

    const safeFileName = fileName || "Agreement.docx";
    lastDocxFileName = safeFileName;

    if (fileNameLabel) fileNameLabel.textContent = safeFileName;

    if (locationLabel) {
        const location = lastDocxFilePath || `Downloads/${safeFileName}`;
        locationLabel.textContent = `Saved to: ${location}`;
    }

    if (openLink) {
        openLink.href = objectUrl || "#";
        if (objectUrl) {
            openLink.removeAttribute("download");
            openLink.setAttribute("data-filename", safeFileName);
            openLink.target = "_blank";
            openLink.rel = "noopener";
            openLink.dataset.openReady = "true";
        }
        openLink.classList.toggle("pointer-events-none", !objectUrl);
        openLink.classList.toggle("opacity-50", !objectUrl);
    }

    if (modal) showModal(modal);
}

async function loadTauriFsHelpers() {
    const globalFs = window.__TAURI__?.fs;
    const globalPath = window.__TAURI__?.path;

    const writeBinaryFile = globalFs?.writeBinaryFile;
    const downloadDir = globalPath?.downloadDir;
    const join = globalPath?.join;

    if (writeBinaryFile && downloadDir && join) {
        return { writeBinaryFile, downloadDir, join };
    }

    // If Tauri isn't exposing the fs/path globals (likely because permissions
    // are denied or we're running in a pure browser context), bail early so we
    // don't attempt asset-protocol imports that are blocked by CORS.
    console.warn(
        "Tauri fs/path helpers not available on window.__TAURI__; using browser download instead."
    );

    return null;
}

function getTauriFsHelpers() {
    if (!fsHelperPromise) {
        fsHelperPromise = loadTauriFsHelpers();
    }
    return fsHelperPromise;
}

async function saveDocxToDownloads(blob, fileName) {
    if (!window.__TAURI__) return "";

    try {
        const helpers = await getTauriFsHelpers();
        if (!helpers) {
            console.warn(
                "Tauri filesystem helpers unavailable; browser download fallback will be used."
            );
            return "";
        }

        const { writeBinaryFile, downloadDir, join } = helpers;

        const downloads = await downloadDir();
        if (!downloads) return "";

        const targetPath = await join(downloads, fileName);
        const bytes = new Uint8Array(await blob.arrayBuffer());
        await writeBinaryFile({ path: targetPath, contents: bytes });
        return targetPath;
    } catch (err) {
        console.error("Unable to save DOCX to downloads via Tauri FS", err);
        return "";
    }
}

async function tryOpenDocxWithTauri(event) {
    if (!window.__TAURI__ || !lastDocxFilePath) return false;

    try {
        const { opener } = window.__TAURI__;
        const openPath = opener?.openPath || opener?.open;

        if (typeof openPath !== "function") {
            throw new Error("Tauri opener plugin is unavailable");
        }

        await openPath(lastDocxFilePath);
        if (event) event.preventDefault();
        return true;
    } catch (err) {
        console.error("Unable to open DOCX with Tauri opener", err);
        showToast("Couldn't open the saved DOCX automatically. Opening download instead.", "warning");
        return false;
    }
}

function wireDocxExportModal() {
    if (docxModalWired) return;
    docxModalWired = true;

    const closeBtn = document.getElementById("docxExportClose");
    const modal = document.getElementById("docxExportModal");
    const openLink = document.getElementById("docxExportOpen");

    if (closeBtn) {
        closeBtn.addEventListener("click", () => {
            if (modal) hideModal(modal);
        });
    }

    if (modal) {
        modal.addEventListener("click", (e) => {
            if (e.target === modal) {
                hideModal(modal);
            }
        });
    }

    if (openLink) {
        openLink.addEventListener("click", async (e) => {
            if (!lastDocxDownloadUrl) {
                e.preventDefault();
                showToast("Export a DOCX first, then try opening it again.", "warning");
                return;
            }

            if (typeof navigator !== "undefined" && navigator.msSaveOrOpenBlob && lastDocxBlob) {
                e.preventDefault();
                navigator.msSaveOrOpenBlob(lastDocxBlob, lastDocxFileName || "Agreement.docx");
                return;
            }

            const openedWithTauri = await tryOpenDocxWithTauri(e);
            if (openedWithTauri) return;

            // Let the browser handle the navigation so it is treated as a
            // user-initiated action and isn't blocked by pop-up blockers.
            openLink.href = lastDocxDownloadUrl;
            openLink.target = "_blank";
            openLink.rel = "noopener";
            openLink.removeAttribute("download");
        });
    }
}

/**
 * Applies markdown-style bold formatting (**text**) to DOCX XML
 * Converts **text** markers to proper <w:b/> tags in the DOCX XML
 * @param {string} xml - The DOCX XML content
 * @returns {string} Modified XML with bold formatting applied
 */
export function applyMarkdownBoldToDocxXml(xml) {
    const regex =
        /<w:r\b[^>]*>(\s*<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<w:t([^>]*)>([\s\S]*?)<\/w:t>\s*<\/w:r>/g;

    return xml.replace(regex, (match, rPr, tAttrs, text) => {
        rPr = rPr || "";
        tAttrs = tAttrs || "";
        if (!text || !text.includes("**")) return match;

        const parts = text.split("**");
        if (parts.length < 3 || parts.length % 2 === 0) return match;

        let result = "";

        parts.forEach((part, index) => {
            if (!part) return;
            const isBold = index % 2 === 1;

            if (isBold) {
                let boldPr = rPr;
                if (boldPr) {
                    if (!boldPr.includes("<w:b")) {
                        boldPr = boldPr.replace("</w:rPr>", "<w:b/></w:rPr>");
                    }
                } else {
                    boldPr = "<w:rPr><w:b/></w:rPr>";
                }
                result += `<w:r>${boldPr}<w:t${tAttrs}>${part}</w:t></w:r>`;
            } else {
                result += `<w:r>${rPr}<w:t${tAttrs}>${part}</w:t></w:r>`;
            }
        });

        return result;
    });
}

/**
 * Exports the agreement form data as a DOCX file
 * Uses the tenant_template.docx file and fills it with form data
 */
export async function exportDocxFromTemplate() {
    wireDocxExportModal();

    const data = collectFormDataForTemplate();

    // Validate key fields before exporting
    if (!data.agreement_date || !data.Tenant_Full_Name) {
        if (!confirm("Some key fields look empty. Export DOCX anyway?")) {
            return;
        }
    }

    const templatePath = "tenant_template.docx";

    try {
        const res = await fetch(templatePath);
        if (!res.ok) {
            throw new Error("Could not load template: HTTP " + res.status);
        }

        const arrayBuffer = await res.arrayBuffer();
        const uint8 = new Uint8Array(arrayBuffer);
        const zip = new PizZip(uint8);
        const doc = new docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        doc.setData(data);
        doc.render();

        const outZip = doc.getZip();
        let xml = outZip.file("word/document.xml").asText();
        xml = applyMarkdownBoldToDocxXml(xml);
        outZip.file("word/document.xml", xml);

        const out = outZip.generate({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });

        const sanitize = (value, fallback) => {
            const safeValue = (value || fallback || "")
                .trim()
                .replace(/\s+/g, "_")
                .replace(/[^\w.-]/g, "");
            return safeValue || fallback || "";
        };

        const tenantName = sanitize(data.Tenant_Full_Name, "Tenant");
        const grnNumber = sanitize(data["GRN number"], "GRN");
        const aadhaar = sanitize(data.tenant_Aadhar, "Aadhaar");
        const fileName = `${tenantName}_${grnNumber}_${aadhaar}_Agreement.docx`;

        revokeLastDocxUrl();
        lastDocxBlob = out;
        lastDocxDownloadUrl = URL.createObjectURL(out);

        const savedPath = await saveDocxToDownloads(out, fileName);
        lastDocxFilePath = savedPath;

        if (!savedPath) {
            saveAs(out, fileName);
        }

        syncDocxExportModal(fileName, lastDocxDownloadUrl);
        const successMessage =
            savedPath && window.__TAURI__
                ? `DOCX saved to ${savedPath}`
                : `DOCX downloaded as "${fileName}"`;
        showToast(successMessage, "success");
    } catch (err) {
        console.error("Error exporting DOCX:", err);
        alert("Could not export DOCX. Check console for details.");
    }
}
