/**
 * DOCX Export Feature Module
 * 
 * Handles exporting agreements as DOCX files using a template.
 */

import { downloadDir, join } from "@tauri-apps/api/path";
import { collectFormDataForTemplate } from "../tenants/form.js";
import { hideModal, showModal, showToast } from "../../utils/ui.js";

let lastDocxDownloadUrl = "";
let lastDocxFileName = "";
let lastDocxBlob = null;
let lastDocxSavedPath = "";
let docxModalWired = false;

function revokeLastDocxUrl() {
    if (lastDocxDownloadUrl) {
        URL.revokeObjectURL(lastDocxDownloadUrl);
        lastDocxDownloadUrl = "";
    }
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
        const locationText = lastDocxSavedPath
            ? `Saved to: ${lastDocxSavedPath}`
            : `Saved to: Downloads/${safeFileName}`;
        locationLabel.textContent = locationText;
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

        if (lastDocxSavedPath) {
            openLink.removeAttribute("download");
            openLink.dataset.savedPath = lastDocxSavedPath;
        } else {
            openLink.removeAttribute("data-saved-path");
        }

        const hasOpenTarget = Boolean(objectUrl || lastDocxSavedPath);
        openLink.classList.toggle("pointer-events-none", !hasOpenTarget);
        openLink.classList.toggle("opacity-50", !hasOpenTarget);
    }

    if (modal) showModal(modal);
}

async function openDocxWithSystemApp(targetPath) {
    try {
        const opener = window.__TAURI__?.opener;
        if (opener?.openPath && targetPath) {
            await opener.openPath(targetPath);
            return true;
        }
    } catch (err) {
        console.error("Tauri opener failed to open file", err);
    }

    return false;
}

async function saveDocxToDownloads(blob, fileName) {
    const fs = window.__TAURI__?.fs;
    if (!fs?.writeFile) return null;

    try {
        let downloadsDir = null;
        try {
            downloadsDir = await downloadDir();
        } catch (err) {
            console.error("Unable to access tauri path downloadDir", err);
        }

        if (!downloadsDir && window.__TAURI__?.path?.downloadDir) {
            downloadsDir = await window.__TAURI__.path.downloadDir();
        }
        if (!downloadsDir) return null;

        let targetPath = null;
        if (typeof join === "function") {
            targetPath = await join(downloadsDir, fileName);
        } else if (window.__TAURI__?.path?.join) {
            targetPath = await window.__TAURI__.path.join(downloadsDir, fileName);
        }
        if (!targetPath) return null;
        const contents = new Uint8Array(await blob.arrayBuffer());

        await fs.writeFile({
            path: targetPath,
            contents,
        });

        lastDocxSavedPath = targetPath;
        return targetPath;
    } catch (err) {
        console.error("Unable to save DOCX to downloads directory", err);
        return null;
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
        openLink.addEventListener("click", (e) => {
            if (lastDocxSavedPath) {
                e.preventDefault();
                openDocxWithSystemApp(lastDocxSavedPath).then((opened) => {
                    if (!opened) {
                        if (lastDocxDownloadUrl) {
                            const fallbackAnchor = document.createElement("a");
                            fallbackAnchor.href = lastDocxDownloadUrl;
                            fallbackAnchor.target = "_blank";
                            fallbackAnchor.rel = "noopener";
                            fallbackAnchor.download = lastDocxFileName || "Agreement.docx";
                            fallbackAnchor.click();
                        }

                        showToast(
                            "Couldn't open the saved DOCX. Please open it from your file manager.",
                            "error"
                        );
                    }
                });
                return;
            }

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

            // Let the browser handle the navigation so it is treated as a
            // user-initiated action and isn't blocked by pop-up blockers.
            openLink.href = lastDocxDownloadUrl;
            openLink.target = "_blank";
            openLink.rel = "noopener";
            openLink.download = lastDocxFileName || "Agreement.docx";
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
    lastDocxSavedPath = "";

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

        const savedPath = await saveDocxToDownloads(out, fileName);

        revokeLastDocxUrl();
        lastDocxBlob = out;
        lastDocxDownloadUrl = URL.createObjectURL(out);

        if (!savedPath) {
            saveAs(out, fileName);
        }

        syncDocxExportModal(fileName, lastDocxDownloadUrl);

        const successMessage = savedPath
            ? `DOCX saved to "${savedPath}"`
            : `DOCX downloaded as "${fileName}"`;
        showToast(successMessage, "success");
    } catch (err) {
        console.error("Error exporting DOCX", err);
        alert(
            "Could not export DOCX. Check the console for details and ensure 'tenant_template.docx' is available."
        );
    }
}
