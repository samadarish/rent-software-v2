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

let tauriApiPromise = null;

function normalizeFileUrl(pathOrUrl) {
    if (!pathOrUrl) return "";

    // If it is already a URL, just return it
    try {
        const candidate = new URL(pathOrUrl, "file://");
        if (candidate.protocol !== "file:") return pathOrUrl;
    } catch {
        // Not a parsable URL; treat as a path
    }

    const rawPath = pathOrUrl.replace(/\\/g, "/");
    if (rawPath.startsWith("file://")) return rawPath;

    // Prepend an extra slash for Windows drive letters (C:/)
    const needsDriveSlash = /^[A-Za-z]:/.test(rawPath);
    const prefix = needsDriveSlash ? "file:///" : "file://";
    return `${prefix}${rawPath}`;
}

async function loadTauriApis() {
    if (tauriApiPromise) return tauriApiPromise;

    tauriApiPromise = (async () => {
        // Only attempt to use Tauri when running inside its WebView context
        if (typeof window === "undefined") return null;

        const tauriGlobal = window.__TAURI__ || window.__TAURI_INTERNALS__;
        if (!tauriGlobal) return null;

        // Prefer globals if they are already injected to avoid module resolution issues
        const writeBinaryFile =
            tauriGlobal.fs?.writeBinaryFile || tauriGlobal.core?.fs?.writeBinaryFile;
        const BaseDirectory = tauriGlobal.fs?.BaseDirectory || tauriGlobal.core?.fs?.BaseDirectory;
        const downloadDir = tauriGlobal.path?.downloadDir;
        const join = tauriGlobal.path?.join;
        const convertFileSrc = tauriGlobal.tauri?.convertFileSrc;
        const open =
            tauriGlobal.opener?.open ||
            tauriGlobal.plugin?.opener?.open ||
            tauriGlobal.plugins?.opener?.open ||
            tauriGlobal.shell?.open;

        const hasSaveApis = Boolean(writeBinaryFile && (BaseDirectory || (downloadDir && join)));
        const hasAnyApis = hasSaveApis || Boolean(open);

        if (!hasAnyApis) return null;

        if (!hasSaveApis) {
            console.warn(
                "Tauri APIs detected but required methods are missing for saving; opener support only"
            );
        }

        return {
            writeBinaryFile: hasSaveApis ? writeBinaryFile : null,
            BaseDirectory: hasSaveApis ? BaseDirectory : null,
            downloadDir: hasSaveApis ? downloadDir : null,
            join: hasSaveApis ? join : null,
            convertFileSrc,
            open,
        };
    })();

    return tauriApiPromise;
}

async function saveDocxToDownloads(blob, fileName) {
    const tauriApis = await loadTauriApis();
    if (!tauriApis) return null;

    const { writeBinaryFile, BaseDirectory, downloadDir, join } = tauriApis;
    if (!writeBinaryFile) return null;

    const buffer = new Uint8Array(await blob.arrayBuffer());

    if (BaseDirectory?.Download) {
        const args =
            writeBinaryFile.length > 1
                ? [fileName, buffer, { dir: BaseDirectory.Download }]
                : [{ path: fileName, contents: buffer, dir: BaseDirectory.Download }];
        await writeBinaryFile(...args);

        if (downloadDir && join) {
            try {
                const downloadsDir = await downloadDir();
                return await join(downloadsDir, fileName);
            } catch {
                // If we cannot resolve the final path, still report success without path
                return fileName;
            }
        }

        return fileName;
    }

    if (downloadDir && join) {
        const downloadsDir = await downloadDir();
        const targetPath = await join(downloadsDir, fileName);
        const args =
            writeBinaryFile.length > 1
                ? [targetPath, buffer]
                : [{ path: targetPath, contents: buffer }];
        await writeBinaryFile(...args);
        return targetPath;
    }

    return null;
}

function revokeLastDocxUrl() {
    if (lastDocxDownloadUrl) {
        URL.revokeObjectURL(lastDocxDownloadUrl);
        lastDocxDownloadUrl = "";
    }
}

function syncDocxExportModal(fileName, objectUrl, filePath) {
    const modal = document.getElementById("docxExportModal");
    const fileNameLabel = document.getElementById("docxExportFileName");
    const locationLabel = document.getElementById("docxExportLocation");
    const openLink = document.getElementById("docxExportOpen");

    const safeFileName = fileName || "Agreement.docx";
    lastDocxFileName = safeFileName;
    lastDocxFilePath = filePath || "";

    if (fileNameLabel) fileNameLabel.textContent = safeFileName;

    if (locationLabel) {
        locationLabel.textContent = lastDocxFilePath
            ? `Saved to: ${lastDocxFilePath}`
            : `Downloaded: ${safeFileName}`;
    }

    if (openLink) {
        openLink.href = "#";
        openLink.dataset.filePath = lastDocxFilePath;
        openLink.dataset.objectUrl = objectUrl || "";
        openLink.dataset.filename = safeFileName;
        openLink.dataset.openReady = lastDocxFilePath || objectUrl ? "true" : "false";
        openLink.rel = "noopener";
        const canAttemptOpen = Boolean(lastDocxFilePath || objectUrl);
        openLink.classList.toggle("pointer-events-none", !canAttemptOpen);
        openLink.classList.toggle("opacity-50", !canAttemptOpen);
    }

    if (modal) showModal(modal);
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
        const openWithTauri = async (targetPath) => {
            const tauriApis = await loadTauriApis();
            if (!tauriApis?.open || !targetPath) return false;

            let tauriUrl = targetPath;
            if (tauriApis.convertFileSrc) {
                try {
                    tauriUrl = tauriApis.convertFileSrc(targetPath);
                } catch (error) {
                    console.warn("convertFileSrc failed; using raw path", error);
                }
            }

            const normalizedPath = normalizeFileUrl(tauriUrl);

            try {
                const openArgs = tauriApis.open.length > 1 ? [normalizedPath, { newInstance: true }] : [normalizedPath];
                await tauriApis.open(...openArgs);
                return true;
            } catch (error) {
                console.warn("Tauri open failed, retrying with raw path", error);
                try {
                    await tauriApis.open(targetPath);
                    return true;
                } catch (innerError) {
                    console.error("DOCX open error:", innerError);
                    return false;
                }
            }
        };

        openLink.addEventListener("click", async (e) => {
            e.preventDefault();

            if (!lastDocxDownloadUrl && !lastDocxFilePath) {
                showToast("Export a DOCX first, then try opening it again.", "warning");
                return;
            }

            if (lastDocxFilePath) {
                const opened = await openWithTauri(lastDocxFilePath);
                if (opened) return;

                showToast(
                    "Could not open the DOCX directly. Please open it from your Downloads folder.",
                    "error"
                );
                return;
            }

            if (lastDocxDownloadUrl) {
                const opened = await openWithTauri(lastDocxDownloadUrl);
                if (opened) return;

                window.open(lastDocxDownloadUrl, "_blank", "noopener,noreferrer");
                return;
            }

            if (typeof navigator !== "undefined" && navigator.msSaveOrOpenBlob && lastDocxBlob) {
                navigator.msSaveOrOpenBlob(lastDocxBlob, lastDocxFileName || "Agreement.docx");
                return;
            }

            showToast(
                "The file has already been downloaded. Please open it manually from your Downloads folder.",
                "info"
            );
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
export function exportDocxFromTemplate() {
    wireDocxExportModal();

    const data = collectFormDataForTemplate();

    // Validate key fields before exporting
    if (!data.agreement_date || !data.Tenant_Full_Name) {
        if (!confirm("Some key fields look empty. Export DOCX anyway?")) {
            return;
        }
    }

    const templatePath = "tenant_template.docx";

    fetch(templatePath)
        .then((res) => {
            if (!res.ok) {
                throw new Error("Could not load template: HTTP " + res.status);
            }
            return res.arrayBuffer();
        })
        .then(async (arrayBuffer) => {
            try {
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
                    mimeType:
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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

                let savedFilePath = "";
                try {
                    savedFilePath = (await saveDocxToDownloads(out, fileName)) || "";
                } catch (saveError) {
                    console.warn("Could not persist DOCX via Tauri APIs", saveError);
                }

                saveAs(out, fileName);
                syncDocxExportModal(fileName, lastDocxDownloadUrl, savedFilePath);
                showToast(`DOCX downloaded as "${fileName}"`, "success");
            } catch (e) {
                console.error("DOCX render error:", e);
                alert("Error while filling template. Check console for details.");
            }
        })
        .catch((err) => {
            console.error("Error loading template via fetch:", err);
            alert(
                "Could not load DOCX template. Make sure 'tenant_template.docx' is in the public/ folder."
            );
        });
}
