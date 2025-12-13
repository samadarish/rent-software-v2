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
        locationLabel.textContent = `Saved to: Downloads/${safeFileName}`;
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
        .then((arrayBuffer) => {
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

                saveAs(out, fileName);
                syncDocxExportModal(fileName, lastDocxDownloadUrl);
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
