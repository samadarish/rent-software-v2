/**
 * Flow/Navigation Feature Module
 * 
 * Handles switching between different application modes:
 * - Agreement mode (create agreement and save tenant)
 * - Create Tenant New mode (active tenant)
 * - Add Past Tenant mode (past tenant)
 */

import { currentFlow, setCurrentFlow } from "../../state.js";
import { isDraftDirty, loadDraftForFlow, openUnsavedDraftModal, syncDraftUiForFlow } from "../shared/drafts.js";
import { loadTenantDirectory } from "../tenants/tenants.js";
import { applyLandlordDefaultsToForm } from "../../api/config.js";
import { refreshPaymentsIfNeeded } from "../billing/payments.js";
import { smoothToggle } from "../../utils/ui.js";

const DRAFT_GUARDED_FLOWS = new Set(["agreement", "createTenantNew", "addPastTenant"]);

/**
 * Switches the application to a different flow/mode
 * Updates UI elements, button visibility, and form sections based on the mode
 * @param {"dashboard" | "agreement" | "createTenantNew" | "addPastTenant" | "viewTenants" | "generateBill" | "payments"} mode - The flow mode to switch to
 */
export function switchFlow(mode, options = { bypassGuard: false }) {
    if (
        currentFlow !== mode &&
        DRAFT_GUARDED_FLOWS.has(currentFlow) &&
        isDraftDirty(currentFlow) &&
        !options.bypassGuard
    ) {
        openUnsavedDraftModal(() => switchFlow(mode, { bypassGuard: true }));
        return;
    }

    setCurrentFlow(mode);
    document.dispatchEvent(new CustomEvent("flow:changed", { detail: { mode } }));

    // Update layout class so equal widths only apply on the agreement form
    const appLayout = document.getElementById("appLayout");
    if (appLayout) {
        appLayout.classList.remove("layout-agreement", "layout-directory");
        if (mode === "agreement") {
            appLayout.classList.add("layout-agreement");
        } else {
            appLayout.classList.add("layout-directory");
        }
    }

    // Show main sections
    const dashboardSection = document.getElementById("dashboardSection");
    const formSection = document.getElementById("formSection");
    const formRightSection = document.getElementById("formRightSection");
    const tenantListSection = document.getElementById("tenantListSection");
    const generateBillSection = document.getElementById("generateBillSection");
    const paymentsSection = document.getElementById("paymentsSection");

    const isDashboard = mode === "dashboard";
    const isGenerateBill = mode === "generateBill";
    const isPayments = mode === "payments";
    const isViewTenants = mode === "viewTenants";

    smoothToggle(dashboardSection, isDashboard);
    smoothToggle(formSection, !(isDashboard || isGenerateBill || isPayments || isViewTenants));
    smoothToggle(formRightSection, true);
    smoothToggle(tenantListSection, isViewTenants);
    smoothToggle(generateBillSection, isGenerateBill);
    smoothToggle(paymentsSection, isPayments);

    // Update navigation button active states
    const dashboardBtn = document.getElementById("navDashboardBtn");
    const createTenantBtn = document.getElementById("navCreateTenantBtn");
    const createAgreementBtn = document.getElementById("navCreateAgreementBtn");
    const viewTenantsBtn = document.getElementById("navViewTenantsBtn");
    const generateBillBtn = document.getElementById("navGenerateBillBtn");
    const paymentsBtn = document.getElementById("navPaymentsBtn");

    const activeClasses = ["bg-slate-900", "text-white", "shadow-sm"];
    const inactiveTextClasses = ["text-slate-700"];
    const hoverClass = "hover:bg-slate-100";

    [dashboardBtn, createTenantBtn, createAgreementBtn, viewTenantsBtn, generateBillBtn, paymentsBtn]
        .filter(Boolean)
        .forEach((btn) => {
            btn.classList.remove(...activeClasses);
            btn.classList.add(hoverClass, ...inactiveTextClasses);
        });

    if (mode === "dashboard" && dashboardBtn) {
        dashboardBtn.classList.remove(hoverClass, ...inactiveTextClasses);
        dashboardBtn.classList.add(...activeClasses);
    }

    if (mode === "agreement" && createAgreementBtn) {
        createAgreementBtn.classList.remove(hoverClass, ...inactiveTextClasses);
        createAgreementBtn.classList.add(...activeClasses);
    }

    if ((mode === "createTenantNew" || mode === "addPastTenant") && createTenantBtn) {
        createTenantBtn.classList.remove(hoverClass, ...inactiveTextClasses);
        createTenantBtn.classList.add(...activeClasses);
    }

    if (mode === "viewTenants" && viewTenantsBtn) {
        viewTenantsBtn.classList.remove(hoverClass, ...inactiveTextClasses);
        viewTenantsBtn.classList.add(...activeClasses);
    }

    if (mode === "generateBill" && generateBillBtn) {
        generateBillBtn.classList.remove(hoverClass, ...inactiveTextClasses);
        generateBillBtn.classList.add(...activeClasses);
    }

    if (mode === "payments" && paymentsBtn) {
        paymentsBtn.classList.remove(hoverClass, ...inactiveTextClasses);
        paymentsBtn.classList.add(...activeClasses);
    }

    // Update form title
    const titleEl = document.getElementById("mainFormTitle");
    if (titleEl) {
        if (mode === "dashboard") {
            titleEl.textContent = "Dashboard";
        } else if (mode === "agreement") {
            titleEl.textContent = "Agreement Form";
        } else if (mode === "createTenantNew") {
            titleEl.textContent = "Create Tenant – Active";
        } else if (mode === "addPastTenant") {
            titleEl.textContent = "Create Tenant – Past";
        } else if (mode === "viewTenants") {
            titleEl.textContent = "Tenant Directory";
        } else if (mode === "generateBill") {
            titleEl.textContent = "Generate Bills";
        } else if (mode === "payments") {
            titleEl.textContent = "Payments";
        }
    }

    // Toggle action buttons in Top and Bottom bars based on mode
    const toggleButtons = (selector, show) => {
        document.querySelectorAll(selector).forEach(btn => {
            if (show) btn.classList.remove("hidden");
            else btn.classList.add("hidden");
        });
    };

    if (mode === "dashboard") {
        toggleButtons(".btn-save-agreement", false);
        toggleButtons(".btn-export-docx", false);
        toggleButtons(".btn-create-new", false);
        toggleButtons(".btn-save-past", false);
    } else if (mode === "agreement") {
        toggleButtons(".btn-save-agreement", true);
        toggleButtons(".btn-export-docx", true);
        toggleButtons(".btn-create-new", false);
        toggleButtons(".btn-save-past", false);
    } else if (mode === "createTenantNew") {
        toggleButtons(".btn-save-agreement", false);
        toggleButtons(".btn-export-docx", false);
        toggleButtons(".btn-create-new", true);
        toggleButtons(".btn-save-past", false);
    } else if (mode === "addPastTenant") {
        toggleButtons(".btn-save-agreement", false);
        toggleButtons(".btn-export-docx", false);
        toggleButtons(".btn-create-new", false);
        toggleButtons(".btn-save-past", true);
    } else if (mode === "viewTenants" || mode === "generateBill" || mode === "payments") {
        toggleButtons(".btn-save-agreement", false);
        toggleButtons(".btn-export-docx", false);
        toggleButtons(".btn-create-new", false);
        toggleButtons(".btn-save-past", false);
    }

    // Toggle sidebar content (clauses for agreement, placeholder for tenant modes)
    const clausesAgreementContent = document.getElementById("clausesAgreementContent");
    const clausesPlaceholderContent = document.getElementById("clausesPlaceholderContent");
    const directorySidebarContent = document.getElementById("directorySidebarContent");
    const utilitySidebarContent = document.getElementById("utilitySidebarContent");
    const grnInput = document.getElementById("grn_number");
    const noGrnCheckbox = document.getElementById("grnNoCheckbox");

    if (clausesAgreementContent && clausesPlaceholderContent && directorySidebarContent) {
        const showClauses = mode === "agreement";
        const showDirectory = mode === "viewTenants";
        const showTenantSidebar = mode === "createTenantNew" || mode === "addPastTenant";
        const showUtilitySidebar = isDashboard || isGenerateBill || isPayments;

        smoothToggle(clausesAgreementContent, showClauses);
        smoothToggle(directorySidebarContent, showDirectory);
        smoothToggle(clausesPlaceholderContent, showTenantSidebar);
        smoothToggle(utilitySidebarContent, showUtilitySidebar);
    }

    if (grnInput && noGrnCheckbox) {
        const allowNoGrn = mode === "createTenantNew" || mode === "addPastTenant";
        noGrnCheckbox.closest("label")?.classList.toggle("hidden", !allowNoGrn);

        if (!allowNoGrn) {
            grnInput.disabled = false;
            grnInput.dataset.noGrn = "0";
            grnInput.value = "";
            if (grnInput.dataset.prevGrn) {
                grnInput.value = grnInput.dataset.prevGrn || grnInput.value;
                delete grnInput.dataset.prevGrn;
            }
            noGrnCheckbox.checked = false;
        }
    }

    // Show tenancy end date field only for past tenants
    const tenancyEndGroup = document.getElementById("tenancyEndGroup");
    if (tenancyEndGroup) {
        if (mode === "addPastTenant") {
            tenancyEndGroup.classList.remove("hidden");
        } else {
            tenancyEndGroup.classList.add("hidden");
        }
    }

    // Show actions card only in agreement mode
    const actionsCard = document.getElementById("actionsCard");
    if (actionsCard) {
        smoothToggle(actionsCard, mode === "agreement");
    }

    if (mode === "dashboard") {
        if (dashboardSection) dashboardSection.classList.remove("hidden");
        if (formSection) formSection.classList.add("hidden");
        if (tenantListSection) tenantListSection.classList.add("hidden");
        if (generateBillSection) generateBillSection.classList.add("hidden");
        if (paymentsSection) paymentsSection.classList.add("hidden");
        if (formRightSection) formRightSection.classList.remove("hidden");
        if (appLayout) {
            appLayout.classList.remove("layout-agreement");
            appLayout.classList.add("layout-directory");
        }
        return;
    }

    if (mode === "generateBill") {
        if (formSection) formSection.classList.add("hidden");
        if (tenantListSection) tenantListSection.classList.add("hidden");
        if (generateBillSection) generateBillSection.classList.remove("hidden");
        if (formRightSection) formRightSection.classList.remove("hidden");
        if (appLayout) {
            appLayout.classList.remove("layout-agreement");
            appLayout.classList.add("layout-directory");
        }

        return;
    }

    if (mode === "payments") {
        if (formSection) formSection.classList.add("hidden");
        if (tenantListSection) tenantListSection.classList.add("hidden");
        if (generateBillSection) generateBillSection.classList.add("hidden");
        if (paymentsSection) paymentsSection.classList.remove("hidden");
        if (formRightSection) formRightSection.classList.remove("hidden");
        if (appLayout) {
            appLayout.classList.remove("layout-agreement");
            appLayout.classList.add("layout-directory");
        }
        refreshPaymentsIfNeeded();
        return;
    }

    if (mode === "viewTenants") {
        if (formSection) formSection.classList.add("hidden");
        if (formRightSection) formRightSection.classList.remove("hidden");
        if (tenantListSection) tenantListSection.classList.remove("hidden");
        if (appLayout) {
            appLayout.classList.remove("layout-agreement");
            appLayout.classList.add("layout-directory");
        }

        loadTenantDirectory();
        return;
    }

    // Load draft data specific to this mode
    syncDraftUiForFlow(mode);
    loadDraftForFlow(mode);

    if (mode !== "viewTenants") {
        applyLandlordDefaultsToForm();
    }
}
