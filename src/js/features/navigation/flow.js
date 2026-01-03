/**
 * Flow/Navigation Feature Module
 * 
 * Handles switching between different application modes:
 * - Agreement mode (create agreement and save tenant)
 * - Create Tenant New mode (active tenant)
 */

import { currentFlow, setCurrentFlow } from "../../state.js";
import { isDraftDirty, loadDraftForFlow, openUnsavedDraftModal, syncDraftUiForFlow } from "../shared/drafts.js";
import { applyLandlordDefaultsToForm } from "../../api/config.js";
import { setNoGrnState } from "../tenants/formState.js";
import { smoothToggle } from "../../utils/ui.js";

const DRAFT_GUARDED_FLOWS = new Set(["agreement", "createTenantNew"]);
let tenantDirectoryInitialized = false;
let billingInitialized = false;
let paymentsInitialized = false;

async function ensureTenantDirectoryInitialized() {
    const mod = await import("../tenants/tenants.js");
    if (!tenantDirectoryInitialized) {
        mod.initTenantDirectory();
        tenantDirectoryInitialized = true;
    }
    return mod;
}

async function ensureBillingInitialized() {
    const mod = await import("../billing/billing.js");
    if (!billingInitialized) {
        mod.initBillingFeature();
        billingInitialized = true;
    }
    return mod;
}

async function ensurePaymentsInitialized() {
    const mod = await import("../billing/payments.js");
    if (!paymentsInitialized) {
        mod.initPaymentsFeature();
        paymentsInitialized = true;
    }
    return mod;
}

/**
 * Switches the application to a different flow/mode
 * Updates UI elements, button visibility, and form sections based on the mode
 * @param {"dashboard" | "agreement" | "createTenantNew" | "viewTenants" | "generateBill" | "payments" | "exportData"} mode - The flow mode to switch to
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
    const exportDataSection = document.getElementById("exportDataSection");

    const isDashboard = mode === "dashboard";
    const isGenerateBill = mode === "generateBill";
    const isPayments = mode === "payments";
    const isViewTenants = mode === "viewTenants";
    const isExportData = mode === "exportData";
    const isFormFlow = mode === "agreement" || mode === "createTenantNew";
    const toggleSection = (element, show) =>
        smoothToggle(element, show, show ? {} : { duration: 0 });

    toggleSection(dashboardSection, isDashboard);
    toggleSection(formSection, isFormFlow);
    smoothToggle(formRightSection, true);
    toggleSection(tenantListSection, isViewTenants);
    toggleSection(generateBillSection, isGenerateBill);
    toggleSection(paymentsSection, isPayments);
    toggleSection(exportDataSection, isExportData);

    // Update navigation button active states
    const navButtons = {
        dashboard: "navDashboardBtn",
        agreement: "navCreateAgreementBtn",
        createTenantNew: "navCreateTenantBtn",
        viewTenants: "navViewTenantsBtn",
        generateBill: "navGenerateBillBtn",
        payments: "navPaymentsBtn",
        exportData: "navExportDataBtn",
    };

    const activeClasses = ["bg-slate-900", "text-white", "shadow-sm"];
    const inactiveTextClasses = ["text-slate-700"];
    const hoverClass = "hover:bg-slate-100";

    Object.values(navButtons).forEach((id) => {
        const btn = document.getElementById(id);
        if (!btn) return;
        btn.classList.remove(...activeClasses);
        btn.classList.add(hoverClass, ...inactiveTextClasses);
    });

    const activeButtonId = navButtons[mode];
    if (activeButtonId) {
        const activeButton = document.getElementById(activeButtonId);
        if (activeButton) {
            activeButton.classList.remove(hoverClass, ...inactiveTextClasses);
            activeButton.classList.add(...activeClasses);
        }
    }

    // Update form title
    const titleEl = document.getElementById("mainFormTitle");
    const titleMap = {
        dashboard: "Dashboard",
        agreement: "Agreement Form",
        createTenantNew: "Create Tenant – Active",
        viewTenants: "Tenant Directory",
        generateBill: "Generate Bills",
        payments: "Payments",
        exportData: "Export data",
    };
    if (titleEl && titleMap[mode]) {
        titleEl.textContent = titleMap[mode];
    }

    // Toggle action buttons in Top and Bottom bars based on mode
    const toggleButtons = (selector, show) => {
        document.querySelectorAll(selector).forEach(btn => {
            if (show) btn.classList.remove("hidden");
            else btn.classList.add("hidden");
        });
    };

    const actionVisibility = {
        agreement: {
            ".btn-save-agreement": true,
            ".btn-export-docx": true,
            ".btn-create-new": false,
        },
        createTenantNew: {
            ".btn-save-agreement": false,
            ".btn-export-docx": false,
            ".btn-create-new": true,
        },
        default: {
            ".btn-save-agreement": false,
            ".btn-export-docx": false,
            ".btn-create-new": false,
        },
    };

    const actionConfig = actionVisibility[mode] || actionVisibility.default;
    Object.entries(actionConfig).forEach(([selector, show]) => toggleButtons(selector, show));

    // Toggle sidebar content (clauses for agreement, placeholder for tenant modes)
    const clausesAgreementContent = document.getElementById("clausesAgreementContent");
    const clausesPlaceholderContent = document.getElementById("clausesPlaceholderContent");
    const directorySidebarContent = document.getElementById("directorySidebarContent");
    const utilitySidebarContent = document.getElementById("utilitySidebarContent");
    const noGrnCheckbox = document.getElementById("grnNoCheckbox");

    if (clausesAgreementContent && clausesPlaceholderContent && directorySidebarContent) {
        const showClauses = mode === "agreement";
        const showDirectory = mode === "viewTenants";
        const showTenantSidebar = mode === "createTenantNew";
        const showUtilitySidebar = isDashboard || isGenerateBill || isPayments || isExportData;

        toggleSection(clausesAgreementContent, showClauses);
        toggleSection(directorySidebarContent, showDirectory);
        toggleSection(clausesPlaceholderContent, showTenantSidebar);
        toggleSection(utilitySidebarContent, showUtilitySidebar);
    }

    if (noGrnCheckbox) {
        const allowNoGrn = mode === "createTenantNew";
        noGrnCheckbox.closest("label")?.classList.toggle("hidden", !allowNoGrn);

        if (!allowNoGrn) {
            setNoGrnState({ checked: false, source: "flow" });
        }
    }

    // Show actions card only in agreement mode
    const actionsCard = document.getElementById("actionsCard");
    if (actionsCard) {
        toggleSection(actionsCard, mode === "agreement");
    }

    if (mode === "payments") {
        void ensurePaymentsInitialized().then((mod) => mod.refreshPaymentsIfNeeded());
    }

    if (mode === "viewTenants") {
        void ensureTenantDirectoryInitialized().then((mod) => mod.loadTenantDirectory());
    }

    if (mode === "generateBill") {
        void ensureBillingInitialized();
    }

    if (isFormFlow) {
        // Load draft data specific to this mode
        syncDraftUiForFlow(mode);
        loadDraftForFlow(mode);
    }

    if (mode !== "viewTenants") {
        applyLandlordDefaultsToForm();
    }
}
