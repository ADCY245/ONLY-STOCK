// The frontend can be served directly by Flask or from the development-only
// static server on ports 3000/3412. In the latter case, API requests must go to Flask.
const IS_LOCAL_STATIC_FRONTEND = ["3000", "3412"].includes(window.location.port);
const API_BASE_URL = window.ONLY_STOCK_API_BASE_URL || (
    window.location.protocol.startsWith("http") && !IS_LOCAL_STATIC_FRONTEND
        ? window.location.origin
        : "http://localhost:5412"
);
const API_CREDENTIALS_MODE = API_BASE_URL === window.location.origin ? "same-origin" : "include";

const state = {
    inventory: [],
    logs: [],
    user: null,
    users: [],
    excelPreview: null,
};

const ROLE_LEVELS = {
    user: 1,
    workshop: 2,
    admin: 3,
};

const PAGE_META = {
    overview: {
        eyebrow: "Dashboard",
        title: "Inventory Overview",
        description: "See your stock health and move into each workspace from the menu.",
    },
    "add-item": {
        eyebrow: "Workspace",
        title: "Add Inventory Item",
        description: "Create new stock records with category, quantity, and unit details.",
    },
    inventory: {
        eyebrow: "Workspace",
        title: "Inventory Manager",
        description: "Search, filter, update, and delete items from a dedicated page.",
    },
    warehouse: {
        eyebrow: "Workspace",
        title: "Warehouse View",
        description: "Interactive 3D warehouse and stock locations",
    },
    excel: {
        eyebrow: "Workspace",
        title: "Excel Tools",
        description: "Import, update, and export inventory using category-specific Excel workbooks.",
    },
    "inventory-adjustments": {
        eyebrow: "Workspace",
        title: "Inventory Adjustments",
        description: "Create formatted inventory adjustment XLS files from pasted rows or uploaded sheets.",
    },
    admin: {
        eyebrow: "Admin",
        title: "Admin Panel",
        description: "Manage users and review stock activity from one admin-only workspace.",
    },
};

const itemForm = document.getElementById("itemForm");
const loginForm = document.getElementById("loginForm");
const signupForm = document.getElementById("signupForm");
const forgotForm = document.getElementById("forgotForm");
const loginMessage = document.getElementById("loginMessage");
const signupMessage = document.getElementById("signupMessage");
const forgotMessage = document.getElementById("forgotMessage");
const authShell = document.getElementById("authShell");
const appShell = document.getElementById("appShell");
const authTabs = [...document.querySelectorAll("[data-auth-tab]")];
const authPanels = [...document.querySelectorAll("[data-auth-panel]")];
const excelForm = document.getElementById("excelForm");
const excelFileInput = document.getElementById("excelFile");
const refreshButton = document.getElementById("refreshButton");
const adminPanelButton = document.getElementById("adminPanelButton");
const excelExportButton = document.getElementById("excelExportButton");
const importTemplateButton = document.getElementById("importTemplateButton");
const currentStockButton = document.getElementById("currentStockButton");
const exportCurrentButton = document.getElementById("exportCurrentButton");
const excelPreviewButton = document.getElementById("excelPreviewButton");
const excelApplyButton = document.getElementById("excelApplyButton");
const excelUpdateWarning = document.getElementById("excelUpdateWarning");
const excelPreviewSection = document.getElementById("excelPreviewSection");
const excelValidationSection = document.getElementById("excelValidationSection");
const excelDetailSection = document.getElementById("excelDetailSection");
const excelApplySection = document.getElementById("excelApplySection");
const excelPreviewTotal = document.getElementById("excelPreviewTotal");
const excelSheetSummary = document.getElementById("excelSheetSummary");
const excelWorkbookWarnings = document.getElementById("excelWorkbookWarnings");
const excelValidCount = document.getElementById("excelValidCount");
const excelWarningCount = document.getElementById("excelWarningCount");
const excelErrorCount = document.getElementById("excelErrorCount");
const excelErrorsButton = document.getElementById("excelErrorsButton");
const excelErrorsPanel = document.getElementById("excelErrorsPanel");
const excelCategoryPreview = document.getElementById("excelCategoryPreview");
const excelApplyTitle = document.getElementById("excelApplyTitle");
const excelApplyDescription = document.getElementById("excelApplyDescription");
const excelAddCount = document.getElementById("excelAddCount");
const excelUpdateCount = document.getElementById("excelUpdateCount");
const excelDeleteCount = document.getElementById("excelDeleteCount");
const excelUnchangedCount = document.getElementById("excelUnchangedCount");
const excelUpdateScope = document.getElementById("excelUpdateScope");
const adjustmentsForm = document.getElementById("adjustmentsForm");
const adjustmentsUploadForm = document.getElementById("adjustmentsUploadForm");
const adjustmentsFile = document.getElementById("adjustmentsFile");
const adjustmentsDate = document.getElementById("adjustmentsDate");
const adjustmentsReason = document.getElementById("adjustmentsReason");
const adjustmentItems = document.getElementById("adjustmentItems");
const addAdjustmentItemButton = document.getElementById("addAdjustmentItemButton");
const previewAdjustmentsButton = document.getElementById("previewAdjustmentsButton");
const adjustmentsPreview = document.getElementById("adjustmentsPreview");
const adjustmentsPasteInput = document.getElementById("adjustmentsPasteInput");
const importPastedAdjustmentsButton = document.getElementById("importPastedAdjustmentsButton");
const adjustmentItemTemplate = document.getElementById("adjustmentItemTemplate");
const adjustmentBatchTemplate = document.getElementById("adjustmentBatchTemplate");
const inventoryTableBody = document.getElementById("inventoryTableBody");
const inventoryTableHead = document.getElementById("inventoryTableHead");
const dynamicItemFields = document.getElementById("dynamicItemFields");
const categoryList = document.getElementById("categoryList");
const categoryCount = document.getElementById("categoryCount");
const sidebarCategoryCount = document.getElementById("sidebarCategoryCount");
const sidebarItemCount = document.getElementById("sidebarItemCount");
const sidebarLogCount = document.getElementById("sidebarLogCount");
const overviewItemCount = document.getElementById("overviewItemCount");
const overviewCategoryCount = document.getElementById("overviewCategoryCount");
const overviewLowStockCount = document.getElementById("overviewLowStockCount");
const overviewLogBadge = document.getElementById("overviewLogBadge");
const overviewLogs = document.getElementById("overviewLogs");
const treeView = document.getElementById("treeView");
const logsList = document.getElementById("logsList");
const usersTableBody = document.getElementById("usersTableBody");
const statusText = document.getElementById("statusText");
const formMessage = document.getElementById("formMessage");
const excelMessage = document.getElementById("excelMessage");
const adjustmentsMessage = document.getElementById("adjustmentsMessage");
const adminUsersMessage = document.getElementById("adminUsersMessage");
const pageEyebrow = document.getElementById("pageEyebrow");
const pageTitle = document.getElementById("pageTitle");
const pageDescription = document.getElementById("pageDescription");
const pages = [...document.querySelectorAll("[data-page]")];
const pageLinks = [...document.querySelectorAll("[data-page-link]")];
const roleBoundElements = [...document.querySelectorAll("[data-min-role]")];
const writeColumns = [...document.querySelectorAll("[data-write-column]")];
const reasonDialog = document.getElementById("reasonDialog");
const reasonForm = document.getElementById("reasonForm");
const reasonTitle = document.getElementById("reasonTitle");
const reasonPrompt = document.getElementById("reasonPrompt");
const reasonInput = document.getElementById("reasonInput");
const reasonError = document.getElementById("reasonError");
const reasonCancelButton = document.getElementById("reasonCancelButton");

const searchInput = document.getElementById("searchInput");
const categoryFilter = document.getElementById("categoryFilter");
const brandFilter = document.getElementById("brandFilter");
const typeFilter = document.getElementById("typeFilter");
const lowStockOnly = document.getElementById("lowStockOnly");
const lowStockThreshold = document.getElementById("lowStockThreshold");
const thicknessFilter = document.getElementById("thicknessFilter");
const formCategorySelect = document.getElementById("formCategorySelect");
const inventoryHelpText = document.getElementById("inventoryHelpText");
const sessionEmail = document.getElementById("sessionEmail");
const sessionRole = document.getElementById("sessionRole");
const logoutButton = document.getElementById("logoutButton");

const CATEGORY_OPTIONS = [
    { code: "01", label: "Rubber Blankets" },
    { code: "02", label: "Metalback Blankets" },
    { code: "03", label: "Underlay Blanket" },
    { code: "04", label: "Blanket Barring" },
    { code: "05", label: "Calibrated Underpacking Paper" },
    { code: "06", label: "Calibrated Underpacking Film" },
    { code: "07", label: "Creasing Matrix" },
    { code: "08", label: "Cutting Rules" },
    { code: "09", label: "Creasing Rules" },
    { code: "10", label: "Litho Perforation Rules" },
    { code: "11", label: "Cutting String" },
    { code: "12", label: "Ejection Rubber" },
    { code: "13", label: "Strip Plate" },
    { code: "14", label: "Anti Marking Film" },
    { code: "15", label: "Ink Duct Foil" },
    { code: "16", label: "Productive Foil" },
    { code: "17", label: "Presspahn Sheets" },
    { code: "18", label: "Washing Solutions" },
    { code: "19", label: "Fountain Solutions" },
    { code: "20", label: "Plate Care Products" },
    { code: "21", label: "Roller Care Products" },
    { code: "22", label: "Blanket Maintenance Products" },
    { code: "23", label: "Auto Wash Cloth" },
    { code: "24", label: "ICP Paper" },
    { code: "25", label: "Spray Powder" },
    { code: "26", label: "Sponges" },
    { code: "27", label: "Dampening Hose" },
    { code: "28", label: "Tesamol Tape" },
    { code: "29", label: "CTCP Plates" },
];

function makeCategoryKey(value) {
    return String(value || "")
        .trim()
        .replace(/^\d+\s*-\s*/, "")
        .toLowerCase()
        .replace(/&/g, "and")
        .replace(/[^a-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "");
}

CATEGORY_OPTIONS.forEach((option) => {
    option.key = makeCategoryKey(option.label);
});

function getCategoryOption(value) {
    const text = String(value || "").trim().toLowerCase();
    if (!text) {
        return null;
    }
    return CATEGORY_OPTIONS.find((option) => (
        option.key === text ||
        option.label.toLowerCase() === text ||
        `${option.code} - ${option.label}`.toLowerCase() === text ||
        makeCategoryKey(text) === option.key
    )) || null;
}

function getCategoryKey(value) {
    return getCategoryOption(value)?.key || makeCategoryKey(value);
}

function getCategoryLabel(value) {
    return getCategoryOption(value)?.label || String(value || "").replace(/^\d+\s*-\s*/, "").trim();
}

const UNIT_OPTIONS = [
    "pcs",
    "box",
    "boxes",
    "pack",
    "packs",
    "roll",
    "rolls",
    "sheet",
    "sheets",
    "set",
    "sets",
    "kg",
    "g",
    "ltr",
    "ml",
];

const CHEMICAL_CATEGORIES = new Set([
    "Washing Solutions",
    "Fountain Solutions",
    "Plate Care Products",
    "Roller Care Products",
    "Blanket Maintenance Products",
]);

const GENERIC_DIMENSIONAL_CATEGORIES = new Set([
    "Blanket Barring",
    "Cutting String",
    "Ejection Rubber",
    "Strip Plate",
    "Anti Marking Film",
    "Ink Duct Foil",
    "Productive Foil",
    "Presspahn Sheets",
    "Auto Wash Cloth",
    "ICP Paper",
    "Dampening Hose",
    "Tesamol Tape",
]);

const CREASING_MATRIX_SIZES = {
    "9": [
        "0.3 X 1.0", "0.3 X 1.1", "0.3 X 1.2", "0.3 X 1.3", "0.3 X 1.5",
        "0.4 X 1.0", "0.4 X 1.1", "0.4 X 1.2", "0.4 X 1.3", "0.4 X 1.4",
        "0.4 X 1.5", "0.4 X 1.6", "0.4 X 1.7", "0.5 X 1.2", "0.5 X 1.3",
        "0.5 X 1.4", "0.5 X 1.5", "0.5 X 1.6", "0.5 X 1.7", "0.5 X 1.8",
        "0.5 X 1.9", "0.6 X 1.1", "0.6 X 1.5", "0.6 X 1.6", "0.6 X 1.7",
        "0.6 X 1.9", "0.6 X 2.1", "0.6 X 2.3", "0.6 X 2.5",
    ],
    "11": [
        "0.7 X 1.1", "0.7 X 1.5", "0.7 X 2.1", "0.7 X 2.3", "0.7 X 2.5",
        "0.7 X 2.7", "0.8 X 2.1", "0.8 X 2.3", "0.8 X 2.5", "0.8 X 2.7",
        "0.8 X 3.0",
    ],
    "13": ["1.0 X 3.0", "1.0 X 3.2", "1.0 X 3.5", "1.0 X 4.0", "1.4 X 5.0"],
};

const CTCP_PLATE_SIZES = {
    "0.30": ["650 X 550", "730 X 600", "620 X 482"],
    "0.20": ["520 X 400"],
};
const ROLL_PAPER_CATEGORY = "Calibrated Underpacking Paper";
const ROLL_PAPER_STOCK_UNIT = "m\u00B2";
const RUBBER_BLANKET_CATEGORY = "Rubber Blankets";
const RUBBER_BLANKET_STOCK_UNIT = "m\u00B2";
const ROLL_WIDTH_UNITS = ["mm", "m", "inch"];
const ROLL_LENGTH_UNITS = ["m", "mm", "inch"];
const ROLL_THICKNESS_UNITS = ["mm", "micron"];
const M3Z_ROLL_STORAGE_TYPE = "roll";
const M3Z_CUT_PIECE_STORAGE_TYPE = "cut_piece";
const M3Z_THICKNESS_OPTIONS = [
    { micron: 500, mm: 0.50, label: "500 micron (0.50 mm)" },
    { micron: 400, mm: 0.40, label: "400 micron (0.40 mm)" },
    { micron: 300, mm: 0.30, label: "300 micron (0.30 mm)" },
    { micron: 250, mm: 0.25, label: "250 micron (0.25 mm)" },
    { micron: 200, mm: 0.20, label: "200 micron (0.20 mm)" },
    { micron: 150, mm: 0.15, label: "150 micron (0.15 mm)" },
    { micron: 100, mm: 0.10, label: "100 micron (0.10 mm)" },
];

// Rubber Blanket names and validation rules are loaded from the shared JSON
// catalog used by the backend and Excel generator.
const INVENTORY_CONFIG_URL = "data/inventory-config.json?v=shared-catalog-20260831-v2";
let RUBBER_BLANKET_RULES = [];

function normalizeRubberBlanketConfigEntry(entry) {
    return {
        key: String(entry.key || "").trim(),
        name: String(entry.name || "").trim(),
        aliases: Array.isArray(entry.aliases) ? entry.aliases.map((value) => String(value).trim()).filter(Boolean) : [],
        thicknessMode: entry.thickness_mode === "fixed" ? "fixed" : "select",
        thickness: Number.isFinite(Number(entry.thickness)) ? Number(entry.thickness) : null,
        thicknessOptions: Array.isArray(entry.thickness_options)
            ? entry.thickness_options.map(Number).filter((value) => Number.isFinite(value))
            : [],
        widths: Array.isArray(entry.widths)
            ? entry.widths
                .filter((pair) => Array.isArray(pair) && pair.length === 2)
                .map(([nominal, actual]) => [Number(nominal), Number(actual)])
                .filter(([nominal, actual]) => Number.isFinite(nominal) && Number.isFinite(actual))
            : [],
        printTypes: Array.isArray(entry.print_types)
            ? entry.print_types.map((value) => String(value).trim().toUpperCase()).filter(Boolean)
            : [],
    };
}

async function loadInventoryConfig() {
    const response = await fetch(INVENTORY_CONFIG_URL, { cache: "no-store" });
    if (!response.ok) {
        throw new Error("Unable to load the inventory catalog");
    }
    const config = await response.json();
    const entries = Array.isArray(config?.rubber_blankets) ? config.rubber_blankets : [];
    const rules = entries.map(normalizeRubberBlanketConfigEntry).filter((entry) => entry.name && entry.widths.length && entry.thicknessOptions.length);
    if (!rules.length) {
        throw new Error("The inventory catalog has no valid Rubber Blanket entries");
    }
    RUBBER_BLANKET_RULES = rules;
}

const CHEMICAL_PRODUCTS = [
    { name: "Chem R-ol", category: "Washing Solutions", displayFormat: "Chem R-ol (5L Pack)", packSize: 5, unit: "ltr", containerType: "bottle", containersPerBox: [4] },
    { name: "FS Clean", category: "Fountain Solutions", displayFormat: "FS Clean (5L Pack)", packSize: 5, unit: "ltr", containerType: "bottle", containersPerBox: [4] },
    { name: "Anilox Clean", category: "Roller Care Products", displayFormat: "Anilox Clean 5L", packSize: 5, unit: "ltr", containerType: "bottle", containersPerBox: [4] },
    { name: "Roll-o-clean", category: "Roller Care Products", displayFormat: "Roll-o-clean 1kg", packSize: 1, unit: "kg", containerType: "bottle", containersPerBox: [12, 15, 18] },
    { name: "MT-R Clean", category: "Roller Care Products", displayFormat: "MT-R Clean 1L", packSize: 1, unit: "ltr", containerType: "bottle", containersPerBox: [12] },
    { name: "Blanket Clean", category: "Blanket Maintenance Products", displayFormat: "Blanket Clean 1L", packSize: 1, unit: "ltr", containerType: "bottle", containersPerBox: [5] },
    { name: "Blanket Clean UV", category: "Blanket Maintenance Products", displayFormat: "Blanket Clean UV 1L", packSize: 1, unit: "ltr", containerType: "bottle", containersPerBox: [5] },
    { name: "ALU Clean", category: "Plate Care Products", displayFormat: "ALU Clean 1L", packSize: 1, unit: "ltr", containerType: "bottle", containersPerBox: [12] },
    { name: "ALU Clean UV", category: "Plate Care Products", displayFormat: "ALU Clean UV 1L", packSize: 1, unit: "ltr", containerType: "bottle", containersPerBox: [12] },
    { name: "Calx De Glazer", category: "Roller Care Products", displayFormat: "Calx De Glazer 1L", packSize: 1, unit: "ltr", containerType: "bottle", containersPerBox: [12] },
];
const DEFAULT_ADJUSTMENT_ITEM_NAME = "Image - Print Master BL - 1070 mm - 1.95mm";
const FIXED_ADJUSTMENT_REASON = "Stock Update based on 1.04.26";
const DEFAULT_ADJUSTMENT_WAREHOUSE = "Main Location";

const CATEGORY_RULES = {
    "Rubber Blankets": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: false,
        requiresType: false,
        unitOptions: [RUBBER_BLANKET_STOCK_UNIT],
        defaultUnit: RUBBER_BLANKET_STOCK_UNIT,
        specialized: "rubber_blankets",
    },
    "Metalback Blankets": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: true,
        requiresType: true,
        unitOptions: ["pcs", "rolls"],
        defaultUnit: "pcs",
        supportsBatchRollNo: true,
    },
    "Underlay Blanket": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "micron",
        requiresBrand: true,
        requiresType: true,
        unitOptions: ["pcs", "rolls"],
        defaultUnit: "pcs",
    },
    [ROLL_PAPER_CATEGORY]: {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: false,
        requiresType: false,
        unitOptions: ["m²", "sheets"],
        defaultUnit: "m²",
        specialized: "calibrated_underpacking_paper",
    },
    "Calibrated Underpacking Film": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "micron",
        requiresBrand: true,
        requiresType: true,
        unitOptions: ["pcs", "rolls"],
        defaultUnit: "pcs",
    },
    "Creasing Matrix": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: false,
        requiresType: false,
        unitOptions: ["pkt"],
        defaultUnit: "pkt",
    },
    "CTCP Plates": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: false,
        requiresType: false,
        unitOptions: ["box"],
        defaultUnit: "box",
    },
    "Cutting Rules": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "pt",
        requiresBrand: true,
        requiresType: true,
        typeOptions: ["coil", "pkt"],
        unitOptions: ["coil", "pkt"],
        defaultUnit: "coil",
        unitLinkedToType: true,
    },
    "Creasing Rules": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "pt",
        requiresBrand: true,
        requiresType: true,
        typeOptions: ["coil", "pkt"],
        unitOptions: ["coil", "pkt"],
        defaultUnit: "coil",
        unitLinkedToType: true,
    },
    "Litho Perforation Rules": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "pt",
        requiresBrand: true,
        requiresType: true,
        typeOptions: ["coil", "pkt"],
        unitOptions: ["coil", "pkt"],
        defaultUnit: "coil",
        unitLinkedToType: true,
    },
    "Washing Solutions": {
        usesDimensions: false,
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: true,
        typeIsFormat: true,
        unitOptions: ["ltr", "kg"],
        defaultUnit: "ltr",
    },
    "Fountain Solutions": {
        usesDimensions: false,
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: true,
        typeIsFormat: true,
        unitOptions: ["ltr", "kg"],
        defaultUnit: "ltr",
    },
    "Plate Care Products": {
        usesDimensions: false,
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: true,
        typeIsFormat: true,
        unitOptions: ["ltr", "kg"],
        defaultUnit: "ltr",
    },
    "Roller Care Products": {
        usesDimensions: false,
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: true,
        typeIsFormat: true,
        unitOptions: ["ltr", "kg"],
        defaultUnit: "ltr",
    },
    "Blanket Maintenance Products": {
        usesDimensions: false,
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: true,
        typeIsFormat: true,
        unitOptions: ["ltr", "kg"],
        defaultUnit: "ltr",
    },
};

function setMessage(element, text, tone = "") {
    element.textContent = text || "";
    if (tone) {
        element.dataset.tone = tone;
    } else {
        delete element.dataset.tone;
    }
}

function getRoleLevel(role) {
    return ROLE_LEVELS[role] || 0;
}

function userHasRole(minRole) {
    return getRoleLevel(state.user?.role) >= getRoleLevel(minRole);
}

function isReadOnlyUser() {
    return state.user?.role === "user";
}

function setAuthTab(tabName) {
    authTabs.forEach((tab) => {
        tab.classList.toggle("is-active", tab.dataset.authTab === tabName);
    });
    authPanels.forEach((panel) => {
        panel.classList.toggle("is-active", panel.dataset.authPanel === tabName);
    });
}

function updateRoleVisibility() {
    roleBoundElements.forEach((element) => {
        const minRole = element.dataset.minRole;
        const allowed = !minRole || userHasRole(minRole);
        element.classList.toggle("is-hidden", !allowed);
    });

    writeColumns.forEach((cell) => {
        cell.classList.toggle("is-hidden", isReadOnlyUser());
    });

    inventoryHelpText.textContent = isReadOnlyUser()
        ? "You have read-only access. You can view overview data, inventory tree, and inventory items."
        : "Record stock coming in or going out. Quantity never goes below zero.";

    sessionEmail.textContent = state.user?.email || "Not signed in";
    sessionRole.textContent = state.user ? `Role: ${state.user.role}` : "Role";
}

function setAuthenticatedUser(user) {
    state.user = user || null;
    const isLoggedIn = Boolean(state.user);
    authShell.classList.toggle("is-hidden", isLoggedIn);
    appShell.classList.toggle("is-hidden", !isLoggedIn);
    updateRoleVisibility();

    if (!isLoggedIn) {
        state.inventory = [];
        state.logs = [];
        state.users = [];
        window.dispatchEvent(new CustomEvent("onlystock:inventory-updated", { detail: { items: [] } }));
        return;
    }

    const page = getCurrentPage();
    if ((page === "add-item" || page === "excel" || page === "inventory-adjustments") && !userHasRole("workshop")) {
        window.location.hash = "#overview";
    }
    if (page === "admin" && !userHasRole("admin")) {
        window.location.hash = "#overview";
    }
}

function getCurrentPage() {
    const page = window.location.hash.replace(/^#/, "");
    return PAGE_META[page] ? page : "overview";
}

function applyPageMeta(page) {
    const meta = PAGE_META[page];
    pageEyebrow.textContent = meta.eyebrow;
    pageTitle.textContent = meta.title;
    pageDescription.textContent = meta.description;
}

function showPage(page) {
    pages.forEach((section) => {
        section.classList.toggle("is-active", section.dataset.page === page);
    });

    pageLinks.forEach((link) => {
        link.classList.toggle("is-active", link.dataset.pageLink === page);
    });

    applyPageMeta(page);
    window.dispatchEvent(new CustomEvent("onlystock:page-changed", { detail: { page } }));
}

function getCategoryRule(category) {
    const categoryLabel = getCategoryLabel(category);
    return CATEGORY_RULES[categoryLabel] || {
        usesDimensions: GENERIC_DIMENSIONAL_CATEGORIES.has(categoryLabel),
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: false,
        unitLinkedToType: false,
        unitOptions: UNIT_OPTIONS,
        defaultUnit: "pcs",
    };
}

function requiresBatchRollNo(category, unit) {
    const rule = getCategoryRule(category);
    return Boolean(rule.supportsBatchRollNo && String(unit || "").trim().toLowerCase() === "rolls");
}

function isRollItem(itemOrUnit) {
    const unit = typeof itemOrUnit === "string" ? itemOrUnit : itemOrUnit?.unit;
    return String(unit || "").trim().toLowerCase() === "rolls";
}

function parsePositiveNumber(value) {
    const match = String(value ?? "").trim().match(/\d+(?:\.\d+)?/);
    const number = match ? Number(match[0]) : Number(value);
    return Number.isFinite(number) && number > 0 ? number : null;
}

function getRollWidthMeters(item) {
    const width = parsePositiveNumber(item?.width);
    return width === null ? null : width / 1000;
}

function getRollAreaSqm(widthValue, lengthValue) {
    const widthMeters = parsePositiveNumber(widthValue);
    const lengthMeters = parsePositiveNumber(lengthValue);
    if (widthMeters === null || lengthMeters === null) {
        return null;
    }
    return (widthMeters / 1000) * lengthMeters;
}

function formatInputNumber(value) {
    return Number(value).toFixed(6).replace(/0+$/, "").replace(/\.$/, "");
}

function formatFixedQuantity(value, decimals = 2) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
        return "-";
    }
    return new Intl.NumberFormat("en-IN", {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
    }).format(number);
}

function convertRollDimension(value, fromUnit, toUnit) {
    const metersPerUnit = { mm: 0.001, m: 1, inch: 0.0254 };
    const number = Number(value);
    if (!Number.isFinite(number) || !metersPerUnit[fromUnit] || !metersPerUnit[toUnit]) {
        return null;
    }
    return (number * metersPerUnit[fromUnit]) / metersPerUnit[toUnit];
}

function convertRollThickness(value, fromUnit, toUnit) {
    const number = Number(value);
    if (!Number.isFinite(number) || !ROLL_THICKNESS_UNITS.includes(fromUnit) || !ROLL_THICKNESS_UNITS.includes(toUnit)) {
        return null;
    }
    const micron = fromUnit === "mm" ? number * 1000 : number;
    return toUnit === "mm" ? micron / 1000 : micron;
}

function convertRollPaperControlValue(unitControl) {
    if (!unitControl) {
        return;
    }
    const valueName = unitControl.name.replace(/_unit$/, "");
    const valueControl = dynamicItemFields.querySelector(`[name="${valueName}"]`);
    if (!valueControl) {
        return;
    }
    const nextUnit = unitControl.value;
    const previousUnit = unitControl.dataset.previousUnit || nextUnit;
    const converter = valueName === "thickness" ? convertRollThickness : convertRollDimension;
    const converted = converter(valueControl.value, previousUnit, nextUnit);
    if (converted !== null) {
        valueControl.value = formatInputNumber(converted);
    }
    unitControl.dataset.previousUnit = nextUnit;
}

function getRollPaperCalculation() {
    const width = Number(dynamicItemFields.querySelector('[name="width"]')?.value);
    const widthUnit = dynamicItemFields.querySelector('[name="width_unit"]')?.value;
    const length = Number(dynamicItemFields.querySelector('[name="length"]')?.value);
    const lengthUnit = dynamicItemFields.querySelector('[name="length_unit"]')?.value;
    const thickness = Number(dynamicItemFields.querySelector('[name="thickness"]')?.value);
    const thicknessUnit = dynamicItemFields.querySelector('[name="thickness_unit"]')?.value;
    const rolls = Number(dynamicItemFields.querySelector('[name="number_of_rolls"]')?.value);
    const widthMeters = convertRollDimension(width, widthUnit, "m");
    const lengthMeters = convertRollDimension(length, lengthUnit, "m");
    const thicknessMicron = convertRollThickness(thickness, thicknessUnit, "micron");
    const areaPerRoll = widthMeters === null || lengthMeters === null ? null : widthMeters * lengthMeters;
    const totalArea = areaPerRoll === null || !Number.isFinite(rolls) ? null : areaPerRoll * rolls;
    return { widthMeters, lengthMeters, thicknessMicron, areaPerRoll, totalArea, rolls };
}

function getM3ZCalculation() {
    const storageType = normalizeM3ZStorageType(dynamicItemFields.querySelector('[name="storage_type"]')?.value);
    const width = Number(dynamicItemFields.querySelector('[name="width"]')?.value);
    const widthUnit = dynamicItemFields.querySelector('[name="width_unit"]')?.value;
    const length = Number(dynamicItemFields.querySelector('[name="length"]')?.value);
    const lengthUnit = dynamicItemFields.querySelector('[name="length_unit"]')?.value;
    const thicknessMicron = Number(dynamicItemFields.querySelector('[name="thickness_micron"]')?.value);
    const widthMeters = convertRollDimension(width, widthUnit, "m");
    const lengthMeters = convertRollDimension(length, lengthUnit, "m");
    const areaPerSheet = widthMeters === null || lengthMeters === null || width <= 0 || length <= 0
        ? null
        : widthMeters * lengthMeters;
    const countName = storageType === M3Z_ROLL_STORAGE_TYPE ? "number_of_rolls" : "number_of_sheets";
    const count = Number(dynamicItemFields.querySelector(`[name="${countName}"]`)?.value);
    const validCount = Number.isInteger(count) && count > 0 ? count : null;
    const totalArea = storageType === M3Z_ROLL_STORAGE_TYPE && areaPerSheet !== null && validCount !== null
        ? areaPerSheet * validCount
        : null;
    return {
        storageType,
        widthMeters,
        lengthMeters,
        thicknessMicron: M3Z_THICKNESS_OPTIONS.some((option) => option.micron === thicknessMicron) ? thicknessMicron : null,
        areaPerSheet,
        areaPerRoll: areaPerSheet,
        count: validCount,
        totalArea,
        totalStock: storageType === M3Z_ROLL_STORAGE_TYPE ? totalArea : validCount,
    };
}

function getRubberBlanketRule(name) {
    return RUBBER_BLANKET_RULES.find((rule) => rule.name === name) || null;
}

function getRubberBlanketWidths(rule, thickness) {
    if (!rule) {
        return [];
    }
    if (rule.widthsByThickness) {
        return rule.widthsByThickness[Number(thickness).toFixed(2)] || [];
    }
    return rule.widths || [];
}

function getSelectedRubberBlanketWidth() {
    const widthControl = dynamicItemFields.querySelector('[name="blanket_width"]');
    const option = widthControl?.options[widthControl.selectedIndex];
    if (!option) {
        return null;
    }
    const nominal = Number(option.dataset.nominal);
    const actual = Number(option.dataset.actual);
    return Number.isFinite(nominal) && nominal > 0 && Number.isFinite(actual) && actual > 0
        ? { nominal, actual }
        : null;
}

function getRubberBlanketCalculation() {
    const storageType = normalizeM3ZStorageType(dynamicItemFields.querySelector('[name="storage_type"]')?.value);
    const width = getSelectedRubberBlanketWidth();
    const length = Number(dynamicItemFields.querySelector('[name="length"]')?.value);
    const lengthUnit = dynamicItemFields.querySelector('[name="length_unit"]')?.value;
    const countName = storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? "number_of_sheets" : "number_of_rolls";
    const count = Number(dynamicItemFields.querySelector(`[name="${countName}"]`)?.value);
    const lengthMeters = length > 0 ? convertRollDimension(length, lengthUnit, "m") : null;
    const actualWidthMeters = width ? width.actual / 1000 : null;
    const areaPerPiece = actualWidthMeters === null || lengthMeters === null
        ? null
        : actualWidthMeters * lengthMeters;
    const validCount = Number.isInteger(count) && count > 0 ? count : null;
    const totalArea = areaPerPiece === null || validCount === null || storageType !== M3Z_ROLL_STORAGE_TYPE
        ? null
        : areaPerPiece * validCount;
    return {
        storageType,
        width,
        actualWidthMeters,
        lengthMeters,
        count: validCount,
        areaPerRoll: storageType === M3Z_ROLL_STORAGE_TYPE ? areaPerPiece : null,
        areaPerSheet: areaPerPiece,
        totalArea,
        totalStock: storageType === M3Z_ROLL_STORAGE_TYPE ? totalArea : validCount,
    };
}

function roundStockQuantity(value) {
    return Math.round((Number(value) + Number.EPSILON) * 10000) / 10000;
}

function formatQuantity(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
        return String(value ?? "");
    }
    return new Intl.NumberFormat("en-IN", {
        maximumFractionDigits: 4,
    }).format(number);
}

function getDisplayUnit(unit) {
    return isRollItem(unit) ? "sq.m" : unit;
}

function getSpecializedKind(item) {
    if (item?.packaging?.kind) {
        return item.packaging.kind;
    }
    if (item?.category === ROLL_PAPER_CATEGORY) {
        return "calibrated_underpacking_paper";
    }
    if (item?.category === RUBBER_BLANKET_CATEGORY && (item?.blanket_name || item?.packaging?.kind === "rubber_blankets")) {
        return "rubber_blankets";
    }
    if (item?.category === "Creasing Matrix") {
        return "creasing_matrix";
    }
    if (item?.category === "CTCP Plates") {
        return "ctcp_plates";
    }
    if (CHEMICAL_CATEGORIES.has(item?.category) && getChemicalProduct(getDisplayValue(item?.brand))) {
        return "chemical";
    }
    return "generic";
}

function getItemPackaging(item) {
    if (item?.packaging?.kind) {
        return item.packaging;
    }
    const kind = getSpecializedKind(item);
    if (kind === "creasing_matrix") {
        return { kind, units_per_box: 10, container_type: "packet" };
    }
    if (kind === "ctcp_plates") {
        return { kind, sheets_per_box: 50, container_type: "box" };
    }
    if (kind === "rubber_blankets") {
        const storageType = item?.storage_type || M3Z_ROLL_STORAGE_TYPE;
        if (storageType === M3Z_CUT_PIECE_STORAGE_TYPE) {
            return { kind, stock_unit: "sheets", movement_units: ["sheets"], storage_type: storageType };
        }
        return { kind, stock_unit: RUBBER_BLANKET_STOCK_UNIT, movement_units: ["rolls", "m²"] };
    }
    if (kind === "calibrated_underpacking_paper") {
        const storageType = item?.storage_type || M3Z_ROLL_STORAGE_TYPE;
        return {
            kind,
            stock_unit: storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? "sheets" : ROLL_PAPER_STOCK_UNIT,
            movement_units: [storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? "sheets" : "rolls"],
            storage_type: storageType,
        };
    }
    if (kind === "chemical") {
        const product = getChemicalProduct(getDisplayValue(item.brand));
        const typeMatch = String(item.type || "").match(/\/\s*(\d+)\s+per\s+box$/i);
        const containersPerBox = typeMatch ? Number(typeMatch[1]) : product.containersPerBox[0];
        return {
            kind,
            pack_size: product.packSize,
            pack_unit: product.unit,
            container_type: product.containerType,
            containers_per_box: containersPerBox,
            display_format: product.displayFormat,
        };
    }
    return null;
}

function getItemPackaging(item) {
    const storedPackaging = item?.packaging;
    const kind = storedPackaging?.kind || getSpecializedKind(item);
    if (kind === "rubber_blankets") {
        const storageType = item?.storage_type || M3Z_ROLL_STORAGE_TYPE;
        return {
            ...(storedPackaging || {}),
            kind,
            stock_unit: storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? "sheets" : RUBBER_BLANKET_STOCK_UNIT,
            movement_units: storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? ["sheets"] : ["rolls", RUBBER_BLANKET_STOCK_UNIT],
            storage_type: storageType,
        };
    }
    if (kind === "calibrated_underpacking_paper") {
        const storageType = item?.storage_type || M3Z_ROLL_STORAGE_TYPE;
        return {
            ...(storedPackaging || {}),
            kind,
            stock_unit: storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? "sheets" : ROLL_PAPER_STOCK_UNIT,
            movement_units: storageType === M3Z_CUT_PIECE_STORAGE_TYPE ? ["sheets"] : ["rolls"],
            storage_type: storageType,
        };
    }
    if (storedPackaging?.kind) {
        return storedPackaging;
    }
    if (kind === "creasing_matrix") {
        return { kind, units_per_box: 10, container_type: "packet" };
    }
    if (kind === "ctcp_plates") {
        return { kind, sheets_per_box: 50, container_type: "box" };
    }
    if (kind === "chemical") {
        const product = getChemicalProduct(getDisplayValue(item.brand));
        if (!product) {
            return null;
        }
        const typeMatch = String(item.type || "").match(/\/\s*(\d+)\s+per\s+box$/i);
        const containersPerBox = typeMatch ? Number(typeMatch[1]) : product.containersPerBox[0];
        return {
            kind,
            pack_size: product.packSize,
            pack_unit: product.unit,
            container_type: product.containerType,
            containers_per_box: containersPerBox,
            display_format: product.displayFormat,
        };
    }
    return null;
}

function getStockBreakdown(item) {
    if (item?.stock_breakdown) {
        return item.stock_breakdown;
    }
    const packaging = getItemPackaging(item);
    const quantity = Number(item?.quantity || 0);
    if (!packaging) {
        return null;
    }
    if (packaging.kind === "creasing_matrix") {
        const packets = Math.round(quantity);
        return { packets, boxes: Math.floor(packets / 10), loose_units: packets % 10 };
    }
    if (packaging.kind === "ctcp_plates") {
        const boxes = Math.round(quantity);
        return { boxes, total_sheets: boxes * 50 };
    }
    if (packaging.kind === "chemical") {
        const containers = quantity / packaging.pack_size;
        const wholeContainers = Number.isInteger(containers) ? containers : containers;
        return {
            containers: wholeContainers,
            boxes: Number.isInteger(containers) ? Math.floor(containers / packaging.containers_per_box) : null,
            loose_units: Number.isInteger(containers) ? containers % packaging.containers_per_box : null,
            normalized_quantity: quantity,
            normalized_unit: item.unit,
        };
    }
    if (packaging.kind === "calibrated_underpacking_paper") {
        const areaPerRoll = getRollPaperAreaPerRoll(item);
        const storageType = item?.storage_type || M3Z_ROLL_STORAGE_TYPE;
        if (storageType === M3Z_CUT_PIECE_STORAGE_TYPE) {
            return {
                sheets: Number.isInteger(Number(item?.number_of_sheets)) ? Number(item.number_of_sheets) : null,
                area_per_sheet_sqm: Number.isFinite(areaPerRoll) ? areaPerRoll : null,
                total_sheets: quantity,
                normalized_quantity: quantity,
                normalized_unit: item.unit || "sheets",
            };
        }
        return {
            rolls: Number.isInteger(Number(item?.number_of_rolls)) ? Number(item.number_of_rolls) : null,
            area_per_roll_sqm: areaPerRoll,
            total_area_sqm: quantity,
            normalized_quantity: quantity,
            normalized_unit: item.unit || ROLL_PAPER_STOCK_UNIT,
        };
    }
    if (packaging.kind === "rubber_blankets") {
        const storedArea = Number(item?.area_per_roll_sqm);
        const actualWidth = Number(item?.actual_width);
        const length = Number(item?.length ?? item?.height);
        const lengthMeters = convertRollDimension(length, item?.length_unit || "m", "m");
        const areaPerRoll = Number.isFinite(storedArea) && storedArea > 0
            ? storedArea
            : (Number.isFinite(actualWidth) && actualWidth > 0 && lengthMeters !== null
                ? actualWidth / 1000 * lengthMeters
                : null);
        if ((item?.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
            return {
                sheets: Number.isInteger(Number(item?.number_of_sheets)) ? Number(item.number_of_sheets) : null,
                area_per_sheet_sqm: Number(item?.area_per_sheet_sqm) || areaPerRoll,
                total_sheets: quantity,
                normalized_quantity: quantity,
                normalized_unit: item.unit || "sheets",
            };
        }
        return {
            rolls: Number.isInteger(Number(item?.number_of_rolls)) ? Number(item.number_of_rolls) : null,
            area_per_roll_sqm: areaPerRoll,
            total_area_sqm: quantity,
            normalized_quantity: quantity,
            normalized_unit: item.unit || RUBBER_BLANKET_STOCK_UNIT,
        };
    }
    return null;
}

function getRollPaperAreaPerRoll(item) {
    const storedArea = Number(item?.area_per_roll_sqm);
    if (Number.isFinite(storedArea) && storedArea > 0) {
        return storedArea;
    }
    const widthMeters = convertRollDimension(item?.width, item?.width_unit, "m");
    const lengthMeters = convertRollDimension(item?.length ?? item?.height, item?.length_unit, "m");
    return widthMeters === null || lengthMeters === null ? null : widthMeters * lengthMeters;
}

function getRollPaperThicknessMicron(item) {
    const storedThickness = Number(item?.thickness_micron);
    if (Number.isFinite(storedThickness) && storedThickness > 0) {
        return storedThickness;
    }
    return convertRollThickness(item?.thickness, item?.thickness_unit, "micron");
}

function formatItemSize(item) {
    if (item?.width && item?.height) {
        return `${item.width} X ${item.height}`;
    }
    return String(item?.size || "-").replace(/\s+x\s+/i, " X ");
}

function formatChemicalUnit(unit) {
    return String(unit || "").toLowerCase() === "ltr" ? "L" : unit;
}

function getSpecializedStockSummary(item) {
    const kind = getSpecializedKind(item);
    const breakdown = getStockBreakdown(item) || {};
    if (kind === "calibrated_underpacking_paper") {
        if ((item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
            return formatQuantity(breakdown.sheets ?? item.number_of_sheets ?? item.quantity) + " sheets";
        }
        return formatQuantity(breakdown.rolls ?? item.number_of_rolls ?? 0) + " rolls | " + formatFixedQuantity(breakdown.total_area_sqm ?? item.quantity) + " m²";
    }
    if (kind === "rubber_blankets") {
        return `${formatQuantity(breakdown.rolls ?? item.number_of_rolls ?? 0)} roll(s) · ${formatFixedQuantity(item.quantity)} m²`;
    }
    if (kind === "creasing_matrix") {
        return `${formatQuantity(breakdown.packets ?? item.quantity)} pkt · ${formatQuantity(breakdown.boxes || 0)} boxes · ${formatQuantity(breakdown.loose_units || 0)} loose`;
    }
    if (kind === "ctcp_plates") {
        return `${formatQuantity(breakdown.boxes ?? item.quantity)} boxes · ${formatQuantity(breakdown.total_sheets || 0)} sheets`;
    }
    if (kind === "chemical") {
        const packaging = getItemPackaging(item);
        const containerType = packaging?.container_type || "container";
        return `${formatQuantity(breakdown.containers || 0)} ${containerType}${Number(breakdown.containers) === 1 ? "" : "s"} · ${formatQuantity(item.quantity)} ${formatChemicalUnit(item.unit)}`;
    }
    return `${formatQuantity(item.quantity)} ${getDisplayUnit(item.unit)}`;
}

function formatIstDateTime(value) {
    if (!value) {
        return "";
    }
    return new Intl.DateTimeFormat("en-IN", {
        dateStyle: "short",
        timeStyle: "medium",
        timeZone: "Asia/Kolkata",
    }).format(new Date(value));
}

function convertMovementToSqm(item, amount, movementUnit) {
    if (movementUnit === "sqm") {
        return amount;
    }
    const widthMeters = getRollWidthMeters(item);
    if (widthMeters === null) {
        throw new Error("roll width is required to convert this movement");
    }
    if (movementUnit === "mtr") {
        return widthMeters * amount;
    }
    if (movementUnit === "inch") {
        return widthMeters * amount * 0.0254;
    }
    return amount;
}

function getDynamicFormValues() {
    return Object.fromEntries(new FormData(itemForm).entries());
}

function setDynamicControlValue(name, value) {
    const control = dynamicItemFields.querySelector(`[name="${name}"]`);
    if (control && value !== undefined && value !== null) {
        control.value = String(value);
    }
}

function getChemicalProduct(productName) {
    return CHEMICAL_PRODUCTS.find((product) => product.name === productName) || null;
}

function getChemicalProductsForCategory(category) {
    return CHEMICAL_PRODUCTS.filter((product) => product.category === category);
}

function renderAddItemForm(category, values = {}) {
    const categoryKey = getCategoryKey(category);
    const categoryLabel = getCategoryLabel(categoryKey);
    if (!categoryKey) {
        dynamicItemFields.innerHTML = '<p class="empty-state dynamic-form-empty">Select a category to see its item-specific fields.</p>';
        return;
    }
    if (categoryKey === "creasing_matrix") {
        renderCreasingMatrixForm(values);
        return;
    }
    if (categoryKey === "ctcp_plates") {
        renderCTCPPlatesForm(values);
        return;
    }
    if (categoryKey === "calibrated_underpacking_paper") {
        renderCalibratedUnderpackingPaperForm(values);
        return;
    }
    if (categoryKey === "rubber_blankets") {
        renderRubberBlanketForm(values);
        return;
    }
    if (CHEMICAL_CATEGORIES.has(categoryLabel)) {
        renderChemicalForm(categoryLabel, values);
        return;
    }
    renderGenericAddItemForm(categoryLabel, values);
}

function rubberBlanketWidthKey(nominal, actual) {
    return `${Number(nominal)}|${Number(actual)}`;
}

function renderRubberBlanketForm(values = {}) {
    dynamicItemFields.innerHTML = `
        <div class="specialized-form-heading">
            <strong>Add Item - Rubber Blankets</strong>
            <span>Choose how you have this material.</span>
        </div>
        <label for="rubberBlanketStorageTypeSelect">
            Storage Type
            <select id="rubberBlanketStorageTypeSelect" name="storage_type" required>
                <option value="roll">As Rolls (stored in Sq.m.)</option>
                <option value="cut_piece">As Cut Pieces (stored in Sheets)</option>
            </select>
        </label>
        <div class="dynamic-item-fields" data-rubber-mode-fields></div>
        <div class="calculation-card rubber-blanket-calculation" data-calculation="rubber-blankets" aria-live="polite"></div>
        <p class="field-hint">Roll stock is normalized to square metres. Cut pieces are stored and counted in Sheets.</p>
    `;
    setDynamicControlValue("storage_type", normalizeM3ZStorageType(values.storage_type));
    renderRubberBlanketModeFields(values);
}

function renderRubberBlanketModeFields(values = {}) {
    const modeFields = dynamicItemFields.querySelector("[data-rubber-mode-fields]");
    const storageControl = dynamicItemFields.querySelector('[name="storage_type"]');
    if (!modeFields || !storageControl) {
        return;
    }
    const storageType = normalizeM3ZStorageType(storageControl.value);
    modeFields.innerHTML = `
        <label for="rubberBlanketNameSelect">
            Blanket Name
            <select id="rubberBlanketNameSelect" name="blanket_name" required>
                ${RUBBER_BLANKET_RULES.map((rule) => `<option value="${rule.name}">${rule.name}</option>`).join("")}
            </select>
        </label>
        <label id="rubberBlanketThicknessField" for="rubberBlanketThicknessControl"></label>
        <label for="rubberBlanketWidthSelect">
            Width
            <select id="rubberBlanketWidthSelect" name="blanket_width" required></select>
            <select id="rubberBlanketWidthUnitSelect" name="width_unit" required>
                ${ROLL_WIDTH_UNITS.map((unit) => `<option value="${unit}">${unit}</option>`).join("")}
            </select>
            <small id="rubberBlanketActualWidth" class="field-hint"></small>
        </label>
        <label for="rubberBlanketLengthInput">
            Length
            <input id="rubberBlanketLengthInput" type="number" name="length" min="0" step="0.001" placeholder="30.00" required>
        </label>
        <label for="rubberBlanketLengthUnitSelect">
            Length Unit
            <select id="rubberBlanketLengthUnitSelect" name="length_unit" required>
                ${ROLL_LENGTH_UNITS.map((unit) => `<option value="${unit}">${unit}</option>`).join("")}
            </select>
        </label>
        <label id="rubberBlanketPrintTypeField" for="rubberBlanketPrintTypeSelect" hidden></label>
        ${storageType === M3Z_ROLL_STORAGE_TYPE
            ? `<label for="rubberBlanketRollQuantityInput">Number of Rolls<input id="rubberBlanketRollQuantityInput" type="number" name="number_of_rolls" min="1" step="1" placeholder="5" required></label>
               <input type="hidden" name="unit" value="${RUBBER_BLANKET_STOCK_UNIT}">`
            : `<label for="rubberBlanketSheetsInput">Number of Sheets<input id="rubberBlanketSheetsInput" type="number" name="number_of_sheets" min="1" step="1" placeholder="50" required></label>
               <input type="hidden" name="unit" value="sheets">`}
    `;

    const requestedName = values.blanket_name || values.brand;
    const selectedRule = getRubberBlanketRule(requestedName) || RUBBER_BLANKET_RULES[0];
    setDynamicControlValue("blanket_name", selectedRule.name);
    setDynamicControlValue("length", values.length ?? values.height ?? "");
    setDynamicControlValue("length_unit", values.length_unit || "m");
    setDynamicControlValue("width_unit", values.width_unit || "mm");
    if (storageType === M3Z_ROLL_STORAGE_TYPE) {
        setDynamicControlValue("number_of_rolls", values.number_of_rolls ?? values.rolls ?? "");
    } else {
        setDynamicControlValue("number_of_sheets", values.number_of_sheets ?? values.sheets ?? "");
    }
    const widthUnitControl = dynamicItemFields.querySelector('[name="width_unit"]');
    const lengthUnitControl = dynamicItemFields.querySelector('[name="length_unit"]');
    widthUnitControl.dataset.previousUnit = widthUnitControl.value;
    lengthUnitControl.dataset.previousUnit = lengthUnitControl.value;
    updateRubberBlanketDependentFields(values);
}

function updateRubberBlanketDependentFields(values = {}) {
    const nameControl = dynamicItemFields.querySelector('[name="blanket_name"]');
    const thicknessField = document.getElementById("rubberBlanketThicknessField");
    const widthControl = dynamicItemFields.querySelector('[name="blanket_width"]');
    const printTypeField = document.getElementById("rubberBlanketPrintTypeField");
    const widthUnit = dynamicItemFields.querySelector('[name="width_unit"]')?.value || "mm";
    const actualWidthOutput = document.getElementById("rubberBlanketActualWidth");
    const rule = getRubberBlanketRule(nameControl?.value);
    if (!rule || !thicknessField || !widthControl || !printTypeField) {
        return;
    }

    const requestedThickness = Number(values.thickness ?? dynamicItemFields.querySelector('[name="thickness"]')?.value);
    const thickness = rule.thicknessOptions.includes(requestedThickness)
        ? requestedThickness
        : (rule.thickness ?? rule.thicknessOptions[0]);
    if (rule.thicknessMode === "fixed") {
        thicknessField.innerHTML = `Thickness<input id="rubberBlanketThicknessControl" type="number" name="thickness" value="${thickness.toFixed(2)}" readonly><input type="hidden" name="thickness_unit" value="mm"><small class="field-hint">Fixed by blanket master.</small>`;
    } else {
        thicknessField.innerHTML = `Thickness<select id="rubberBlanketThicknessControl" name="thickness" required>${rule.thicknessOptions.map((option) => `<option value="${option.toFixed(2)}">${option.toFixed(2)} mm</option>`).join("")}</select><input type="hidden" name="thickness_unit" value="mm">`;
        setDynamicControlValue("thickness", thickness.toFixed(2));
    }

    const widths = getRubberBlanketWidths(rule, thickness);
    const currentOption = widthControl.options[widthControl.selectedIndex];
    const preferredWidth = values.blanket_width || (
        values.nominal_width || values.width
            ? rubberBlanketWidthKey(values.nominal_width ?? values.width, values.actual_width ?? values.width)
            : currentOption?.value
    );
    widthControl.innerHTML = widths.map(([nominal, actual]) => {
        const displayNominal = convertRollDimension(nominal, "mm", widthUnit);
        const displayActual = convertRollDimension(actual, "mm", widthUnit);
        const label = nominal === actual
            ? `${formatInputNumber(displayNominal)} ${widthUnit}`
            : `${formatInputNumber(displayNominal)} ${widthUnit} (Actual ${formatInputNumber(displayActual)} ${widthUnit})`;
        return `<option value="${rubberBlanketWidthKey(nominal, actual)}" data-nominal="${nominal}" data-actual="${actual}">${label}</option>`;
    }).join("");
    if ([...widthControl.options].some((option) => option.value === preferredWidth)) {
        widthControl.value = preferredWidth;
    }
    const selectedWidth = getSelectedRubberBlanketWidth();
    if (actualWidthOutput) {
        actualWidthOutput.textContent = selectedWidth
            ? `Actual Width: ${formatInputNumber(convertRollDimension(selectedWidth.actual, "mm", widthUnit))} ${widthUnit}`
            : "";
    }

    if (rule.printTypes.length) {
        printTypeField.hidden = false;
        printTypeField.innerHTML = `Print Type<select id="rubberBlanketPrintTypeSelect" name="print_type" required>${rule.printTypes.map((option) => `<option value="${option}">${option === "P" ? "Printed (P)" : "Without Print (W/O)"}</option>`).join("")}</select>`;
        const requestedPrintType = values.print_type || values.type;
        if (rule.printTypes.includes(requestedPrintType)) {
            setDynamicControlValue("print_type", requestedPrintType);
        }
    } else {
        printTypeField.hidden = true;
        printTypeField.innerHTML = "";
    }
    updateSpecializedCalculation();
}

function renderCreasingMatrixForm(values = {}) {
    dynamicItemFields.innerHTML = `
        <div class="specialized-form-heading">
            <strong>Creasing Matrix</strong>
            <span>10 packets = 1 box</span>
        </div>
        <label for="creasingThicknessSelect">
            Thickness
            <select id="creasingThicknessSelect" name="thickness" required>
                <option value="9">9 mm</option>
                <option value="11">11 mm</option>
                <option value="13">13 mm</option>
            </select>
        </label>
        <label for="creasingSizeSelect">
            Size
            <select id="creasingSizeSelect" name="size" required></select>
        </label>
        <label for="creasingQuantityInput">
            Quantity in packets
            <input id="creasingQuantityInput" type="number" name="quantity" min="0" step="1" placeholder="23" required>
        </label>
        <input type="hidden" name="unit" value="pkt">
        <div class="calculation-card" data-calculation="creasing" aria-live="polite"></div>
    `;
    setDynamicControlValue("thickness", values.thickness || "9");
    updateCreasingSizeOptions(values.size);
    setDynamicControlValue("quantity", values.quantity || "");
    updateSpecializedCalculation();
}

function updateCreasingSizeOptions(preferredSize) {
    const thicknessControl = dynamicItemFields.querySelector('[name="thickness"]');
    const sizeControl = dynamicItemFields.querySelector('[name="size"]');
    if (!thicknessControl || !sizeControl) {
        return;
    }
    const sizes = CREASING_MATRIX_SIZES[thicknessControl.value] || [];
    const currentSize = preferredSize || sizeControl.value;
    sizeControl.innerHTML = sizes.map((size) => `<option value="${size}">${size}</option>`).join("");
    if (sizes.includes(currentSize)) {
        sizeControl.value = currentSize;
    }
}

function renderChemicalForm(category, values = {}) {
    const products = getChemicalProductsForCategory(category);
    dynamicItemFields.innerHTML = `
        <label for="chemicalProductSelect">
            Product
            <select id="chemicalProductSelect" name="product" required>
                ${products.map((product) => `<option value="${product.name}">${product.displayFormat}</option>`).join("")}
            </select>
        </label>
        <label for="chemicalContainersPerBoxSelect">
            Bottles per box
            <select id="chemicalContainersPerBoxSelect" name="containers_per_box" required></select>
        </label>
        <label for="chemicalBoxesInput">
            Boxes
            <input id="chemicalBoxesInput" type="number" name="boxes" min="0" step="1" value="0" required>
        </label>
        <label for="chemicalLooseInput">
            Loose bottles
            <input id="chemicalLooseInput" type="number" name="loose_units" min="0" step="1" value="0" required>
        </label>
        <div class="pack-format-card" data-pack-format aria-live="polite"></div>
        <div class="calculation-card" data-calculation="chemical" aria-live="polite"></div>
    `;
    const selectedProduct = products.some((product) => product.name === values.product)
        ? values.product
        : products[0]?.name;
    setDynamicControlValue("product", selectedProduct || "");
    updateChemicalPackOptions(values.containers_per_box);
    setDynamicControlValue("boxes", values.boxes ?? "0");
    setDynamicControlValue("loose_units", values.loose_units ?? "0");
    updateSpecializedCalculation();
}

function updateChemicalPackOptions(preferredValue) {
    const productControl = dynamicItemFields.querySelector('[name="product"]');
    const perBoxControl = dynamicItemFields.querySelector('[name="containers_per_box"]');
    const product = getChemicalProduct(productControl?.value);
    if (!product || !perBoxControl) {
        return;
    }
    const previousValue = Number(preferredValue || perBoxControl.value);
    perBoxControl.innerHTML = product.containersPerBox
        .map((count) => `<option value="${count}">${count} bottles</option>`)
        .join("");
    perBoxControl.value = product.containersPerBox.includes(previousValue)
        ? String(previousValue)
        : String(product.containersPerBox[0]);
    const packCard = dynamicItemFields.querySelector("[data-pack-format]");
    const displayUnit = product.unit === "ltr" ? "L" : product.unit;
    packCard.innerHTML = `<strong>${product.displayFormat}</strong><span>${formatQuantity(product.packSize)} ${displayUnit} ${product.containerType} × ${perBoxControl.value} ${product.containerType}s/box</span>`;
}

function renderCTCPPlatesForm(values = {}) {
    dynamicItemFields.innerHTML = `
        <div class="specialized-form-heading">
            <strong>CTCP Plates</strong>
            <span>50 sheets = 1 box</span>
        </div>
        <label for="ctcpThicknessSelect">
            Thickness
            <select id="ctcpThicknessSelect" name="thickness" required>
                <option value="0.30">0.30</option>
                <option value="0.20">0.20</option>
            </select>
        </label>
        <label for="ctcpSizeSelect">
            Size
            <select id="ctcpSizeSelect" name="size" required></select>
        </label>
        <label for="ctcpBoxesInput">
            Quantity in boxes
            <input id="ctcpBoxesInput" type="number" name="boxes" min="0" step="1" placeholder="7" required>
        </label>
        <input type="hidden" name="unit" value="box">
        <div class="calculation-card" data-calculation="ctcp" aria-live="polite"></div>
    `;
    setDynamicControlValue("thickness", values.thickness || "0.30");
    updateCTCPSizeOptions(values.size);
    setDynamicControlValue("boxes", values.boxes || "");
    updateSpecializedCalculation();
}

function updateCTCPSizeOptions(preferredSize) {
    const thicknessControl = dynamicItemFields.querySelector('[name="thickness"]');
    const sizeControl = dynamicItemFields.querySelector('[name="size"]');
    if (!thicknessControl || !sizeControl) {
        return;
    }
    const sizes = CTCP_PLATE_SIZES[thicknessControl.value] || [];
    const currentSize = preferredSize || sizeControl.value;
    sizeControl.innerHTML = sizes.map((size) => `<option value="${size}">${size}</option>`).join("");
    if (sizes.includes(currentSize)) {
        sizeControl.value = currentSize;
    }
}

function normalizeM3ZStorageType(value) {
    const normalized = String(value || "").trim().toLowerCase().replace(/-/g, "_");
    if (["cut_piece", "cut_pieces", "cutpiece", "cutpieces", "sheet", "sheets"].includes(normalized)) {
        return M3Z_CUT_PIECE_STORAGE_TYPE;
    }
    return M3Z_ROLL_STORAGE_TYPE;
}

function getM3ZThicknessOption(values = {}) {
    const micronValue = Number(values.thickness_micron);
    if (M3Z_THICKNESS_OPTIONS.some((option) => option.micron === micronValue)) {
        return M3Z_THICKNESS_OPTIONS.find((option) => option.micron === micronValue);
    }
    const thickness = Number(values.thickness);
    const thicknessUnit = String(values.thickness_unit || "mm").toLowerCase();
    const micron = thicknessUnit === "micron" ? thickness : thickness * 1000;
    return M3Z_THICKNESS_OPTIONS.find((option) => Math.abs(option.micron - micron) < 0.000001)
        || M3Z_THICKNESS_OPTIONS[1];
}

function updateM3ZThicknessHidden() {
    const thicknessControl = dynamicItemFields.querySelector('[name="thickness_micron"]');
    const thicknessHidden = dynamicItemFields.querySelector('[name="thickness"]');
    const option = M3Z_THICKNESS_OPTIONS.find((entry) => entry.micron === Number(thicknessControl?.value));
    if (thicknessHidden && option) {
        thicknessHidden.value = formatInputNumber(option.mm);
    }
}

function renderM3ZModeFields(values = {}) {
    const modeFields = dynamicItemFields.querySelector("[data-m3z-mode-fields]");
    const storageControl = dynamicItemFields.querySelector('[name="storage_type"]');
    if (!modeFields || !storageControl) {
        return;
    }
    const storageType = normalizeM3ZStorageType(storageControl.value);
    const thicknessOption = getM3ZThicknessOption(values);
    modeFields.innerHTML = `
        <label for="m3zThicknessSelect">
            Thickness
            <select id="m3zThicknessSelect" name="thickness_micron" required>
                ${M3Z_THICKNESS_OPTIONS.map((option) => `<option value="${option.micron}">${option.label}</option>`).join("")}
            </select>
            <small class="field-hint">Stored as ${thicknessOption.micron} micron; equivalent mm is retained.</small>
        </label>
        <label for="m3zWidthInput">
            Width
            <input id="m3zWidthInput" type="number" name="width" min="0" step="0.001" placeholder="1320" required>
        </label>
        <label for="m3zWidthUnitSelect">
            Width Unit
            <select id="m3zWidthUnitSelect" name="width_unit" required>
                ${ROLL_WIDTH_UNITS.map((unit) => `<option value="${unit}">${unit}</option>`).join("")}
            </select>
        </label>
        <label for="m3zLengthInput">
            Length
            <input id="m3zLengthInput" type="number" name="length" min="0" step="0.001" placeholder="100" required>
        </label>
        <label for="m3zLengthUnitSelect">
            Length Unit
            <select id="m3zLengthUnitSelect" name="length_unit" required>
                ${ROLL_LENGTH_UNITS.map((unit) => `<option value="${unit}">${unit}</option>`).join("")}
            </select>
        </label>
        ${storageType === M3Z_ROLL_STORAGE_TYPE
            ? `<label for="m3zRollsInput">Number of Rolls<input id="m3zRollsInput" type="number" name="number_of_rolls" min="1" step="1" placeholder="8" required></label>
               <input type="hidden" name="unit" value="${ROLL_PAPER_STOCK_UNIT}">`
            : `<label for="m3zSheetsInput">Number of Sheets<input id="m3zSheetsInput" type="number" name="number_of_sheets" min="1" step="1" placeholder="50" required></label>
               <input type="hidden" name="unit" value="sheets">`}
        <input type="hidden" name="thickness" value="${thicknessOption.mm}">
        <input type="hidden" name="thickness_unit" value="mm">
    `;
    setDynamicControlValue("thickness_micron", thicknessOption.micron);
    setDynamicControlValue("width", values.width ?? "");
    setDynamicControlValue("width_unit", values.width_unit || "mm");
    setDynamicControlValue("length", values.length ?? values.height ?? "");
    setDynamicControlValue("length_unit", values.length_unit || "m");
    if (storageType === M3Z_ROLL_STORAGE_TYPE) {
        setDynamicControlValue("number_of_rolls", values.number_of_rolls ?? values.rolls ?? "");
    } else {
        setDynamicControlValue("number_of_sheets", values.number_of_sheets ?? values.sheets ?? "");
    }
    ["width_unit", "length_unit"].forEach((name) => {
        const control = dynamicItemFields.querySelector(`[name="${name}"]`);
        if (control) {
            control.dataset.previousUnit = control.value;
        }
    });
    updateSpecializedCalculation();
}

function renderCalibratedUnderpackingPaperForm(values = {}) {
    const storageType = normalizeM3ZStorageType(values.storage_type);
    dynamicItemFields.innerHTML = `
        <div class="specialized-form-heading">
            <strong>Calibrated Underpacking Paper (M3Z)</strong>
            <span>Choose how this stock is stored</span>
        </div>
        <label for="m3zStorageTypeSelect">
            Storage Type
            <select id="m3zStorageTypeSelect" name="storage_type" required>
                <option value="roll">As Rolls (stored in Sq.m.)</option>
                <option value="cut_piece">As Cut Pieces (stored in Sheets)</option>
            </select>
        </label>
        <div class="dynamic-item-fields" data-m3z-mode-fields></div>
        <div class="calculation-card roll-paper-calculation" data-calculation="m3z" aria-live="polite"></div>
        <p class="field-hint">Roll mode stores width × length × rolls in m². Cut Pieces mode stores the number of sheets; dimensions show the reference area per sheet.</p>
    `;
    setDynamicControlValue("storage_type", storageType);
    renderM3ZModeFields(values);
}

function renderGenericAddItemForm(category, values = {}) {
    const rule = getCategoryRule(category);
    const unit = rule.unitOptions.includes(values.unit) ? values.unit : rule.defaultUnit;
    const selectedIsRoll = isRollItem(unit);
    const fields = [];
    if (rule.requiresBrand) {
        fields.push('<label for="formBrandInput">Brand<input id="formBrandInput" type="text" name="brand" placeholder="Day" required></label>');
    }
    if (rule.requiresType) {
        fields.push('<label for="formTypeInput">Type<input id="formTypeInput" type="text" name="type" placeholder="UV" required></label>');
    }
    fields.push(`<label for="formUnitSelect">Unit<select id="formUnitSelect" name="unit" required>${rule.unitOptions.map((option) => `<option value="${option}">${option}</option>`).join("")}</select></label>`);
    if (requiresBatchRollNo(category, unit)) {
        fields.push('<label for="formBatchRollNoInput">Batch / Roll No.<input id="formBatchRollNoInput" type="text" name="batch_roll_no" placeholder="BR-001" required></label>');
    }
    if (rule.usesDimensions) {
        fields.push('<label for="formWidthInput">Width<input id="formWidthInput" type="text" name="width" placeholder="1040" required></label>');
        fields.push(`<label for="formHeightInput">${selectedIsRoll ? "Length (mtr)" : "Length"}<input id="formHeightInput" type="text" name="height" placeholder="${selectedIsRoll ? "30" : "920"}" required></label>`);
    }
    if (rule.requiresThickness) {
        fields.push(`<label for="formThicknessInput">Thickness<input id="formThicknessInput" type="text" name="thickness" placeholder="Enter thickness in ${rule.thicknessUnit}" required><small class="field-hint">Thickness unit: ${rule.thicknessUnit}.</small></label>`);
    }
    const quantityStep = selectedIsRoll ? "0.0001" : (rule.quantityAllowsDecimal ? "0.01" : "1");
    const quantityHint = selectedIsRoll
        ? "Calculated in sq.m from width and length."
        : (rule.quantityAllowsDecimal ? "Decimals are allowed for this category." : "Use the item's stock unit.");
    fields.push(`<label for="formQuantityInput">Quantity<input id="formQuantityInput" type="number" name="quantity" min="0" step="${quantityStep}" ${selectedIsRoll ? "readonly" : ""} required><small class="field-hint">${quantityHint}</small></label>`);
    dynamicItemFields.innerHTML = fields.join("");
    ["brand", "type", "batch_roll_no", "width", "height", "thickness", "quantity"].forEach((name) => setDynamicControlValue(name, values[name] || ""));
    setDynamicControlValue("unit", unit);
    updateRollQuantityEstimate();
}

function updateRollQuantityEstimate() {
    const unitControl = dynamicItemFields.querySelector('[name="unit"]');
    const widthControl = dynamicItemFields.querySelector('[name="width"]');
    const heightControl = dynamicItemFields.querySelector('[name="height"]');
    const quantityControl = dynamicItemFields.querySelector('[name="quantity"]');
    if (!unitControl || !isRollItem(unitControl.value) || !widthControl || !heightControl || !quantityControl) {
        return;
    }
    const area = getRollAreaSqm(widthControl.value, heightControl.value);
    quantityControl.value = area === null ? "" : String(roundStockQuantity(area));
}

function updateSpecializedCalculation() {
    const categoryKey = getCategoryKey(formCategorySelect.value);
    const category = getCategoryLabel(categoryKey);
    const calculation = dynamicItemFields.querySelector("[data-calculation]");
    if (!calculation) {
        return;
    }
    if (categoryKey === "creasing_matrix") {
        const packets = Math.max(0, Math.floor(Number(dynamicItemFields.querySelector('[name="quantity"]')?.value || 0)));
        calculation.innerHTML = `<span>${formatQuantity(packets)} packets</span><strong>${Math.floor(packets / 10)} boxes · ${packets % 10} loose packets</strong>`;
        return;
    }
   if (categoryKey === "ctcp_plates") {
       const boxes = Math.max(0, Math.floor(Number(dynamicItemFields.querySelector('[name="boxes"]')?.value || 0)));
       calculation.innerHTML = `<span>${formatQuantity(boxes)} boxes</span><strong>${formatQuantity(boxes * 50)} total sheets</strong>`;
       return;
    }
    if (categoryKey === "calibrated_underpacking_paper") {
        const result = getM3ZCalculation();
        const widthText = result.widthMeters === null ? "-" : `${result.widthMeters.toFixed(3)} m`;
        const lengthText = result.lengthMeters === null ? "-" : `${result.lengthMeters.toFixed(3)} m`;
        const thicknessText = result.thicknessMicron === null ? "-" : `${formatFixedQuantity(result.thicknessMicron, 0)} micron`;
        if (result.storageType === M3Z_CUT_PIECE_STORAGE_TYPE) {
            const areaText = result.areaPerSheet === null ? "-" : `${formatFixedQuantity(result.areaPerSheet, 4)} m²`;
            const sheetsText = result.count === null ? "-" : formatQuantity(result.count);
            calculation.innerHTML = `
                <span>Width: ${widthText}</span>
                <span>Thickness: ${thicknessText}</span>
                <span>Length: ${lengthText}</span>
                <span>Area / Sheet (reference): ${areaText}</span>
                <strong>Total Stock: ${sheetsText} Sheets</strong>
            `;
        } else {
            const areaText = result.areaPerRoll === null ? "-" : `${formatFixedQuantity(result.areaPerRoll)} m²`;
            const totalText = result.totalArea === null ? "-" : `${formatFixedQuantity(result.totalArea)} m²`;
            const rollsText = result.count === null ? "-" : formatQuantity(result.count);
            calculation.innerHTML = `
                <span>Width: ${widthText}</span>
                <span>Thickness: ${thicknessText}</span>
                <span>Length: ${lengthText}</span>
                <span>Area / Roll: ${areaText}</span>
                <span>Number of Rolls: ${rollsText}</span>
                <strong>Total Stock: ${totalText}</strong>
            `;
        }
        return;
    }
    if (categoryKey === "rubber_blankets") {
        const result = getRubberBlanketCalculation();
        const widthText = result.actualWidthMeters === null ? "-" : `${result.actualWidthMeters.toFixed(3)} m`;
        const lengthText = result.lengthMeters === null ? "-" : `${result.lengthMeters.toFixed(3)} m`;
        if (result.storageType === M3Z_CUT_PIECE_STORAGE_TYPE) {
            const areaText = result.areaPerSheet === null ? "-" : `${formatFixedQuantity(result.areaPerSheet, 4)} Sq.m`;
            const sheetsText = result.count === null ? "-" : formatQuantity(result.count);
            calculation.innerHTML = `
                <strong>Calculation (Stored in Sheets)</strong>
                <span>Area / Sheet (for reference): ${areaText}</span>
                <span>No. of Sheets: ${sheetsText}</span>
                <strong>Total Stock: ${sheetsText} Sheets</strong>
            `;
        } else {
            const areaText = result.areaPerRoll === null ? "-" : `${formatFixedQuantity(result.areaPerRoll)} Sq.m`;
            const totalText = result.totalArea === null ? "-" : `${formatFixedQuantity(result.totalArea)} Sq.m`;
            const rollsText = result.count === null ? "-" : formatQuantity(result.count);
            calculation.innerHTML = `
                <strong>Calculation (Stored in Sq.m.)</strong>
                <span>Actual Width: ${widthText}</span>
                <span>Length: ${lengthText}</span>
                <span>Area / Roll: ${areaText}</span>
                <span>No. of Rolls: ${rollsText}</span>
                <strong>Total Stock: ${totalText}</strong>
            `;
        }
        return;
    }
   if (CHEMICAL_CATEGORIES.has(category)) {
        const product = getChemicalProduct(dynamicItemFields.querySelector('[name="product"]')?.value);
        const perBox = Number(dynamicItemFields.querySelector('[name="containers_per_box"]')?.value || 0);
        const boxes = Math.max(0, Math.floor(Number(dynamicItemFields.querySelector('[name="boxes"]')?.value || 0)));
        const loose = Math.max(0, Math.floor(Number(dynamicItemFields.querySelector('[name="loose_units"]')?.value || 0)));
        const containers = boxes * perBox + loose;
        const total = product ? roundStockQuantity(containers * product.packSize) : 0;
        const unit = product?.unit === "ltr" ? "L" : product?.unit || "";
        calculation.innerHTML = `<span>${formatQuantity(containers)} ${product?.containerType || "container"}${containers === 1 ? "" : "s"}</span><strong>${formatQuantity(total)} ${unit} total</strong>`;
    }
}

function updateCategoryDrivenFields(values = {}) {
    renderAddItemForm(getCategoryKey(formCategorySelect.value), values);
}

function buildParams() {
    const params = new URLSearchParams();

    if (searchInput.value.trim()) {
        params.set("search", searchInput.value.trim());
    }
    if (categoryFilter.value) {
        params.set("category", categoryFilter.value);
    }
    if (brandFilter.value) {
        params.set("brand", brandFilter.value);
    }
    if (thicknessFilter.value.trim()) {
        params.set("thickness", thicknessFilter.value.trim());
    }
    if (typeFilter.value) {
        params.set("type", typeFilter.value);
    }
    if (lowStockOnly.checked) {
        params.set("low_stock", "true");
        params.set("low_stock_threshold", lowStockThreshold.value || "5");
    }

    return params;
}

function getItemKey(item) {
    if (getSpecializedKind(item) === "rubber_blankets") {
        return [
            item.category,
            item.blanket_name || item.brand,
            item.storage_type || M3Z_ROLL_STORAGE_TYPE,
            item.nominal_width || item.width,
            item.actual_width,
            item.length ?? item.height,
            item.length_unit,
            item.thickness,
            item.roll_no || "-",
            item.batch_no || "-",
            item.print_type || "-",
        ].join("|");
    }
    if (getSpecializedKind(item) === "calibrated_underpacking_paper") {
        return [
            item.category,
            item.storage_type || M3Z_ROLL_STORAGE_TYPE,
            item.width_meters ?? convertRollDimension(item.width, item.width_unit, "m") ?? item.width,
            item.length_meters ?? convertRollDimension(item.length ?? item.height, item.length_unit, "m") ?? item.length ?? item.height,
            item.thickness_micron ?? getRollPaperThicknessMicron(item),
        ].join("|");
    }
    return [
        item.category,
        item.brand,
        item.type,
        item.batch_roll_no || "-",
        item.width || "-",
        item.height || "-",
        item.thickness || "-",
    ].join("|");
}

function getLookupPayload(item) {
    return {
        category: item.category,
        brand: item.brand,
        type: item.type,
        batch_roll_no: item.batch_roll_no || "",
        unit: item.unit,
        width: item.width,
        height: item.height,
        width_unit: item.width_unit,
        length: item.length ?? item.height,
        length_unit: item.length_unit,
        thickness: item.thickness,
        thickness_unit: item.thickness_unit,
        thickness_micron: item.thickness_micron,
        storage_type: item.storage_type || M3Z_ROLL_STORAGE_TYPE,
        number_of_rolls: item.number_of_rolls,
        number_of_sheets: item.number_of_sheets,
        blanket_name: item.blanket_name,
        nominal_width: item.nominal_width ?? item.width,
        actual_width: item.actual_width,
        roll_no: item.roll_no || "",
        batch_no: item.batch_no || "",
        print_type: item.print_type || "",
    };
}

function isPlaceholderValue(value) {
    if (typeof value !== "string") {
        return false;
    }
    const normalized = value.trim().toLowerCase();
    return normalized === "__none__" || normalized === "none" || normalized === "_none";
}

function getDisplayValue(value) {
    if (value == null) {
        return "";
    }
    if (typeof value !== "string") {
        return String(value);
    }
    const trimmed = value.trim();
    return isPlaceholderValue(trimmed) ? "" : trimmed;
}

function joinPathParts(parts) {
    return parts.map((part) => getDisplayValue(part)).filter(Boolean).join(" / ");
}

function promptForReason(title, promptText) {
    if (!reasonDialog || typeof reasonDialog.showModal !== "function") {
        const fallbackReason = window.prompt(promptText);
        return Promise.resolve(fallbackReason && fallbackReason.trim() ? fallbackReason.trim() : null);
    }

    reasonTitle.textContent = title;
    reasonPrompt.textContent = promptText;
    reasonInput.value = "";
    reasonError.textContent = "";

    return new Promise((resolve) => {
        const cleanup = () => {
            reasonForm.removeEventListener("submit", handleSubmit);
            reasonCancelButton.removeEventListener("click", handleCancel);
            reasonDialog.removeEventListener("cancel", handleCancel);
        };

        const closeDialog = (value) => {
            cleanup();
            reasonDialog.close();
            resolve(value);
        };

        const handleSubmit = (event) => {
            event.preventDefault();
            const reason = reasonInput.value.trim();
            if (!reason) {
                reasonError.textContent = "Reason is required";
                reasonError.dataset.tone = "error";
                reasonInput.focus();
                return;
            }
            closeDialog(reason);
        };

        const handleCancel = (event) => {
            event.preventDefault();
            closeDialog(null);
        };

        reasonForm.addEventListener("submit", handleSubmit);
        reasonCancelButton.addEventListener("click", handleCancel);
        reasonDialog.addEventListener("cancel", handleCancel);
        reasonDialog.showModal();
        reasonInput.focus();
    });
}

function getLogDetailsMarkup(details) {
    if (!details || Object.keys(details).length === 0) {
        return "<p>No extra details recorded.</p>";
    }

    return Object.entries(details).map(([key, value]) => {
        const label = key.replace(/_/g, " ");
        const text = typeof value === "object" && value !== null
            ? JSON.stringify(value)
            : String(value ?? "");
        return `<p><strong>${label}:</strong> ${text}</p>`;
    }).join("");
}

function getLogItemPreview(log) {
    return joinPathParts([log.category, log.brand, log.type, log.batch_roll_no, log.size]) || "item";
}

function renderLogEntry(log) {
    const itemPreview = getLogItemPreview(log);
    return `
        <details class="log-entry">
            <summary>
                <span class="log-summary-text"><strong>${log.action}</strong> ${itemPreview} via ${log.source}</span>
                <span class="log-time">${formatIstDateTime(log.changed_at)} IST</span>
            </summary>
            <p><strong>Item:</strong> ${itemPreview}</p>
            <p><strong>Quantity:</strong> ${formatQuantity(log.quantity_before)} -> ${formatQuantity(log.quantity_after)} ${getDisplayUnit(log.unit)}</p>
            <p><strong>Reason:</strong> ${log.reason || "Not recorded"}</p>
            <div class="log-details">${getLogDetailsMarkup(log.details)}</div>
        </details>
    `;
}

async function request(path, options = {}) {
    const response = await fetch(`${API_BASE_URL}${path}`, {
        ...options,
        credentials: API_CREDENTIALS_MODE,
        headers: {
            ...(options.body instanceof FormData ? {} : { "Content-Type": "application/json" }),
            ...(options.headers || {}),
        },
    });

    if (options.expectBlob) {
        if (!response.ok) {
            const data = await response.json().catch(() => ({}));
            throw new Error(data.error || "Unable to export Excel");
        }
        return response.blob();
    }

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        const error = new Error(data.error || "Request failed");
        error.status = response.status;
        throw error;
    }

    return data;
}

async function checkAuthSession() {
    try {
        const response = await request("/auth/me");
        setAuthenticatedUser(response.user);
        return Boolean(response.user);
    } catch (error) {
        setAuthenticatedUser(null);
        return false;
    }
}

async function handleLogin(event) {
    event.preventDefault();
    const formData = new FormData(loginForm);

    try {
        const response = await request("/auth/login", {
            method: "POST",
            body: JSON.stringify({
                email: String(formData.get("email") || "").trim(),
                password: String(formData.get("password") || ""),
            }),
        });
        loginForm.reset();
        setMessage(loginMessage, "Login successful", "success");
        setAuthenticatedUser(response.user);
        await initializeAppData();
    } catch (error) {
        setMessage(loginMessage, error.message, "error");
    }
}

async function handleSignup(event) {
    event.preventDefault();
    const formData = new FormData(signupForm);

    try {
        const response = await request("/auth/signup", {
            method: "POST",
            body: JSON.stringify({
                email: String(formData.get("email") || "").trim(),
                password: String(formData.get("password") || ""),
            }),
        });
        signupForm.reset();
        setMessage(signupMessage, "Signup successful", "success");
        setAuthenticatedUser(response.user);
        await initializeAppData();
    } catch (error) {
        setMessage(signupMessage, error.message, "error");
    }
}

async function handleForgotPassword(event) {
    event.preventDefault();
    const formData = new FormData(forgotForm);

    try {
        const response = await request("/auth/forgot-password", {
            method: "POST",
            body: JSON.stringify({
                email: String(formData.get("email") || "").trim(),
                new_password: String(formData.get("new_password") || ""),
            }),
        });
        forgotForm.reset();
        setMessage(forgotMessage, "Password reset successful", "success");
        setAuthenticatedUser(response.user);
        await initializeAppData();
    } catch (error) {
        setMessage(forgotMessage, error.message, "error");
    }
}

async function handleLogout() {
    try {
        await request("/auth/logout", { method: "POST" });
    } catch (error) {
        // Logout should still clear the local session state.
    }
    setAuthenticatedUser(null);
    setAuthTab("login");
}

function handleAdminPanelClick() {
    window.location.hash = "#admin";
}

function populateCategoryFilterOptions(selectedValue = categoryFilter.value) {
    categoryFilter.innerHTML = '<option value="">All categories</option>' +
        CATEGORY_OPTIONS.map((option) => `<option value="${option.label}">${option.code} - ${option.label}</option>`).join("");
    if (CATEGORY_OPTIONS.some((option) => option.label === selectedValue)) {
        categoryFilter.value = selectedValue;
    }
}

function populateFormCategoryOptions(selectedValue = formCategorySelect.value) {
    const selectedCategoryKey = getCategoryKey(selectedValue);
    formCategorySelect.innerHTML = `
        <option value="">Select category</option>
        ${CATEGORY_OPTIONS.map((option) => `<option value="${option.key}">${option.code} - ${option.label}</option>`).join("")}
    `;
    if (CATEGORY_OPTIONS.some((option) => option.key === selectedCategoryKey)) {
        formCategorySelect.value = selectedCategoryKey;
    }
}

function populateSelectOptions(items) {
    const uniqueBrands = [...new Set(items.map((item) => getDisplayValue(item.brand)).filter(Boolean))].sort();
    const uniqueTypes = [...new Set(items.map((item) => getDisplayValue(item.type)).filter(Boolean))].sort();

    const selectedCategory = categoryFilter.value;
    const selectedBrand = brandFilter.value;
    const selectedType = typeFilter.value;

    populateCategoryFilterOptions(selectedCategory);
    brandFilter.innerHTML = '<option value="">All brands</option>' +
        uniqueBrands.map((value) => `<option value="${value}">${value}</option>`).join("");
    typeFilter.innerHTML = '<option value="">All types</option>' +
        uniqueTypes.map((value) => `<option value="${value}">${value}</option>`).join("");

    if (CATEGORY_OPTIONS.some((option) => option.label === selectedCategory)) {
        categoryFilter.value = selectedCategory;
    }
    if (uniqueBrands.includes(selectedBrand)) {
        brandFilter.value = selectedBrand;
    }
    if (uniqueTypes.includes(selectedType)) {
        typeFilter.value = selectedType;
    }

    const selectedFormValues = getDynamicFormValues();
    populateFormCategoryOptions(formCategorySelect.value);
    updateCategoryDrivenFields(selectedFormValues);
}

function renderCategories(items) {
    const counts = items.reduce((map, item) => {
        map[item.category] = (map[item.category] || 0) + 1;
        return map;
    }, {});

    const categories = Object.keys(counts).sort();
    categoryCount.textContent = String(categories.length);
    sidebarCategoryCount.textContent = String(categories.length);
    overviewCategoryCount.textContent = String(categories.length);

    if (categories.length === 0) {
        categoryList.innerHTML = '<li class="empty-state">No categories yet</li>';
        return;
    }

    categoryList.innerHTML = categories
        .map((category) => `<li><span>${category}</span><strong>${counts[category]}</strong></li>`)
        .join("");
}

function buildTree(items) {
    return items.reduce((tree, item) => {
        const brand = getDisplayValue(item.brand);
        const itemType = getDisplayValue(item.type);
        tree[item.category] ??= {};
        tree[item.category][brand] ??= {};
        tree[item.category][brand][itemType] ??= [];
        tree[item.category][brand][itemType].push(item);
        return tree;
    }, {});
}

function renderTree(items) {
    if (items.length === 0) {
        treeView.innerHTML = '<p class="empty-state">No inventory hierarchy available</p>';
        return;
    }

    const tree = buildTree(items);
    const categoryMarkup = Object.keys(tree).sort().map((category) => {
        const brands = tree[category];
        const brandMarkup = Object.keys(brands).sort().map((brand) => {
            const types = brands[brand];
            const typeMarkup = Object.keys(types).sort().map((itemType) => {
                const sizes = types[itemType]
                    .sort((a, b) => a.size.localeCompare(b.size))
                    .map((item) => {
                        const details = [item.size];
                        const batchRollNo = getDisplayValue(item.batch_roll_no);
                        if (batchRollNo) {
                            details.push(`Batch/Roll No: ${batchRollNo}`);
                        }
                        return `<div class="tree-leaf">${details.join(" • ")} - ${getSpecializedStockSummary(item)}</div>`;
                    })
                    .join("");

                if (!itemType) {
                    return sizes;
                }

                return `
                    <details>
                        <summary>${itemType}</summary>
                        <div class="tree-children">${sizes}</div>
                    </details>
                `;
            }).join("");

            if (!brand) {
                return typeMarkup;
            }

            return `
                <details>
                    <summary>${brand}</summary>
                    <div class="tree-children">${typeMarkup}</div>
                </details>
            `;
        }).join("");

        return `
            <details open>
                <summary>${category}</summary>
                <div class="tree-children">${brandMarkup}</div>
            </details>
        `;
    }).join("");

    treeView.innerHTML = categoryMarkup;
}

function renderInventoryTree(items) {
    if (items.length === 0) {
        treeView.innerHTML = '<p class="empty-state">No inventory hierarchy available</p>';
        return;
    }

    const tree = buildTree(items);
    const categoryMarkup = Object.keys(tree).sort().map((category) => {
        const brands = tree[category];
        const brandMarkup = Object.keys(brands).sort().map((brand) => {
            const types = brands[brand];
            const typeMarkup = Object.keys(types).sort().map((itemType) => {
                const sizes = types[itemType]
                    .sort((a, b) => a.size.localeCompare(b.size))
                    .map((item) => {
                        const details = [item.size];
                        const batchRollNo = getDisplayValue(item.batch_roll_no);
                        if (batchRollNo) {
                            details.push(`Batch/Roll No: ${batchRollNo}`);
                        }
                        return `<div class="tree-leaf">${details.join(" / ")} - ${getSpecializedStockSummary(item)}</div>`;
                    })
                    .join("");

                if (!itemType) {
                    return sizes;
                }

                return `
                    <details>
                        <summary>${itemType}</summary>
                        <div class="tree-children">${sizes}</div>
                    </details>
                `;
            }).join("");

            if (!brand) {
                return typeMarkup;
            }

            return `
                <details>
                    <summary>${brand}</summary>
                    <div class="tree-children">${typeMarkup}</div>
                </details>
            `;
        }).join("");

        return `
            <details open>
                <summary>${category}</summary>
                <div class="tree-children">${brandMarkup}</div>
            </details>
        `;
    }).join("");

    treeView.innerHTML = categoryMarkup;
}

function renderOverviewStats(items) {
    const lowStockThresholdValue = Number(lowStockThreshold.value || "5");
    const lowStockCount = items.filter((item) => item.quantity <= lowStockThresholdValue).length;

    overviewItemCount.textContent = String(items.length);
    sidebarItemCount.textContent = String(items.length);
    overviewLowStockCount.textContent = String(lowStockCount);
}

function getInventoryDisplayMode(items) {
    if (items.length && items.every((item) => item.category === RUBBER_BLANKET_CATEGORY)) {
        const storageTypes = new Set(items.map((item) => item.storage_type || M3Z_ROLL_STORAGE_TYPE));
        if (storageTypes.size === 1) {
            return [...storageTypes][0] === M3Z_CUT_PIECE_STORAGE_TYPE
                ? "rubber_blankets_cut_piece"
                : "rubber_blankets_roll";
        }
        return "rubber_blankets_mixed";
    }
    if (items.length && items.every((item) => getSpecializedKind(item) === "calibrated_underpacking_paper")) {
        const storageTypes = new Set(items.map((item) => item.storage_type || M3Z_ROLL_STORAGE_TYPE));
        if (storageTypes.size === 1) {
            return [...storageTypes][0] === M3Z_CUT_PIECE_STORAGE_TYPE
                ? "calibrated_underpacking_paper_cut_piece"
                : "calibrated_underpacking_paper_roll";
        }
        return "calibrated_underpacking_paper_mixed";
    }
    const kinds = new Set(items.map((item) => getSpecializedKind(item)));
    return kinds.size === 1 ? [...kinds][0] : "generic";
}

function getInventoryColumnsOriginal(items, interactive = true) {
    const mode = getInventoryDisplayMode(items);
    let columns;
    if (mode === "rubber_blankets") {
        columns = [
            { key: "blanket", label: "Blanket Name", value: (item) => item.blanket_name || getDisplayValue(item.brand) || "—" },
            { key: "thickness", label: "Thickness", value: (item) => item.thickness ? `${item.thickness} ${item.thickness_unit || "mm"}` : "—" },
            { key: "width", label: "Width", value: (item) => `${item.nominal_width ?? item.width ?? "—"} mm` },
            { key: "actual_width", label: "Actual Width", value: (item) => item.actual_width ? `${item.actual_width} mm actual` : "—" },
            { key: "length", label: "Length", value: (item) => item.length ?? item.height ? `${item.length ?? item.height} ${item.length_unit || "m"}` : "—" },
            { key: "roll_no", label: "Roll No", value: (item) => getDisplayValue(item.roll_no) || "—" },
            { key: "batch_no", label: "Batch No", value: (item) => getDisplayValue(item.batch_no) || "—" },
            { key: "print_type", label: "Print Type", value: (item) => item.print_type === "P" ? "Printed (P)" : (item.print_type === "W/O" ? "Without Print (W/O)" : "—") },
            { key: "area", label: "Area/Roll", value: (item) => {
                const area = getStockBreakdown(item)?.area_per_roll_sqm;
                return area ? `${formatFixedQuantity(area)} m²` : "—";
            }, align: "right" },
            { key: "stock", label: "Stock", value: (item) => `${formatFixedQuantity(item.quantity)} m²`, align: "right" },
        ];
    } else if (mode === "calibrated_underpacking_paper_roll") {
        columns = [
            { key: "item", label: "Item", value: () => ROLL_PAPER_CATEGORY },
            { key: "thickness", label: "Thickness", value: (item) => formatQuantity(getRollPaperThicknessMicron(item)) + " micron" },
            { key: "width", label: "Width", value: (item) => (item.width || "-") + " " + (item.width_unit || "") },
            { key: "length", label: "Length", value: (item) => (item.length ?? item.height ?? "-") + " " + (item.length_unit || "") },
            { key: "rolls", label: "Rolls", value: (item) => formatQuantity(getStockBreakdown(item)?.rolls ?? item.number_of_rolls ?? 0) + " rolls", align: "right" },
            { key: "area", label: "Area/Roll", value: (item) => {
                const area = getRollPaperAreaPerRoll(item);
                return area === null ? "-" : formatQuantity(area) + " m²";
            }, align: "right" },
            { key: "total", label: "Total Stock", value: (item) => formatQuantity(item.quantity) + " m²", align: "right" },
        ];
    } else if (mode === "calibrated_underpacking_paper_cut_piece") {
        columns = [
            { key: "item", label: "Item", value: () => ROLL_PAPER_CATEGORY },
            { key: "thickness", label: "Thickness", value: (item) => formatQuantity(getRollPaperThicknessMicron(item)) + " micron" },
            { key: "width", label: "Width", value: (item) => (item.width || "-") + " " + (item.width_unit || "") },
            { key: "length", label: "Length", value: (item) => (item.length ?? item.height ?? "-") + " " + (item.length_unit || "") },
            { key: "sheets", label: "Sheets", value: (item) => formatQuantity(getStockBreakdown(item)?.sheets ?? item.number_of_sheets ?? item.quantity) + " sheets", align: "right" },
            { key: "area", label: "Area/Sheet", value: (item) => {
                const area = getStockBreakdown(item)?.area_per_sheet_sqm ?? getRollPaperAreaPerRoll(item);
                return area === null || area === undefined ? "-" : formatQuantity(area) + " m²";
            }, align: "right" },
            { key: "total", label: "Total Stock", value: (item) => formatQuantity(item.quantity) + " sheets", align: "right" },
        ];
    } else if (mode === "calibrated_underpacking_paper_mixed") {
        columns = [
            { key: "item", label: "Item", value: () => ROLL_PAPER_CATEGORY },
            { key: "storage", label: "Storage Type", value: (item) => (item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE ? "Cut Pieces" : "Rolls" },
            { key: "thickness", label: "Thickness", value: (item) => formatQuantity(getRollPaperThicknessMicron(item)) + " micron" },
            { key: "width", label: "Width", value: (item) => (item.width || "-") + " " + (item.width_unit || "") },
            { key: "length", label: "Length", value: (item) => (item.length ?? item.height ?? "-") + " " + (item.length_unit || "") },
            { key: "stock", label: "Stock", value: (item) => (item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE ? `${formatQuantity(item.quantity)} sheets` : `${formatQuantity(item.quantity)} m²`, align: "right" },
        ];
    } else if (mode === "creasing_matrix") {
        columns = [
            { key: "item", label: "Item", value: () => "Creasing Matrix" },
            { key: "size", label: "Size", value: formatItemSize },
            { key: "thickness", label: "Thickness", value: (item) => `${item.thickness} mm`, align: "right" },
            { key: "packets", label: "Qty (Pkt)", value: (item) => formatQuantity(getStockBreakdown(item)?.packets ?? item.quantity), align: "right" },
            { key: "boxes", label: "Boxes", value: (item) => formatQuantity(getStockBreakdown(item)?.boxes || 0), align: "right" },
            { key: "loose", label: "Loose Pkts", value: (item) => formatQuantity(getStockBreakdown(item)?.loose_units || 0), align: "right" },
        ];
    } else if (mode === "chemical") {
        columns = [
            { key: "item", label: "Item", value: (item) => getDisplayValue(item.brand) },
            { key: "format", label: "Format", value: (item) => getItemPackaging(item)?.display_format || getDisplayValue(item.type) },
            { key: "containers", label: "Qty (Bottles/Cans)", value: (item) => formatQuantity(getStockBreakdown(item)?.containers || 0), align: "right" },
            { key: "total", label: "Total Ltr/Kg", value: (item) => `${formatQuantity(item.quantity)} ${formatChemicalUnit(item.unit)}`, align: "right" },
            { key: "packaging", label: "Packaging", value: (item) => {
                const packaging = getItemPackaging(item);
                return `${packaging.containers_per_box} ${packaging.container_type}s/box`;
            } },
        ];
    } else if (mode === "ctcp_plates") {
        columns = [
            { key: "item", label: "Item", value: () => "CTCP Plates" },
            { key: "size", label: "Size", value: formatItemSize },
            { key: "thickness", label: "Thickness", value: (item) => item.thickness, align: "right" },
            { key: "boxes", label: "Boxes", value: (item) => formatQuantity(getStockBreakdown(item)?.boxes ?? item.quantity), align: "right" },
            { key: "sheets", label: "Total Sheets", value: (item) => formatQuantity(getStockBreakdown(item)?.total_sheets || 0), align: "right" },
        ];
    } else {
        columns = [
            { key: "category", label: "Category", value: (item) => item.category },
            { key: "brand", label: "Brand / Product", value: (item) => getDisplayValue(item.brand) || "-" },
            { key: "type", label: "Type / Format", value: (item) => getDisplayValue(item.type) || "-" },
            { key: "batch", label: "Batch / Roll No.", value: (item) => getDisplayValue(item.batch_roll_no) || "-" },
            { key: "size", label: "Size", value: formatItemSize },
            { key: "thickness", label: "Thickness", value: (item) => item.thickness || "-", align: "right" },
            { key: "quantity", label: "Quantity", value: (item) => formatQuantity(item.quantity), align: "right" },
            { key: "unit", label: "Unit", value: (item) => getDisplayUnit(item.unit) },
            { key: "stock_details", label: "Packaging", value: (item) => getSpecializedKind(item) === "generic" ? "-" : getSpecializedStockSummary(item) },
        ];
    }
    if (interactive && !isReadOnlyUser()) {
        columns.push({ key: "movement", label: "Stock Movement", action: "movement" });
        columns.push({ key: "delete", label: "Delete", action: "delete" });
    }
    return columns;
}

function getInventoryColumns(items, interactive = true) {
    const mode = getInventoryDisplayMode(items);
    const unitSqm = "m" + String.fromCharCode(178);
    const columns = [];
    const blanketName = (item) => item.blanket_name || getDisplayValue(item.brand) || "-";
    const thickness = (item) => item.thickness ? `${item.thickness} ${item.thickness_unit || "mm"}` : "-";
    const width = (item) => `${item.nominal_width ?? item.width ?? "-"} mm`;
    const length = (item) => item.length ?? item.height ? `${item.length ?? item.height} ${item.length_unit || "m"}` : "-";
    const printType = (item) => item.print_type === "P" ? "Printed (P)" : (item.print_type === "W/O" ? "Without Print (W/O)" : "-");
    const storageType = (item) => (item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE ? "Cut Piece" : "Roll";

    if (mode === "rubber_blankets_roll") {
        columns.push(
            { key: "blanket", label: "Blanket Name", value: blanketName },
            { key: "storage", label: "Storage Type", value: () => "Roll" },
            { key: "thickness", label: "Thickness", value: thickness },
            { key: "width", label: "Width", value: width },
            { key: "actual_width", label: "Actual Width", value: (item) => item.actual_width ? `${item.actual_width} mm actual` : "-" },
            { key: "length", label: "Length", value: length },
            { key: "print_type", label: "Print Type", value: printType },
            { key: "rolls", label: "No. of Rolls", value: (item) => `${formatQuantity(getStockBreakdown(item)?.rolls ?? item.number_of_rolls ?? 0)} Rolls`, align: "right" },
            { key: "area", label: "Area/Roll", value: (item) => {
                const area = getStockBreakdown(item)?.area_per_roll_sqm;
                return area ? `${formatFixedQuantity(area)} ${unitSqm}` : "-";
            }, align: "right" },
            { key: "stock", label: "Total Stock", value: (item) => `${formatFixedQuantity(item.quantity)} ${unitSqm}`, align: "right" },
        );
    } else if (mode === "rubber_blankets_cut_piece") {
        columns.push(
            { key: "blanket", label: "Blanket Name", value: blanketName },
            { key: "storage", label: "Storage Type", value: () => "Cut Piece" },
            { key: "thickness", label: "Thickness", value: thickness },
            { key: "width", label: "Width", value: width },
            { key: "length", label: "Length", value: length },
            { key: "print_type", label: "Print Type", value: printType },
            { key: "sheets", label: "No. of Sheets", value: (item) => `${formatQuantity(getStockBreakdown(item)?.sheets ?? item.number_of_sheets ?? item.quantity)} Sheets`, align: "right" },
            { key: "area", label: "Area/Sheet (Reference)", value: (item) => {
                const area = getStockBreakdown(item)?.area_per_sheet_sqm;
                return area ? `${formatFixedQuantity(area, 4)} ${unitSqm} reference` : "-";
            }, align: "right" },
            { key: "stock", label: "Total Stock", value: (item) => `${formatQuantity(item.quantity)} Sheets`, align: "right" },
        );
    } else if (mode === "rubber_blankets_mixed") {
        columns.push(
            { key: "blanket", label: "Blanket Name", value: blanketName },
            { key: "storage", label: "Storage Type", value: storageType },
            { key: "thickness", label: "Thickness", value: thickness },
            { key: "width", label: "Width", value: width },
            { key: "length", label: "Length", value: length },
            { key: "print_type", label: "Print Type", value: printType },
            { key: "stock", label: "Total Stock", value: (item) => storageType(item) === "Cut Piece" ? `${formatQuantity(item.quantity)} Sheets` : `${formatFixedQuantity(item.quantity)} ${unitSqm}`, align: "right" },
        );
    } else {
        return getInventoryColumnsLegacy(items, interactive);
    }
    if (interactive && !isReadOnlyUser()) {
        columns.push({ key: "movement", label: "Stock Movement", action: "movement" });
        columns.push({ key: "delete", label: "Delete", action: "delete" });
    }
    return columns;
}

function getInventoryColumnsLegacy(items, interactive = true) {
    return getInventoryColumnsOriginal(items, interactive);
    /* Kept as a readable fallback reference for the specialized column contract. */
    const mode = getInventoryDisplayMode(items);
    let columns;
    if (mode === "calibrated_underpacking_paper_roll") {
        columns = [
            { key: "item", label: "Item", value: () => ROLL_PAPER_CATEGORY },
            { key: "thickness", label: "Thickness", value: (item) => formatQuantity(getRollPaperThicknessMicron(item)) + " micron" },
            { key: "width", label: "Width", value: (item) => (item.width || "-") + " " + (item.width_unit || "") },
            { key: "length", label: "Length", value: (item) => (item.length ?? item.height ?? "-") + " " + (item.length_unit || "") },
            { key: "rolls", label: "Rolls", value: (item) => formatQuantity(getStockBreakdown(item)?.rolls ?? item.number_of_rolls ?? 0) + " rolls", align: "right" },
            { key: "area", label: "Area/Roll", value: (item) => { const area = getRollPaperAreaPerRoll(item); return area === null ? "-" : formatQuantity(area) + " m\\u00B2"; }, align: "right" },
            { key: "total", label: "Total Stock", value: (item) => formatQuantity(item.quantity) + " m\\u00B2", align: "right" },
        ];
    } else if (mode === "calibrated_underpacking_paper_cut_piece") {
        columns = [
            { key: "item", label: "Item", value: () => ROLL_PAPER_CATEGORY },
            { key: "thickness", label: "Thickness", value: (item) => formatQuantity(getRollPaperThicknessMicron(item)) + " micron" },
            { key: "width", label: "Width", value: (item) => (item.width || "-") + " " + (item.width_unit || "") },
            { key: "length", label: "Length", value: (item) => (item.length ?? item.height ?? "-") + " " + (item.length_unit || "") },
            { key: "sheets", label: "Sheets", value: (item) => formatQuantity(getStockBreakdown(item)?.sheets ?? item.number_of_sheets ?? item.quantity) + " sheets", align: "right" },
            { key: "area", label: "Area/Sheet", value: (item) => { const area = getStockBreakdown(item)?.area_per_sheet_sqm ?? getRollPaperAreaPerRoll(item); return area == null ? "-" : formatQuantity(area) + " m\\u00B2"; }, align: "right" },
            { key: "total", label: "Total Stock", value: (item) => formatQuantity(item.quantity) + " sheets", align: "right" },
        ];
    } else {
        columns = [
            { key: "category", label: "Category", value: (item) => item.category },
            { key: "brand", label: "Brand / Product", value: (item) => getDisplayValue(item.brand) || "-" },
            { key: "type", label: "Type / Format", value: (item) => getDisplayValue(item.type) || "-" },
            { key: "batch", label: "Batch / Roll No.", value: (item) => getDisplayValue(item.batch_roll_no) || "-" },
            { key: "size", label: "Size", value: formatItemSize },
            { key: "thickness", label: "Thickness", value: (item) => item.thickness || "-", align: "right" },
            { key: "quantity", label: "Quantity", value: (item) => formatQuantity(item.quantity), align: "right" },
            { key: "unit", label: "Unit", value: (item) => getDisplayUnit(item.unit) },
            { key: "stock_details", label: "Packaging", value: (item) => getSpecializedKind(item) === "generic" ? "-" : getSpecializedStockSummary(item) },
        ];
    }
    if (interactive && !isReadOnlyUser()) {
        columns.push({ key: "movement", label: "Stock Movement", action: "movement" });
        columns.push({ key: "delete", label: "Delete", action: "delete" });
    }
    return columns;
}

function createMovementControls(item) {
    const wrapper = document.createElement("div");
    wrapper.className = "update-controls";
    const kind = getSpecializedKind(item);
    const packaging = getItemPackaging(item);
    wrapper.innerHTML = `
        <select class="movement-select" aria-label="Stock direction">
            <option value="in">Stock In</option>
            <option value="out">Stock Out</option>
        </select>
    `;
    if (kind === "calibrated_underpacking_paper") {
        if ((item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
            wrapper.insertAdjacentHTML("beforeend", '<label class="movement-roll-field">Sheets<input class="movement-sheets-input" type="number" value="1" min="1" step="1"></label>');
        } else {
            wrapper.insertAdjacentHTML("beforeend", '<label class="movement-roll-field">Rolls<input class="movement-rolls-input" type="number" value="1" min="1" step="1"></label>');
        }
    } else if (kind === "creasing_matrix" || kind === "chemical") {
        const looseLabel = kind === "chemical" ? `${packaging.container_type}s` : "packets";
        wrapper.insertAdjacentHTML("beforeend", `
            <div class="packaging-movement-inputs">
                <label>Boxes<input class="movement-boxes-input" type="number" value="0" min="0" step="1"></label>
                <label>Loose ${looseLabel}<input class="movement-loose-input" type="number" value="0" min="0" step="1"></label>
            </div>
        `);
    } else if (kind === "ctcp_plates") {
        wrapper.insertAdjacentHTML("beforeend", '<label class="movement-box-field">Boxes<input class="movement-boxes-input" type="number" value="1" min="1" step="1"></label>');
    } else {
        const movementOptions = isRollItem(item)
            ? '<option value="sqm">sq.m</option><option value="mtr">mtr</option><option value="inch">inch</option>'
            : `<option value="item">${getDisplayUnit(item.unit)}</option>`;
        const step = isRollItem(item) ? "0.0001" : (getCategoryRule(item.category).quantityAllowsDecimal ? "0.01" : "1");
        wrapper.insertAdjacentHTML("beforeend", `
            <div class="movement-amount">
                <input class="delta-input" type="number" value="1" min="${step}" step="${step}">
                <select class="movement-unit-select" aria-label="Movement unit">${movementOptions}</select>
            </div>
        `);
    }
    wrapper.insertAdjacentHTML("beforeend", '<button class="secondary-button update-button" type="button">Apply</button>');
    return wrapper;
}

function createMovementControls(item) {
    const wrapper = document.createElement("div");
    wrapper.className = "update-controls";
    const kind = getSpecializedKind(item);
    const packaging = getItemPackaging(item);
    wrapper.innerHTML = `
        <select class="movement-select" aria-label="Stock direction">
            <option value="in">Stock In</option>
            <option value="out">Stock Out</option>
        </select>
    `;
    if (kind === "rubber_blankets") {
        if ((item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
            wrapper.insertAdjacentHTML("beforeend", '<label class="movement-roll-field">Sheets<input class="rubber-movement-sheets-input" type="number" value="1" min="1" step="1"></label>');
        } else {
            wrapper.insertAdjacentHTML("beforeend", `
                <label class="movement-roll-field">Mode
                    <select class="rubber-movement-mode">
                        <option value="rolls">Full Rolls</option>
                        <option value="area">Partial Sq.m.</option>
                    </select>
                </label>
                <label class="movement-roll-field">Amount<input class="rubber-movement-input" type="number" value="1" min="0.0001" step="1"></label>
            `);
        }
    } else if (kind === "calibrated_underpacking_paper") {
        if ((item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
            wrapper.insertAdjacentHTML("beforeend", '<label class="movement-roll-field">Sheets<input class="movement-sheets-input" type="number" value="1" min="1" step="1"></label>');
        } else {
            wrapper.insertAdjacentHTML("beforeend", '<label class="movement-roll-field">Rolls<input class="movement-rolls-input" type="number" value="1" min="1" step="1"></label>');
        }
    } else if (kind === "creasing_matrix" || kind === "chemical") {
        const looseLabel = kind === "chemical" ? `${packaging.container_type}s` : "packets";
        wrapper.insertAdjacentHTML("beforeend", `
            <div class="packaging-movement-inputs">
                <label>Boxes<input class="movement-boxes-input" type="number" value="0" min="0" step="1"></label>
                <label>Loose ${looseLabel}<input class="movement-loose-input" type="number" value="0" min="0" step="1"></label>
            </div>
        `);
    } else if (kind === "ctcp_plates") {
        wrapper.insertAdjacentHTML("beforeend", '<label class="movement-box-field">Boxes<input class="movement-boxes-input" type="number" value="1" min="1" step="1"></label>');
    } else {
        const movementOptions = isRollItem(item)
            ? '<option value="sqm">sq.m</option><option value="mtr">mtr</option><option value="inch">inch</option>'
            : `<option value="item">${getDisplayUnit(item.unit)}</option>`;
        const step = isRollItem(item) ? "0.0001" : (getCategoryRule(item.category).quantityAllowsDecimal ? "0.01" : "1");
        wrapper.insertAdjacentHTML("beforeend", `
            <div class="movement-amount">
                <input class="delta-input" type="number" value="1" min="${step}" step="${step}">
                <select class="movement-unit-select" aria-label="Movement unit">${movementOptions}</select>
            </div>
        `);
    }
    wrapper.insertAdjacentHTML("beforeend", '<button class="secondary-button update-button" type="button">Apply</button>');
    return wrapper;
}

function appendInventoryRow(target, item, columns) {
    const row = document.createElement("tr");
    row.dataset.itemKey = getItemKey(item);
    columns.forEach((column) => {
        const cell = document.createElement("td");
        if (column.align) {
            cell.dataset.align = column.align;
        }
        if (column.action === "movement") {
            cell.appendChild(createMovementControls(item));
        } else if (column.action === "delete") {
            cell.dataset.align = "center";
            cell.innerHTML = '<button class="danger-button delete-button" type="button">Delete</button>';
        } else {
            cell.textContent = column.value(item);
        }
        row.appendChild(cell);
    });
    target.appendChild(row);
}

function renderTable(items) {
    statusText.textContent = `${items.length} item(s) found`;
    const columns = getInventoryColumns(items, true);
    inventoryTableHead.innerHTML = `<tr>${columns.map((column) => `<th${column.align ? ` data-align="${column.align}"` : ""}>${column.label}</th>`).join("")}</tr>`;
    inventoryTableBody.innerHTML = "";
    if (items.length === 0) {
        inventoryTableBody.innerHTML = `<tr><td colspan="${columns.length}" class="empty-state">No inventory available</td></tr>`;
        return;
    }
    items.forEach((item) => appendInventoryRow(inventoryTableBody, item, columns));
}

function renderLogs(logs) {
    if (logs.length === 0) {
        sidebarLogCount.textContent = "0";
        overviewLogBadge.textContent = "0";
        logsList.innerHTML = '<p class="empty-state">No stock history yet</p>';
        overviewLogs.innerHTML = '<p class="empty-state">No stock history yet</p>';
        return;
    }

    sidebarLogCount.textContent = String(logs.length);
    overviewLogBadge.textContent = String(logs.length);

    const markup = logs.map((log) => renderLogEntry(log)).join("");

    logsList.innerHTML = markup;
    overviewLogs.innerHTML = logs.slice(0, 4).map((log) => renderLogEntry(log)).join("");
}

function renderUsers(users) {
    if (!userHasRole("admin")) {
        usersTableBody.innerHTML = '<tr><td colspan="5" class="empty-state">Users are available only for admins.</td></tr>';
        return;
    }

    if (users.length === 0) {
        usersTableBody.innerHTML = '<tr><td colspan="5" class="empty-state">No users found</td></tr>';
        return;
    }

    usersTableBody.innerHTML = "";

    users.forEach((user) => {
        const row = document.createElement("tr");
        const roleOptions = ["user", "workshop", "admin"]
            .map((role) => `<option value="${role}" ${user.role === role ? "selected" : ""}>${role}</option>`)
            .join("");
        const removeDisabled = user.is_superadmin || user.id === state.user?.id ? "disabled" : "";
        const roleDisabled = user.is_superadmin ? "disabled" : "";

        row.innerHTML = `
            <td>${user.email}${user.is_superadmin ? " (superadmin)" : ""}</td>
            <td>
                <select class="user-role-select" data-user-id="${user.id}" ${roleDisabled}>
                    ${roleOptions}
                </select>
            </td>
            <td>${formatIstDateTime(user.created_at) || "-"}</td>
            <td>${formatIstDateTime(user.updated_at) || "-"}</td>
            <td data-align="center">
                <button class="danger-button remove-user-button" type="button" data-user-id="${user.id}" ${removeDisabled}>Remove</button>
            </td>
        `;
        usersTableBody.appendChild(row);
    });
}

async function loadInventory() {
    if (!state.user) {
        return;
    }
    setMessage(statusText, "Loading inventory...");

    try {
        const params = buildParams();
        const path = params.toString() ? `/inventory?${params.toString()}` : "/inventory";
        const items = await request(path);
        state.inventory = Array.isArray(items) ? items : [];

        renderCategories(state.inventory);
        renderInventoryTree(state.inventory);
        renderOverviewStats(state.inventory);
        renderTable(state.inventory);
        populateSelectOptions(state.inventory);
        setMessage(statusText, `${state.inventory.length} item(s) found`);
        window.dispatchEvent(new CustomEvent("onlystock:inventory-updated", { detail: { items: state.inventory } }));
    } catch (error) {
        state.inventory = [];
        renderCategories([]);
        renderInventoryTree([]);
        renderOverviewStats([]);
        renderTable([]);
        setMessage(statusText, error.message, "error");
        window.dispatchEvent(new CustomEvent("onlystock:inventory-updated", { detail: { items: [] } }));
    }
}

async function loadLogs() {
    if (!userHasRole("admin")) {
        state.logs = [];
        sidebarLogCount.textContent = "0";
        overviewLogBadge.textContent = "0";
        logsList.innerHTML = '<p class="empty-state">Stock logs are available only for admins.</p>';
        overviewLogs.innerHTML = '<p class="empty-state">Recent activity is available only for admins.</p>';
        return;
    }
    try {
        const logs = await request("/stock-logs?limit=12");
        state.logs = Array.isArray(logs) ? logs : [];
        renderLogs(state.logs);
    } catch (error) {
        logsList.innerHTML = `<p class="empty-state">${error.message}</p>`;
        overviewLogs.innerHTML = `<p class="empty-state">${error.message}</p>`;
    }
}

async function loadUsers() {
    if (!userHasRole("admin")) {
        state.users = [];
        renderUsers([]);
        return;
    }

    try {
        const users = await request("/admin/users");
        state.users = Array.isArray(users) ? users : [];
        renderUsers(state.users);
        setMessage(adminUsersMessage, `${state.users.length} user(s) found`);
    } catch (error) {
        state.users = [];
        renderUsers([]);
        setMessage(adminUsersMessage, error.message, "error");
    }
}

function parseNonNegativeWholeNumber(value, label) {
    const parsed = Number(value);
    if (!Number.isInteger(parsed) || parsed < 0) {
        throw new Error(`${label} must be a non-negative whole number`);
    }
    return parsed;
}

function parsePositiveFormNumber(value, label) {
    const text = String(value ?? "").trim();
    const parsed = Number(text);
    if (!text || !Number.isFinite(parsed) || parsed <= 0) {
        throw new Error(`${label} must be greater than 0`);
    }
    return parsed;
}

function parsePositiveFormWholeNumber(value, label) {
    const parsed = parsePositiveFormNumber(value, label);
    if (!Number.isInteger(parsed)) {
        throw new Error(`${label} must be a whole number`);
    }
    return parsed;
}

function splitSelectedSize(size) {
    const parts = String(size || "").split(/\s+X\s+/i);
    return parts.length === 2 ? parts : ["", ""];
}

function buildAddItemPayload() {
    updateRollQuantityEstimate();
    const values = getDynamicFormValues();
    const categoryKey = getCategoryKey(formCategorySelect.value);
    const category = getCategoryLabel(categoryKey);
    if (!categoryKey) {
        throw new Error("category is required");
    }

    if (categoryKey === "rubber_blankets") {
        const storageType = normalizeM3ZStorageType(values.storage_type);
        const blanket = getRubberBlanketRule(values.blanket_name);
        if (!blanket) {
            throw new Error("select a valid blanket name");
        }
        const thickness = parsePositiveFormNumber(values.thickness, "thickness");
        if (!blanket.thicknessOptions.includes(thickness)) {
            throw new Error("select a valid thickness for this blanket");
        }
        const selectedWidth = getSelectedRubberBlanketWidth();
        if (!selectedWidth) {
            throw new Error("select a valid width for this blanket");
        }
        const validWidth = getRubberBlanketWidths(blanket, thickness).some(
            ([nominal, actual]) => nominal === selectedWidth.nominal && actual === selectedWidth.actual
        );
        if (!validWidth) {
            throw new Error("selected width is not valid for this blanket and thickness");
        }
        const length = parsePositiveFormNumber(values.length, "length");
        const lengthUnit = String(values.length_unit || "").trim().toLowerCase();
        if (!ROLL_LENGTH_UNITS.includes(lengthUnit)) {
            throw new Error("length unit must be m, mm, or inch");
        }
        const widthUnit = String(values.width_unit || "").trim().toLowerCase();
        if (!ROLL_WIDTH_UNITS.includes(widthUnit)) {
            throw new Error("width unit must be mm, m, or inch");
        }
        const printType = String(values.print_type || "").trim();
        if (blanket.printTypes.length && !blanket.printTypes.includes(printType)) {
            throw new Error("select Printed (P) or Without Print (W/O)");
        }
        if (!blanket.printTypes.length && printType) {
            throw new Error("print type is not applicable to this blanket");
        }
        const lengthMeters = convertRollDimension(length, lengthUnit, "m");
        const areaPerPiece = selectedWidth.actual / 1000 * lengthMeters;
        const rolls = storageType === M3Z_ROLL_STORAGE_TYPE
            ? parsePositiveFormWholeNumber(values.number_of_rolls, "number of rolls")
            : null;
        const sheets = storageType === M3Z_CUT_PIECE_STORAGE_TYPE
            ? parsePositiveFormWholeNumber(values.number_of_sheets, "number of sheets")
            : null;
        return {
            category,
            blanket_name: blanket.name,
            brand: blanket.name,
            type: printType,
            storage_type: storageType,
            thickness,
            thickness_unit: "mm",
            width: selectedWidth.nominal,
            nominal_width: selectedWidth.nominal,
            actual_width: selectedWidth.actual,
            width_unit: "mm",
            height: length,
            length,
            length_unit: lengthUnit,
            print_type: printType,
            number_of_rolls: rolls,
            number_of_sheets: sheets,
            rolls,
            sheets,
            area_per_roll_sqm: storageType === M3Z_ROLL_STORAGE_TYPE ? roundStockQuantity(areaPerPiece) : null,
            area_per_sheet_sqm: roundStockQuantity(areaPerPiece),
            quantity: storageType === M3Z_ROLL_STORAGE_TYPE
                ? roundStockQuantity(areaPerPiece * rolls)
                : sheets,
            unit: storageType === M3Z_ROLL_STORAGE_TYPE ? RUBBER_BLANKET_STOCK_UNIT : "sheets",
        };
    }

    if (categoryKey === "calibrated_underpacking_paper") {
        const storageType = normalizeM3ZStorageType(values.storage_type);
        const thicknessMicron = parsePositiveFormWholeNumber(values.thickness_micron, "thickness");
        const width = parsePositiveFormNumber(values.width, "width");
        const widthUnit = String(values.width_unit || "").trim().toLowerCase();
        const length = parsePositiveFormNumber(values.length, "length");
        const lengthUnit = String(values.length_unit || "").trim().toLowerCase();
        const thicknessOption = M3Z_THICKNESS_OPTIONS.find((option) => option.micron === thicknessMicron);
        if (!thicknessOption) {
            throw new Error("select a valid thickness");
        }
        if (!ROLL_WIDTH_UNITS.includes(widthUnit)) {
            throw new Error("width unit must be mm, m, or inch");
        }
        if (!ROLL_LENGTH_UNITS.includes(lengthUnit)) {
            throw new Error("length unit must be m, mm, or inch");
        }
        const widthMeters = convertRollDimension(width, widthUnit, "m");
        const lengthMeters = convertRollDimension(length, lengthUnit, "m");
        const areaPerSheet = widthMeters * lengthMeters;
        const rolls = storageType === M3Z_ROLL_STORAGE_TYPE
            ? parsePositiveFormWholeNumber(values.number_of_rolls, "number of rolls")
            : null;
        const sheets = storageType === M3Z_CUT_PIECE_STORAGE_TYPE
            ? parsePositiveFormWholeNumber(values.number_of_sheets, "number of sheets")
            : null;
        const quantity = storageType === M3Z_ROLL_STORAGE_TYPE
            ? roundStockQuantity(areaPerSheet * rolls)
            : sheets;
        return {
            category,
            brand: "",
            type: "",
            storage_type: storageType,
            width,
            width_unit: widthUnit,
            height: length,
            length,
            length_unit: lengthUnit,
            thickness: thicknessOption.mm,
            thickness_unit: "mm",
            thickness_micron: thicknessMicron,
            number_of_rolls: rolls,
            number_of_sheets: sheets,
            rolls,
            sheets,
            area_per_roll_sqm: areaPerSheet,
            area_per_sheet_sqm: areaPerSheet,
            quantity,
            unit: storageType === M3Z_ROLL_STORAGE_TYPE ? ROLL_PAPER_STOCK_UNIT : "sheets",
        };
    }

    if (categoryKey === "creasing_matrix") {
        const thickness = String(values.thickness || "");
        const size = String(values.size || "");
        const quantity = parseNonNegativeWholeNumber(values.quantity, "quantity in packets");
        if (!CREASING_MATRIX_SIZES[thickness]?.includes(size)) {
            throw new Error("select a valid size for the chosen Creasing Matrix thickness");
        }
        const [width, height] = splitSelectedSize(size);
        return { category, brand: "", type: "", thickness, size, width, height, quantity, unit: "pkt" };
    }

    if (categoryKey === "ctcp_plates") {
        const thickness = String(values.thickness || "");
        const size = String(values.size || "");
        const boxes = parseNonNegativeWholeNumber(values.boxes, "boxes");
        if (!CTCP_PLATE_SIZES[thickness]?.includes(size)) {
            throw new Error("select a valid size for the chosen CTCP thickness");
        }
        const [width, height] = splitSelectedSize(size);
        return {
            category,
            brand: "",
            type: "",
            thickness,
            size,
            width,
            height,
            boxes,
            quantity: boxes,
            total_sheets: boxes * 50,
            unit: "box",
        };
    }

    if (CHEMICAL_CATEGORIES.has(category)) {
        const product = getChemicalProduct(values.product);
        if (!product || product.category !== category) {
            throw new Error("select a valid chemical product");
        }
        const containersPerBox = parseNonNegativeWholeNumber(values.containers_per_box, "containers per box");
        if (!product.containersPerBox.includes(containersPerBox)) {
            throw new Error("select a valid bottles-per-box configuration");
        }
        const boxes = parseNonNegativeWholeNumber(values.boxes, "boxes");
        const looseUnits = parseNonNegativeWholeNumber(values.loose_units, "loose bottles");
        if (looseUnits >= containersPerBox) {
            throw new Error(`loose bottles must be less than ${containersPerBox}`);
        }
        const containers = boxes * containersPerBox + looseUnits;
        const quantity = roundStockQuantity(containers * product.packSize);
        return {
            category,
            product: product.name,
            brand: product.name,
            type: `${product.packSize}${product.unit}`,
            pack_size: product.packSize,
            container_type: product.containerType,
            containers_per_box: containersPerBox,
            boxes,
            loose_units: looseUnits,
            containers,
            quantity,
            unit: product.unit,
        };
    }

    const rule = getCategoryRule(category);
    const unit = String(values.unit || "").trim();
    if (!unit) {
        throw new Error("unit is required");
    }
    if (rule.requiresBrand && !String(values.brand || "").trim()) {
        throw new Error("brand is required");
    }
    if (rule.requiresType && !String(values.type || "").trim()) {
        throw new Error("type is required");
    }
    const quantity = Number(values.quantity);
    if (!Number.isFinite(quantity) || quantity < 0) {
        throw new Error("quantity must be a non-negative number");
    }
    if (!isRollItem(unit) && !rule.quantityAllowsDecimal && !Number.isInteger(quantity)) {
        throw new Error("quantity must be a non-negative whole number");
    }
    if (rule.unitLinkedToType) {
        const normalizedType = String(values.type || "").trim().toLowerCase();
        if (!rule.typeOptions?.includes(normalizedType) || unit.toLowerCase() !== normalizedType) {
            throw new Error("type and unit must both be coil or pkt");
        }
    }
    if (rule.usesDimensions && (!String(values.width || "").trim() || !String(values.height || "").trim())) {
        throw new Error("width and length are required for this category");
    }
    if (rule.requiresThickness && !String(values.thickness || "").trim()) {
        throw new Error(`thickness is required in ${rule.thicknessUnit}`);
    }
    if (requiresBatchRollNo(category, unit) && !String(values.batch_roll_no || "").trim()) {
        throw new Error("batch / roll no. is required for blanket rolls");
    }
    return {
        category,
        brand: String(values.brand || "").trim(),
        type: String(values.type || "").trim(),
        batch_roll_no: String(values.batch_roll_no || "").trim(),
        width: String(values.width || "").trim(),
        height: String(values.height || "").trim(),
        thickness: String(values.thickness || "").trim(),
        quantity,
        unit,
    };
}

async function handleAddItem(event) {
    event.preventDefault();
    let payload;
    try {
        payload = buildAddItemPayload();
    } catch (error) {
        setMessage(formMessage, error.message, "error");
        return;
    }

    try {
        await request("/add-item", { method: "POST", body: JSON.stringify(payload) });
        itemForm.reset();
        renderAddItemForm("");
        setMessage(formMessage, "Item added successfully", "success");
        if (window.location.hash !== "#inventory") {
            window.location.hash = "#inventory";
        }
        await Promise.all([loadInventory(), loadLogs()]);
    } catch (error) {
        setMessage(formMessage, error.message, "error");
    }
}

function getExcelMode() {
    return excelForm?.querySelector('input[name="excelMode"]:checked')?.value || "import";
}

function escapeExcelHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatExcelCell(value) {
    if (value === null || value === undefined || value === "") {
        return "—";
    }
    if (typeof value === "number") {
        return formatQuantity(value);
    }
    return escapeExcelHtml(value);
}

function resetExcelPreview() {
    state.excelPreview = null;
    [excelPreviewSection, excelValidationSection, excelDetailSection, excelApplySection].forEach((section) => {
        section?.classList.add("is-hidden");
    });
    excelErrorsPanel?.classList.add("is-hidden");
    if (excelApplyButton) {
        excelApplyButton.disabled = true;
    }
}

function renderExcelSheetSummary(sheets) {
    excelSheetSummary.innerHTML = sheets.map((sheet) => {
        const tone = sheet.error_count ? "error" : sheet.row_count ? "valid" : "empty";
        return `
            <article class="excel-sheet-card excel-sheet-card--${tone}" data-sheet-key="${escapeExcelHtml(sheet.key)}">
                <span class="excel-sheet-name">${escapeExcelHtml(sheet.sheet_name)}</span>
                <strong>${sheet.row_count}</strong>
                <small>${sheet.row_count === 1 ? "row" : "rows"}</small>
                <span class="excel-sheet-status">${sheet.error_count ? `${sheet.error_count} error(s)` : sheet.row_count ? "Valid" : "Empty"}</span>
            </article>
        `;
    }).join("");
}

function renderExcelErrors(validation) {
    const issues = [...(validation.errors || [])];
    if (!issues.length) {
        excelErrorsPanel.innerHTML = '<p class="empty-state">No workbook errors found.</p>';
        return;
    }
    excelErrorsPanel.innerHTML = issues.map((issue) => `
        <article class="excel-error-row">
            <strong>${escapeExcelHtml(issue.sheet || "Workbook")}${issue.row ? ` · row ${issue.row}` : ""}</strong>
            <span>${escapeExcelHtml(issue.field || "Row")}</span>
            <p>${escapeExcelHtml(issue.problem || "Invalid value")}</p>
            ${issue.suggestion ? `<small>${escapeExcelHtml(issue.suggestion)}</small>` : ""}
        </article>
    `).join("");
}

function renderExcelCategoryPreview(sheets) {
    excelCategoryPreview.innerHTML = sheets.map((sheet, index) => {
        const visibleRows = sheet.rows.slice(0, 100);
        const headerMarkup = sheet.columns.map((column) => `<th>${escapeExcelHtml(column)}</th>`).join("");
        const rowMarkup = visibleRows.map((row) => `
            <tr class="excel-preview-row excel-preview-row--${row.status}">
                ${sheet.columns.map((column) => `<td>${formatExcelCell(row.values[column])}</td>`).join("")}
            </tr>
        `).join("");
        const emptyMarkup = sheet.row_count === 0
            ? `<p class="empty-state">No populated rows in this sheet.</p>`
            : `<div class="table-wrap"><table><thead><tr>${headerMarkup}</tr></thead><tbody>${rowMarkup}</tbody></table></div>`;
        return `
            <details class="excel-category-accordion" ${index === 0 && sheet.row_count ? "open" : ""}>
                <summary>
                    <span>${escapeExcelHtml(sheet.sheet_name)}</span>
                    <span class="excel-row-badge">${sheet.row_count} row${sheet.row_count === 1 ? "" : "s"}</span>
                </summary>
                ${sheet.sheet_errors?.length ? `<div class="excel-inline-error">${escapeExcelHtml(sheet.sheet_errors[0].problem)}</div>` : ""}
                ${emptyMarkup}
                ${sheet.row_count > visibleRows.length ? `<p class="field-hint">Showing the first ${visibleRows.length} rows.</p>` : ""}
            </details>
        `;
    }).join("");
}

function renderExcelPreview(preview) {
    state.excelPreview = preview;
    [excelPreviewSection, excelValidationSection, excelDetailSection, excelApplySection].forEach((section) => {
        section?.classList.remove("is-hidden");
    });
    excelPreviewTotal.textContent = `${preview.total_rows} row${preview.total_rows === 1 ? "" : "s"}`;
    renderExcelSheetSummary(preview.sheets);

    const warnings = preview.validation.warnings || [];
    excelWorkbookWarnings.innerHTML = warnings.map((warning) => `
        <div class="excel-warning-strip">
            <strong>${escapeExcelHtml(warning.sheet || "Workbook")}</strong>
            <span>${escapeExcelHtml(warning.problem)}</span>
        </div>
    `).join("");

    excelValidCount.textContent = preview.validation.valid_rows;
    excelWarningCount.textContent = preview.validation.warning_count;
    excelErrorCount.textContent = preview.validation.error_rows;
    excelErrorsButton.disabled = preview.validation.error_rows === 0;
    renderExcelErrors(preview.validation);
    renderExcelCategoryPreview(preview.sheets);

    const summary = preview.update_summary;
    excelAddCount.textContent = summary.add;
    excelUpdateCount.textContent = summary.update;
    excelDeleteCount.textContent = summary.delete;
    excelUnchangedCount.textContent = summary.unchanged;
    const isUpdate = preview.mode === "update";
    excelApplyTitle.textContent = isUpdate ? "Update Summary" : "Import Summary";
    excelApplyDescription.textContent = isUpdate
        ? "This action can delete records missing from the non-empty category sheets listed below."
        : `${preview.validation.valid_rows} validated row(s) will be added or used to update matching items.`;
    excelApplyButton.textContent = isUpdate ? "Confirm Update" : "Import Valid Rows";
    excelApplyButton.classList.toggle("danger-button", isUpdate);
    excelApplyButton.classList.toggle("primary-button", !isUpdate);
    excelApplyButton.disabled = !preview.can_apply;
    excelUpdateScope.innerHTML = isUpdate
        ? `<strong>Update scope:</strong> ${summary.scope.length
            ? summary.scope.map((scope) => `<span>${escapeExcelHtml(scope.label)}</span>`).join("")
            : "<span>No non-empty category sheet is in scope.</span>"}`
        : "";
}

async function handleExcelUpload(event) {
    event.preventDefault();
    const file = excelFileInput.files[0];
    if (!file) {
        setMessage(excelMessage, "Select an Excel workbook to preview", "error");
        return;
    }
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
        setMessage(excelMessage, "Only .xlsx and .xls workbooks are supported", "error");
        return;
    }
    const body = new FormData();
    body.append("file", file);
    body.append("mode", getExcelMode());
    body.append("action", "preview");
    excelPreviewButton.disabled = true;
    setMessage(excelMessage, "Parsing and validating workbook…");
    try {
        const response = await request("/upload-excel", { method: "POST", body });
        renderExcelPreview(response);
        setMessage(
            excelMessage,
            response.can_apply
                ? `Preview ready: ${response.validation.valid_rows} valid row(s)`
                : `Preview found ${response.validation.error_rows} error(s)`,
            response.can_apply ? "success" : "error"
        );
    } catch (error) {
        resetExcelPreview();
        setMessage(excelMessage, error.message, "error");
    } finally {
        excelPreviewButton.disabled = false;
    }
}

async function handleExcelApply() {
    const preview = state.excelPreview;
    const file = excelFileInput.files[0];
    if (!preview || !preview.can_apply || !file) {
        setMessage(excelMessage, "Upload and validate the workbook before applying it", "error");
        return;
    }
    const mode = getExcelMode();
    if (mode !== preview.mode) {
        setMessage(excelMessage, "Operation changed. Upload and preview the workbook again.", "error");
        return;
    }
    if (mode === "update") {
        const confirmed = window.confirm(
            `WARNING: This update may delete ${preview.update_summary.delete} inventory record(s) missing from the provided non-empty category sheets. Continue?`
        );
        if (!confirmed) {
            return;
        }
    }
    const reason = await promptForReason(
        mode === "update" ? "Excel Update Reason" : "Excel Import Reason",
        mode === "update"
            ? "Why are you applying this category-scoped inventory update?"
            : "Why are you importing these inventory rows?"
    );
    if (!reason) {
        return;
    }
    const body = new FormData();
    body.append("file", file);
    body.append("mode", mode);
    body.append("action", "apply");
    body.append("reason", reason);
    if (mode === "update") {
        body.append("confirm_update", "true");
    }
    excelApplyButton.disabled = true;
    try {
        const response = await request("/upload-excel", { method: "POST", body });
        setMessage(
            excelMessage,
            `Excel ${response.mode}: ${response.inserted} inserted, ${response.updated} updated, ${response.deleted} deleted, ${response.unchanged} unchanged`,
            "success"
        );
        excelForm.reset();
        excelUpdateWarning.classList.add("is-hidden");
        resetExcelPreview();
        await Promise.all([loadInventory(), loadLogs()]);
    } catch (error) {
        excelApplyButton.disabled = false;
        setMessage(excelMessage, error.message, "error");
    }
}

async function downloadExcel(path, filename, successMessage) {
    try {
        const blob = await request(path, {
            method: "GET",
            expectBlob: true,
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        link.click();
        URL.revokeObjectURL(url);
        setMessage(excelMessage, successMessage, "success");
    } catch (error) {
        setMessage(excelMessage, error.message, "error");
    }
}

async function handleImportTemplateDownload() {
    await downloadExcel("/download-import-template", "only_stock_import_template.xlsx", "Multi-sheet import template downloaded");
}

async function handleUpdateSheetDownload() {
    await downloadExcel("/export-update-excel", "only_stock_update_workbook.xlsx", "Update workbook downloaded");
}

async function handleCurrentStockDownload() {
    await downloadExcel("/export-excel", "only_stock_current_inventory.xlsx", "Current stock workbook downloaded");
}

function createAdjustmentBatchRow(defaultUnit = "Mtr") {
    const fragment = adjustmentBatchTemplate.content.cloneNode(true);
    const row = fragment.querySelector(".adjustment-batch-row");
    row.querySelector(".adjustment-batch-unit").value = defaultUnit;
    return row;
}

function createAdjustmentItem(defaultUnit = "Mtr") {
    const fragment = adjustmentItemTemplate.content.cloneNode(true);
    const item = fragment.querySelector(".adjustment-item");
    item.querySelector(".adjustment-batches").appendChild(createAdjustmentBatchRow(defaultUnit));
    return item;
}

function addAdjustmentItem(defaultUnit = "Mtr") {
    adjustmentItems.appendChild(createAdjustmentItem(defaultUnit));
}

function normalizeUnitLabel(value) {
    const text = String(value || "").trim().toLowerCase();
    if (text === "mtr" || text === "meter" || text === "meters") {
        return "Mtr";
    }
    if (text === "sq.mtr" || text === "square.mtr" || text === "square mtr" || text === "sqmtr" || text === "sq mtr" || text === "sq.meter" || text === "sq meter") {
        return "sq.mtr";
    }
    if (text === "pcs" || text === "piece" || text === "pieces") {
        return "pcs";
    }
    return "Mtr";
}

function splitManufacturerBatch(rawValue, fallbackUnit = "Mtr") {
    const text = String(rawValue || "").trim();
    const match = text.match(/^(.*?)(?:\s+(Mtr|sq\.mtr|square\.mtr|sqmtr|pcs))?$/i);
    if (!match) {
        return { manufacturerBatch: text, usageUnit: fallbackUnit };
    }
    return {
        manufacturerBatch: match[1].trim(),
        usageUnit: normalizeUnitLabel(match[2] || fallbackUnit),
    };
}

function resolvePastedColumnIndex(headerMap, names) {
    for (const name of names) {
        if (headerMap.has(name)) {
            return headerMap.get(name);
        }
    }
    return -1;
}

function isAdjustmentBatchLabel(value) {
    return /^[A-Za-z]{1,3}\d+\s*-\s*/.test(String(value || "").trim());
}

function isAdjustmentItemLabel(value) {
    const text = String(value || "").trim();
    if (!text || isAdjustmentBatchLabel(text)) {
        return false;
    }
    return /image print|master|bl|mm|inch|"/i.test(text);
}

function parseQuantityWithUnit(value) {
    const text = String(value || "").trim();
    const match = text.match(/(-?\d+(?:\.\d+)?)\s*([A-Za-z.]+)?/);
    if (!match) {
        return { quantity: "", unit: "Mtr" };
    }
    return {
        quantity: match[1] || "",
        unit: normalizeUnitLabel(match[2] || "Mtr"),
    };
}

function normalizeSummaryUnit(value) {
    const text = String(value || "").trim().toLowerCase();
    if (text === "square.mtr" || text === "square mtr" || text === "sq.mtr" || text === "sqmtr" || text === "sq mtr") {
        return "sq.mtr";
    }
    return normalizeUnitLabel(text);
}

function hasAdjustmentUnit(value) {
    return /\b(Mtr|sq\.mtr|square\.mtr|sqmtr|pcs)\b/i.test(String(value || ""));
}

function buildManufacturerBatchValue(value) {
    const text = String(value || "").trim();
    if (!text) {
        return "";
    }
    return hasAdjustmentUnit(text) ? text : `${text} Mtr`;
}

function buildGroupedAdjustmentItems(groupedItems) {
    if (!groupedItems.size) {
        throw new Error("No adjustment rows were detected in the pasted table");
    }

    adjustmentItems.innerHTML = "";
    groupedItems.forEach((batches, itemName) => {
        const item = createAdjustmentItem(batches[0]?.unit || "Mtr");
        item.querySelector(".adjustment-item-name").value = itemName;
        const batchContainer = item.querySelector(".adjustment-batches");
        batchContainer.innerHTML = "";
        batches.forEach((batch) => {
            const row = createAdjustmentBatchRow(batch.unit || "Mtr");
            row.querySelector(".adjustment-batch-no").value = batch.batchNo;
            row.querySelector(".adjustment-mfg-no").value = batch.mfgNo;
            row.querySelector(".adjustment-size").value = batch.size;
            row.querySelector(".adjustment-batch-unit").value = batch.unit || "Mtr";
            row.querySelector(".adjustment-cost").value = batch.cost;
            batchContainer.appendChild(row);
        });
        adjustmentItems.appendChild(item);
    });

    renderAdjustmentsPreview();
}

function importPlainTextAdjustments(text) {
    const lines = String(text || "")
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);

    if (lines.length < 2) {
        throw new Error("Paste one item line and at least one batch line");
    }

    const groupedItems = new Map();
    let currentItem = DEFAULT_ADJUSTMENT_ITEM_NAME;

    lines.forEach((line) => {
        const copiedRowMatch = line.match(/^([A-Za-z]{1,3}\d+)\s*-\s*(-?\d+(?:\.\d+)?)\s+([A-Za-z.]+)\s+(Main\s+Location)\s+(-?\d+(?:\.\d+)?)\s+([A-Za-z.]+)\s+(-?\d+(?:\.\d+)?)/i);
        if (copiedRowMatch) {
            const [, batchNo, mfgNo, mfgUnit, warehouse, quantity, quantityUnit, cost] = copiedRowMatch;
            if (!groupedItems.has(currentItem)) {
                groupedItems.set(currentItem, []);
            }
            groupedItems.get(currentItem).push({
                batchNo,
                mfgNo: `${mfgNo} ${normalizeUnitLabel(mfgUnit)}`,
                warehouse,
                size: quantity,
                unit: normalizeSummaryUnit(quantityUnit || mfgUnit),
                cost,
            });
            return;
        }

        const batchMatch = line.match(/^\d+\.\s*([A-Za-z]{1,3}\d+)\s*-\s*(.+?)\s+(-?\d+(?:\.\d+)?)\s+([A-Za-z.]+)\s+(-?\d+(?:\.\d+)?)$/i);
        if (batchMatch) {
            if (!currentItem) {
                return;
            }
            const [, batchNo, manufacturerText, quantity, unitText, cost] = batchMatch;
            if (!groupedItems.has(currentItem)) {
                groupedItems.set(currentItem, []);
            }
            groupedItems.get(currentItem).push({
                batchNo,
                mfgNo: manufacturerText.trim(),
                size: quantity,
                unit: normalizeSummaryUnit(unitText),
                cost,
            });
            return;
        }

        currentItem = line.replace(/^\d+\.\s*/, "").trim() || DEFAULT_ADJUSTMENT_ITEM_NAME;
        if (currentItem && !groupedItems.has(currentItem)) {
            groupedItems.set(currentItem, []);
        }
    });

    buildGroupedAdjustmentItems(groupedItems);
}

function importSummaryAdjustments(rows) {
    const groupedItems = new Map();
    let currentItem = DEFAULT_ADJUSTMENT_ITEM_NAME;

    rows.forEach((cells) => {
        const firstCell = cells[0] || "";
        if (isAdjustmentItemLabel(firstCell)) {
            currentItem = firstCell;
            if (!groupedItems.has(currentItem)) {
                groupedItems.set(currentItem, []);
            }
            return;
        }

        if (!currentItem || !isAdjustmentBatchLabel(firstCell)) {
            return;
        }

        const batchParts = firstCell.split(/\s*-\s*/);
        const batchNo = batchParts.shift() || "";
        const manufacturerText = batchParts.join(" - ").trim();
        const quantityInfo = parseQuantityWithUnit(cells[2] || cells[1] || "");
        const manufacturerInfo = parseQuantityWithUnit(manufacturerText);
        groupedItems.get(currentItem).push({
            batchNo,
            mfgNo: buildManufacturerBatchValue(manufacturerInfo.quantity ? `${manufacturerInfo.quantity} ${manufacturerInfo.unit}` : manufacturerText),
            warehouse: DEFAULT_ADJUSTMENT_WAREHOUSE,
            size: quantityInfo.quantity,
            unit: quantityInfo.unit,
            cost: String(cells[3] || "").trim(),
        });
    });

    buildGroupedAdjustmentItems(groupedItems);
}

function importPastedAdjustments(text) {
    const rawText = String(text || "");
    const trimmedLines = rawText
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);

    const looksLikePlainTextBlock = trimmedLines.length > 0
        && !trimmedLines[0].includes("\t")
        && trimmedLines.some((line) => /^(\d+\.\s*)?[A-Za-z]{1,3}\d+\s*-/.test(line));

    if (!rawText.includes("\t") || looksLikePlainTextBlock) {
        importPlainTextAdjustments(text);
        return;
    }

    const rows = rawText
        .split(/\r?\n/)
        .map((line) => line.split("\t").map((cell) => cell.trim()))
        .filter((cells) => cells.some(Boolean));

    if (rows.length < 2) {
        throw new Error("Paste copied Excel rows including the header row");
    }

    const headerMap = new Map(rows[0].map((value, index) => [value, index]));
    const itemNameIndex = resolvePastedColumnIndex(headerMap, ["Item Name"]);
    const batchReferenceIndex = resolvePastedColumnIndex(headerMap, ["Reference#", "Batch Reference#"]);
    const quantityIndex = resolvePastedColumnIndex(headerMap, ["Quantity Adjusted"]);
    const costPriceIndex = resolvePastedColumnIndex(headerMap, ["Cost Price"]);
    const usageUnitIndex = resolvePastedColumnIndex(headerMap, ["Usage unit"]);
    const manufacturerBatchIndex = resolvePastedColumnIndex(headerMap, ["Manufacturer Batch#", "Description", "Item Desc"]);
    const reasonIndex = resolvePastedColumnIndex(headerMap, ["Reason"]);
    const dateIndex = resolvePastedColumnIndex(headerMap, ["Date"]);

    if ([itemNameIndex, batchReferenceIndex, quantityIndex, costPriceIndex].some((index) => index < 0)) {
        importSummaryAdjustments(rows);
        return;
    }

    const groupedItems = new Map();
    rows.slice(1).forEach((cells) => {
        const itemName = cells[itemNameIndex] || DEFAULT_ADJUSTMENT_ITEM_NAME;
        const rawManufacturerBatch = manufacturerBatchIndex >= 0 ? cells[manufacturerBatchIndex] || "" : "";
        const rawUsageUnit = usageUnitIndex >= 0 ? cells[usageUnitIndex] || "" : "";
        const parsedManufacturer = splitManufacturerBatch(rawManufacturerBatch, normalizeUnitLabel(rawUsageUnit || "Mtr"));
        const batch = {
            batchNo: cells[batchReferenceIndex] || "",
            mfgNo: parsedManufacturer.manufacturerBatch,
            size: cells[quantityIndex] || "",
            unit: parsedManufacturer.usageUnit,
            cost: cells[costPriceIndex] || "",
        };

        if (!groupedItems.has(itemName)) {
            groupedItems.set(itemName, []);
        }
        groupedItems.get(itemName).push(batch);

        if (reasonIndex >= 0 && cells[reasonIndex] && adjustmentsReason && !adjustmentsReason.value.trim()) {
            adjustmentsReason.value = cells[reasonIndex];
        }
        if (dateIndex >= 0 && cells[dateIndex] && adjustmentsDate && !adjustmentsDate.value.trim()) {
            const pastedDate = cells[dateIndex].replace(/\//g, "-");
            if (/^\d{4}-\d{2}-\d{2}$/.test(pastedDate)) {
                adjustmentsDate.value = pastedDate;
            }
        }
    });

    buildGroupedAdjustmentItems(groupedItems);
}

function buildAdjustmentItemsText() {
    const sections = [];
    const items = [...adjustmentItems.querySelectorAll(".adjustment-item")];

    items.forEach((item) => {
        const itemName = item.querySelector(".adjustment-item-name").value.trim();
        const batchRows = [...item.querySelectorAll(".adjustment-batch-row")]
            .map((row, index) => {
                const batchNo = row.querySelector(".adjustment-batch-no").value.trim();
                const mfgNo = row.querySelector(".adjustment-mfg-no").value.trim();
                const size = row.querySelector(".adjustment-size").value.trim();
                const unit = row.querySelector(".adjustment-batch-unit").value.trim();
                const cost = row.querySelector(".adjustment-cost").value.trim();
                if (![batchNo, mfgNo, size, cost].some(Boolean)) {
                    return "";
                }
                const manufacturerValue = buildManufacturerBatchValue(mfgNo);
                const quantityValue = [size, unit].filter(Boolean).join(" ");
                return `${index + 1}. ${batchNo} - ${manufacturerValue} - ${DEFAULT_ADJUSTMENT_WAREHOUSE} - ${quantityValue} - ${cost}`;
            })
            .filter(Boolean);

        if (itemName || batchRows.length) {
            sections.push([itemName, ...batchRows].filter(Boolean).join("\n"));
        }
    });

    return sections.join("\n");
}

function buildAdjustmentPreviewRows() {
    const rows = [];
    const itemBlocks = [...adjustmentItems.querySelectorAll(".adjustment-item")];
    const reason = FIXED_ADJUSTMENT_REASON;

    itemBlocks.forEach((item) => {
        const itemName = item.querySelector(".adjustment-item-name").value.trim();
        item.querySelectorAll(".adjustment-batch-row").forEach((row) => {
            const batchNo = row.querySelector(".adjustment-batch-no").value.trim();
            const mfgNo = row.querySelector(".adjustment-mfg-no").value.trim();
            const size = row.querySelector(".adjustment-size").value.trim();
            const unit = row.querySelector(".adjustment-batch-unit").value.trim();
            const cost = row.querySelector(".adjustment-cost").value.trim();
            if (![itemName, batchNo, mfgNo, size, cost].some(Boolean)) {
                return;
            }
            rows.push({
                itemName,
                batchNo,
                manufacturerBatch: buildManufacturerBatchValue(mfgNo),
                quantity: size,
                cost,
                reason,
            });
        });
    });

    return rows;
}

function renderAdjustmentsPreview() {
    const rows = buildAdjustmentPreviewRows();
    if (!adjustmentsPreview) {
        return;
    }
    if (!rows.length) {
        adjustmentsPreview.innerHTML = '<p class="empty-state">Preview rows will appear here.</p>';
        return;
    }

    adjustmentsPreview.innerHTML = `
        <table>
            <thead>
                <tr>
                    <th>Item Name</th>
                    <th>Batch Ref</th>
                    <th>Mfg Batch</th>
                    <th>Quantity</th>
                    <th>Cost Price</th>
                    <th>Reason</th>
                </tr>
            </thead>
            <tbody>
                ${rows.map((row) => `
                    <tr>
                        <td>${row.itemName || "-"}</td>
                        <td>${row.batchNo || "-"}</td>
                        <td>${row.manufacturerBatch || "-"}</td>
                        <td>${row.quantity || "-"}</td>
                        <td>${row.cost || "-"}</td>
                        <td>${row.reason || "-"}</td>
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;
}

async function handleAdjustmentsDownload(event) {
    event.preventDefault();

    const builtText = buildAdjustmentItemsText();
    const hasText = builtText.trim().length > 0;
    const hasFile = Boolean(adjustmentsFile.files[0]);
    if (!hasText && !hasFile) {
        setMessage(adjustmentsMessage, "Add at least one item or upload an .xls file", "error");
        return;
    }

    const body = new FormData();
    body.append("items", builtText);
    body.append("date", adjustmentsDate.value);
    body.append("reason", FIXED_ADJUSTMENT_REASON);
    if (hasFile) {
        body.append("file", adjustmentsFile.files[0]);
    }

    try {
        const blob = await request("/inventory-adjustments/export", {
            method: "POST",
            body,
            expectBlob: true,
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "inventory_adjustments.xls";
        link.click();
        URL.revokeObjectURL(url);
        setMessage(adjustmentsMessage, "Inventory adjustment sheet downloaded", "success");
    } catch (error) {
        setMessage(adjustmentsMessage, error.message, "error");
    }
}

function handleAdjustmentsPreview() {
    renderAdjustmentsPreview();
    setMessage(adjustmentsMessage, "Preview updated", "success");
}

function handlePastedAdjustmentsImport() {
    try {
        importPastedAdjustments(adjustmentsPasteInput?.value || "");
        setMessage(adjustmentsMessage, "Pasted Excel rows imported", "success");
    } catch (error) {
        setMessage(adjustmentsMessage, error.message, "error");
    }
}

async function handleAdjustmentsUpload(event) {
    event.preventDefault();
    if (!adjustmentsFile.files[0]) {
        setMessage(adjustmentsMessage, "Choose an .xls file to submit", "error");
        return;
    }

    const body = new FormData();
    body.append("file", adjustmentsFile.files[0]);
    body.append("date", adjustmentsDate.value);
    body.append("reason", FIXED_ADJUSTMENT_REASON);

    try {
        const blob = await request("/inventory-adjustments/export", {
            method: "POST",
            body,
            expectBlob: true,
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "inventory_adjustments.xls";
        link.click();
        URL.revokeObjectURL(url);
        setMessage(adjustmentsMessage, "Inventory adjustment sheet downloaded from uploaded Excel", "success");
    } catch (error) {
        setMessage(adjustmentsMessage, error.message, "error");
    }
}

async function handleTableClick(event) {
    if (isReadOnlyUser()) {
        return;
    }
    const button = event.target.closest("button");
    if (!button) {
        return;
    }

    const row = event.target.closest("tr");
    const item = state.inventory.find((entry) => getItemKey(entry) === row?.dataset.itemKey);
    if (!item) {
        return;
    }

    if (button.classList.contains("delete-button")) {
        const confirmed = window.confirm(`Delete ${joinPathParts([item.category, item.brand, item.type, item.batch_roll_no, item.size])}?`);
        if (!confirmed) {
            return;
        }
        const reason = await promptForReason("Delete Item Reason", "Why are you deleting this item?");
        if (!reason) {
            return;
        }

        try {
            await request("/delete-item", {
                method: "DELETE",
                body: JSON.stringify({
                    ...getLookupPayload(item),
                    reason,
                }),
            });
            await Promise.all([loadInventory(), loadLogs()]);
        } catch (error) {
            window.alert(error.message);
        }
    }

    if (button.classList.contains("update-button")) {
        const movementSelect = row.querySelector(".movement-select");
        const kind = getSpecializedKind(item);
        let movementPayload;
        let movementDescription;
        let resetControls;

        if (kind === "rubber_blankets") {
            if ((item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
                const sheetsInput = row.querySelector(".rubber-movement-sheets-input");
                const sheets = Number(sheetsInput?.value || 0);
                if (!Number.isInteger(sheets) || sheets <= 0) {
                    window.alert("sheets must be a positive whole number");
                    return;
                }
                movementPayload = { direction: movementSelect.value, sheets };
                movementDescription = `${formatQuantity(sheets)} sheet${sheets === 1 ? "" : "s"}`;
                resetControls = () => { sheetsInput.value = "1"; };
            } else {
            const modeControl = row.querySelector(".rubber-movement-mode");
            const amountControl = row.querySelector(".rubber-movement-input");
            const amount = Number(amountControl?.value || 0);
            if (!Number.isFinite(amount) || amount <= 0) {
                window.alert("stock movement must be greater than 0");
                return;
            }
            if (modeControl.value === "rolls") {
                if (!Number.isInteger(amount)) {
                    window.alert("full-roll movement must be a positive whole number");
                    return;
                }
                movementPayload = { direction: movementSelect.value, rolls: amount };
                movementDescription = `${formatQuantity(amount)} full roll${amount === 1 ? "" : "s"}`;
            } else {
                movementPayload = { direction: movementSelect.value, quantity: amount, unit: RUBBER_BLANKET_STOCK_UNIT };
                movementDescription = `${formatQuantity(amount)} m² of partial roll stock`;
            }
            resetControls = () => {
                modeControl.value = "rolls";
                amountControl.value = "1";
                amountControl.step = "1";
            };
            }
        } else if (kind === "calibrated_underpacking_paper") {
            if ((item.storage_type || M3Z_ROLL_STORAGE_TYPE) === M3Z_CUT_PIECE_STORAGE_TYPE) {
                const sheetsInput = row.querySelector(".movement-sheets-input");
                const sheets = Number(sheetsInput?.value || 0);
                if (!Number.isInteger(sheets) || sheets <= 0) {
                    window.alert("sheets must be a positive whole number");
                    return;
                }
                movementPayload = { direction: movementSelect.value, sheets };
                movementDescription = `${formatQuantity(sheets)} sheet${sheets === 1 ? "" : "s"}`;
                resetControls = () => { sheetsInput.value = "1"; };
            } else {
                const rollsInput = row.querySelector(".movement-rolls-input");
                const rolls = Number(rollsInput?.value || 0);
                if (!Number.isInteger(rolls) || rolls <= 0) {
                    window.alert("rolls must be a positive whole number");
                    return;
                }
                movementPayload = { direction: movementSelect.value, rolls };
                movementDescription = `${formatQuantity(rolls)} roll${rolls === 1 ? "" : "s"}`;
                resetControls = () => { rollsInput.value = "1"; };
            }
        } else if (kind !== "generic") {
            const boxesInput = row.querySelector(".movement-boxes-input");
            const looseInput = row.querySelector(".movement-loose-input");
            const boxes = Number(boxesInput?.value || 0);
            const looseUnits = Number(looseInput?.value || 0);
            if (!Number.isInteger(boxes) || boxes < 0 || !Number.isInteger(looseUnits) || looseUnits < 0) {
                window.alert("boxes and loose units must be non-negative whole numbers");
                return;
            }
            if (boxes === 0 && looseUnits === 0) {
                window.alert("stock movement must be greater than 0");
                return;
            }
            const packaging = getItemPackaging(item);
            const maximumLoose = kind === "chemical" ? packaging.containers_per_box : packaging.units_per_box;
            if (kind !== "ctcp_plates" && looseUnits >= maximumLoose) {
                window.alert(`loose units must be less than ${maximumLoose}`);
                return;
            }
            movementPayload = { direction: movementSelect.value, boxes, loose_units: looseUnits };
            movementDescription = `${formatQuantity(boxes)} boxes${looseUnits ? ` and ${formatQuantity(looseUnits)} loose units` : ""}`;
            resetControls = () => {
                boxesInput.value = kind === "ctcp_plates" ? "1" : "0";
                if (looseInput) {
                    looseInput.value = "0";
                }
            };
        } else {
            const input = row.querySelector(".delta-input");
            const movementUnitSelect = row.querySelector(".movement-unit-select");
            const movementAmount = Number(input.value);
            if (!Number.isFinite(movementAmount) || movementAmount <= 0) {
                window.alert("stock movement must be greater than 0");
                return;
            }
            if (!isRollItem(item) && !getCategoryRule(item.category).quantityAllowsDecimal && !Number.isInteger(movementAmount)) {
                window.alert("stock movement must be a whole number for this item");
                return;
            }
            let movementInStockUnit;
            try {
                movementInStockUnit = isRollItem(item)
                    ? convertMovementToSqm(item, movementAmount, movementUnitSelect.value)
                    : movementAmount;
            } catch (error) {
                window.alert(error.message);
                return;
            }
            movementPayload = roundStockQuantity(movementSelect.value === "out" ? -movementInStockUnit : movementInStockUnit);
            movementDescription = `${formatQuantity(movementAmount)} ${movementUnitSelect.options[movementUnitSelect.selectedIndex].textContent}`;
            resetControls = () => { input.value = "1"; };
        }

        const reason = await promptForReason("Stock Movement Reason", `Why are you moving ${movementDescription}?`);
        if (!reason) {
            return;
        }

        try {
            await request("/update-stock", {
                method: "PUT",
                body: JSON.stringify({
                    ...getLookupPayload(item),
                    ...(kind === "generic" ? { quantity_change: movementPayload } : { movement: movementPayload }),
                    reason,
                }),
            });
            resetControls();
            movementSelect.value = "in";
            await Promise.all([loadInventory(), loadLogs()]);
        } catch (error) {
            window.alert(error.message);
        }
    }
}

async function handleUsersTableChange(event) {
    const select = event.target.closest(".user-role-select");
    if (!select) {
        return;
    }

    const userId = select.dataset.userId;
    const nextRole = select.value;
    const targetUser = state.users.find((user) => user.id === userId);
    if (!targetUser) {
        return;
    }

    try {
        await request(`/admin/users/${userId}/role`, {
            method: "PUT",
            body: JSON.stringify({ role: nextRole }),
        });
        setMessage(adminUsersMessage, `Updated role for ${targetUser.email}`, "success");
        await loadUsers();
    } catch (error) {
        select.value = targetUser.role;
        setMessage(adminUsersMessage, error.message, "error");
    }
}

async function handleUsersTableClick(event) {
    const button = event.target.closest(".remove-user-button");
    if (!button) {
        return;
    }

    const userId = button.dataset.userId;
    const targetUser = state.users.find((user) => user.id === userId);
    if (!targetUser) {
        return;
    }

    const confirmed = window.confirm(`Remove ${targetUser.email}?`);
    if (!confirmed) {
        return;
    }

    try {
        await request(`/admin/users/${userId}`, {
            method: "DELETE",
        });
        setMessage(adminUsersMessage, `Removed ${targetUser.email}`, "success");
        await loadUsers();
    } catch (error) {
        setMessage(adminUsersMessage, error.message, "error");
    }
}

function debounce(callback, delay = 300) {
    let timeoutId;
    return (...args) => {
        window.clearTimeout(timeoutId);
        timeoutId = window.setTimeout(() => callback(...args), delay);
    };
}

const debouncedLoadInventory = debounce(loadInventory, 250);

async function initializeAppData() {
    if (!state.user) {
        return;
    }
    await Promise.all([loadInventory(), loadLogs(), loadUsers()]);
}

window.addEventListener("hashchange", () => {
    if (!state.user) {
        return;
    }
    const page = getCurrentPage();
    if ((page === "add-item" || page === "excel" || page === "inventory-adjustments") && !userHasRole("workshop")) {
        window.location.hash = "#overview";
        return;
    }
    if (page === "admin" && !userHasRole("admin")) {
        window.location.hash = "#overview";
        return;
    }
    showPage(getCurrentPage());
});

authTabs.forEach((tab) => {
    tab.addEventListener("click", () => setAuthTab(tab.dataset.authTab));
});

loginForm.addEventListener("submit", handleLogin);
signupForm.addEventListener("submit", handleSignup);
forgotForm.addEventListener("submit", handleForgotPassword);
itemForm.addEventListener("submit", handleAddItem);
excelForm.addEventListener("submit", handleExcelUpload);
excelForm.querySelectorAll('input[name="excelMode"]').forEach((control) => {
    control.addEventListener("change", () => {
        const isUpdate = getExcelMode() === "update";
        excelUpdateWarning.classList.toggle("is-hidden", !isUpdate);
        resetExcelPreview();
        setMessage(excelMessage, isUpdate ? "Update mode selected. Upload the workbook to review its exact deletion scope." : "");
    });
});
excelFileInput.addEventListener("change", resetExcelPreview);
excelApplyButton.addEventListener("click", handleExcelApply);
excelErrorsButton.addEventListener("click", () => {
    excelErrorsPanel.classList.toggle("is-hidden");
});
adjustmentsForm.addEventListener("submit", handleAdjustmentsDownload);
adjustmentsUploadForm?.addEventListener("submit", handleAdjustmentsUpload);
addAdjustmentItemButton?.addEventListener("click", () => {
    addAdjustmentItem("Mtr");
});
previewAdjustmentsButton?.addEventListener("click", handleAdjustmentsPreview);
importPastedAdjustmentsButton?.addEventListener("click", handlePastedAdjustmentsImport);
formCategorySelect.addEventListener("change", () => renderAddItemForm(getCategoryKey(formCategorySelect.value)));
dynamicItemFields.addEventListener("change", (event) => {
    const categoryKey = getCategoryKey(formCategorySelect.value);
    const category = getCategoryLabel(categoryKey);
    const controlName = event.target.name;
    if (categoryKey === "rubber_blankets" && controlName === "storage_type") {
        renderRubberBlanketModeFields(getDynamicFormValues());
    } else if (categoryKey === "rubber_blankets" && controlName === "width_unit") {
        updateRubberBlanketDependentFields(getDynamicFormValues());
    } else if (categoryKey === "rubber_blankets" && controlName === "length_unit") {
        convertRollPaperControlValue(event.target);
    } else if (categoryKey === "rubber_blankets" && controlName === "blanket_name") {
        updateRubberBlanketDependentFields(getDynamicFormValues());
    } else if (categoryKey === "rubber_blankets" && controlName === "thickness") {
        updateRubberBlanketDependentFields(getDynamicFormValues());
    } else if (categoryKey === "calibrated_underpacking_paper" && controlName === "storage_type") {
        renderM3ZModeFields(getDynamicFormValues());
    } else if (categoryKey === "calibrated_underpacking_paper" && ["width_unit", "length_unit"].includes(controlName)) {
        convertRollPaperControlValue(event.target);
    } else if (categoryKey === "calibrated_underpacking_paper" && controlName === "thickness_micron") {
        updateM3ZThicknessHidden();
    } else if (categoryKey === "creasing_matrix" && controlName === "thickness") {
        updateCreasingSizeOptions();
    } else if (categoryKey === "ctcp_plates" && controlName === "thickness") {
        updateCTCPSizeOptions();
    } else if (CHEMICAL_CATEGORIES.has(category) && controlName === "product") {
        updateChemicalPackOptions();
    } else if (CHEMICAL_CATEGORIES.has(category) && controlName === "containers_per_box") {
        updateChemicalPackOptions(event.target.value);
    } else if (!CHEMICAL_CATEGORIES.has(category) && !["creasing_matrix", "ctcp_plates", "calibrated_underpacking_paper", "rubber_blankets"].includes(categoryKey) && controlName === "unit") {
        renderGenericAddItemForm(category, getDynamicFormValues());
    }
    updateRollQuantityEstimate();
    updateSpecializedCalculation();
});
dynamicItemFields.addEventListener("input", (event) => {
    if (event.target.name === "type") {
        const rule = getCategoryRule(formCategorySelect.value);
        const normalizedType = event.target.value.trim().toLowerCase();
        const unitControl = dynamicItemFields.querySelector('[name="unit"]');
        if (rule.unitLinkedToType && rule.typeOptions?.includes(normalizedType) && unitControl) {
            unitControl.value = normalizedType;
        }
    }
    updateRollQuantityEstimate();
    updateSpecializedCalculation();
});
logoutButton.addEventListener("click", handleLogout);
refreshButton.addEventListener("click", async () => {
    await initializeAppData();
});
adminPanelButton.addEventListener("click", handleAdminPanelClick);
importTemplateButton.addEventListener("click", handleImportTemplateDownload);
currentStockButton.addEventListener("click", handleCurrentStockDownload);
exportCurrentButton.addEventListener("click", handleCurrentStockDownload);
excelExportButton.addEventListener("click", handleUpdateSheetDownload);
inventoryTableBody.addEventListener("click", handleTableClick);
inventoryTableBody.addEventListener("change", (event) => {
    if (!event.target.classList.contains("rubber-movement-mode")) {
        return;
    }
    const input = event.target.closest(".update-controls")?.querySelector(".rubber-movement-input");
    if (input) {
        input.step = event.target.value === "rolls" ? "1" : "0.0001";
        input.value = "1";
    }
});
usersTableBody.addEventListener("change", handleUsersTableChange);
usersTableBody.addEventListener("click", handleUsersTableClick);
searchInput.addEventListener("input", debouncedLoadInventory);
categoryFilter.addEventListener("change", loadInventory);
brandFilter.addEventListener("change", loadInventory);
thicknessFilter.addEventListener("input", debouncedLoadInventory);
typeFilter.addEventListener("change", loadInventory);
lowStockOnly.addEventListener("change", loadInventory);
lowStockThreshold.addEventListener("input", debouncedLoadInventory);

adjustmentItems?.addEventListener("click", (event) => {
    const addBatchButton = event.target.closest(".add-batch-button");
    if (addBatchButton) {
        const item = addBatchButton.closest(".adjustment-item");
        item?.querySelector(".adjustment-batches")?.appendChild(createAdjustmentBatchRow("Mtr"));
        return;
    }

    const removeBatchButton = event.target.closest(".remove-batch-button");
    if (removeBatchButton) {
        const item = removeBatchButton.closest(".adjustment-item");
        const rows = item ? [...item.querySelectorAll(".adjustment-batch-row")] : [];
        if (rows.length === 1) {
            rows[0].querySelectorAll("input").forEach((input) => {
                input.value = "";
            });
            rows[0].querySelector(".adjustment-batch-unit").value = "Mtr";
            return;
        }
        removeBatchButton.closest(".adjustment-batch-row")?.remove();
        return;
    }

    const removeItemButton = event.target.closest(".remove-item-button");
    if (removeItemButton) {
        const items = [...adjustmentItems.querySelectorAll(".adjustment-item")];
        if (items.length === 1) {
            adjustmentItems.innerHTML = "";
            addAdjustmentItem("Mtr");
            return;
        }
        removeItemButton.closest(".adjustment-item")?.remove();
    }
});

adjustmentItems?.addEventListener("input", () => {
    renderAdjustmentsPreview();
});

adjustmentsReason?.addEventListener("input", renderAdjustmentsPreview);
adjustmentsPasteInput?.addEventListener("paste", () => {
    window.setTimeout(() => {
        handlePastedAdjustmentsImport();
    }, 0);
});

if (adjustmentsDate) {
    const today = new Date();
    const offsetDate = new Date(today.getTime() - today.getTimezoneOffset() * 60000);
    adjustmentsDate.value = offsetDate.toISOString().slice(0, 10);
}

if (adjustmentItems && adjustmentItems.children.length === 0) {
    addAdjustmentItem("Mtr");
}
renderAdjustmentsPreview();

window.OnlyStockWarehouseBridge = Object.freeze({
    getInventory() {
        return state.inventory.map((item) => ({ ...item }));
    },
    getUser() {
        return state.user ? { ...state.user } : null;
    },
    loadAllInventory() {
        return request("/inventory");
    },
    viewInventoryItem(itemId) {
        const item = state.inventory.find((candidate) => candidate.id === itemId);
        if (item) {
            const searchTerm = [item.blanket_name, item.brand, item.type, item.category]
                .find((value) => value && value !== "__none__");
            searchInput.value = searchTerm || item.category || "";
        }
        window.location.hash = "#inventory";
        return loadInventory();
    },
});

async function bootstrapOnlyStock() {
    try {
        await loadInventoryConfig();
    } catch (error) {
        setMessage(formMessage, error.message, "error");
    }

    populateCategoryFilterOptions();
    populateFormCategoryOptions();
    showPage(getCurrentPage());
    renderAddItemForm(formCategorySelect.value);
    setAuthTab("login");
    setAuthenticatedUser(null);
    const isAuthenticated = await checkAuthSession();
    if (isAuthenticated) {
        showPage(getCurrentPage());
        await initializeAppData();
    }
}

bootstrapOnlyStock();
