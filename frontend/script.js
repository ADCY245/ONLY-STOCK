const API_BASE_URL = window.location.protocol.startsWith("http")
    ? window.location.origin
    : "http://localhost:5000";

const state = {
    inventory: [],
    logs: [],
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
    excel: {
        eyebrow: "Workspace",
        title: "Excel Tools",
        description: "Import bulk updates and export the current inventory to spreadsheet format.",
    },
    logs: {
        eyebrow: "Workspace",
        title: "Stock Logs",
        description: "Review recent manual and Excel-driven stock changes in one place.",
    },
};

const itemForm = document.getElementById("itemForm");
const excelForm = document.getElementById("excelForm");
const excelFileInput = document.getElementById("excelFile");
const refreshButton = document.getElementById("refreshButton");
const exportButton = document.getElementById("exportButton");
const excelExportButton = document.getElementById("excelExportButton");
const inventoryTableBody = document.getElementById("inventoryTableBody");
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
const statusText = document.getElementById("statusText");
const formMessage = document.getElementById("formMessage");
const excelMessage = document.getElementById("excelMessage");
const rowTemplate = document.getElementById("inventoryRowTemplate");
const pageEyebrow = document.getElementById("pageEyebrow");
const pageTitle = document.getElementById("pageTitle");
const pageDescription = document.getElementById("pageDescription");
const pages = [...document.querySelectorAll("[data-page]")];
const pageLinks = [...document.querySelectorAll("[data-page-link]")];

const searchInput = document.getElementById("searchInput");
const categoryFilter = document.getElementById("categoryFilter");
const brandFilter = document.getElementById("brandFilter");
const typeFilter = document.getElementById("typeFilter");
const lowStockOnly = document.getElementById("lowStockOnly");
const lowStockThreshold = document.getElementById("lowStockThreshold");
const thicknessFilter = document.getElementById("thicknessFilter");
const formCategorySelect = document.getElementById("formCategorySelect");
const formUnitSelect = document.getElementById("formUnitSelect");
const formWidthInput = document.getElementById("formWidthInput");
const formHeightInput = document.getElementById("formHeightInput");
const formThicknessInput = document.getElementById("formThicknessInput");
const thicknessHint = document.getElementById("thicknessHint");

const formBrandInput = itemForm.querySelector('input[name="brand"]');
const formTypeInput = itemForm.querySelector('input[name="type"]');
const brandLabel = formBrandInput.closest("label");
const typeLabel = formTypeInput.closest("label");
const widthLabel = formWidthInput.closest("label");
const heightLabel = formHeightInput.closest("label");
const thicknessLabel = formThicknessInput.closest("label");

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
];

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

const CATEGORY_RULES = {
    "Rubber Blankets": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: true,
        requiresType: true,
        unitOptions: ["pcs", "rolls"],
        defaultUnit: "pcs",
    },
    "Metalback Blankets": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "mm",
        requiresBrand: true,
        requiresType: true,
        unitOptions: ["pcs", "rolls"],
        defaultUnit: "pcs",
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
    "Calibrated Underpacking Paper": {
        usesDimensions: true,
        requiresThickness: true,
        thicknessUnit: "micron",
        requiresBrand: true,
        requiresType: true,
        unitOptions: ["pcs", "rolls"],
        defaultUnit: "pcs",
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
}

function getCategoryRule(category) {
    return CATEGORY_RULES[category] || {
        usesDimensions: false,
        requiresThickness: false,
        requiresBrand: true,
        requiresType: true,
        quantityAllowsDecimal: false,
        unitLinkedToType: false,
        unitOptions: UNIT_OPTIONS,
        defaultUnit: "pcs",
    };
}

function updateCategoryDrivenFields() {
    const rule = getCategoryRule(formCategorySelect.value);
    const previousUnit = formUnitSelect.value;

    brandLabel.style.display = rule.requiresBrand ? "" : "none";
    typeLabel.style.display = rule.requiresType ? "" : "none";
    formBrandInput.disabled = !rule.requiresBrand;
    formTypeInput.disabled = !rule.requiresType;
    if (!rule.requiresBrand) {
        formBrandInput.value = "";
    }
    if (!rule.requiresType) {
        formTypeInput.value = "";
    }

    formWidthInput.disabled = !rule.usesDimensions;
    formHeightInput.disabled = !rule.usesDimensions;
    formThicknessInput.disabled = !rule.requiresThickness;

    widthLabel.style.display = rule.usesDimensions ? "" : "none";
    heightLabel.style.display = rule.usesDimensions ? "" : "none";
    thicknessLabel.style.display = rule.requiresThickness ? "" : "none";

    if (!rule.usesDimensions) {
        formWidthInput.value = "";
        formHeightInput.value = "";
    }
    if (!rule.requiresThickness) {
        formThicknessInput.value = "";
        formThicknessInput.placeholder = "Not required";
        thicknessHint.textContent = "Thickness is not required for this category.";
    } else {
        formThicknessInput.placeholder = `Enter thickness in ${rule.thicknessUnit}`;
        thicknessHint.textContent = `Thickness unit for this category: ${rule.thicknessUnit}.`;
    }

    formUnitSelect.innerHTML = `
        <option value="">Select unit</option>
        ${rule.unitOptions.map((unit) => `<option value="${unit}">${unit}</option>`).join("")}
    `;

    if (rule.unitOptions.includes(previousUnit)) {
        formUnitSelect.value = previousUnit;
    } else {
        formUnitSelect.value = rule.defaultUnit;
    }

    if (rule.unitLinkedToType && rule.typeOptions) {
        const normalizedType = formTypeInput.value.trim().toLowerCase();
        if (rule.typeOptions.includes(normalizedType)) {
            formUnitSelect.value = normalizedType;
        }
    }
}

formTypeInput.addEventListener("input", () => {
    const rule = getCategoryRule(formCategorySelect.value);
    if (!rule.unitLinkedToType || !rule.typeOptions) {
        return;
    }
    const normalized = formTypeInput.value.trim().toLowerCase();
    if (rule.typeOptions.includes(normalized)) {
        formUnitSelect.value = normalized;
    }
});

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
    return [
        item.category,
        item.brand,
        item.type,
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
        width: item.width,
        height: item.height,
        thickness: item.thickness,
    };
}

async function request(path, options = {}) {
    const response = await fetch(`${API_BASE_URL}${path}`, {
        ...options,
        headers: {
            ...(options.body instanceof FormData ? {} : { "Content-Type": "application/json" }),
            ...(options.headers || {}),
        },
    });

    if (options.expectBlob) {
        if (!response.ok) {
            throw new Error("Unable to export Excel");
        }
        return response.blob();
    }

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        throw new Error(data.error || "Request failed");
    }

    return data;
}

function populateSelectOptions(items) {
    const uniqueBrands = [...new Set(items.map((item) => item.brand))].sort();
    const uniqueTypes = [...new Set(items.map((item) => item.type))].sort();

    const selectedCategory = categoryFilter.value;
    const selectedBrand = brandFilter.value;
    const selectedType = typeFilter.value;

    categoryFilter.innerHTML = '<option value="">All categories</option>' +
        CATEGORY_OPTIONS.map((option) => `<option value="${option.label}">${option.code} - ${option.label}</option>`).join("");
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

    const selectedFormCategory = formCategorySelect.value;
    formCategorySelect.innerHTML = `
        <option value="">Select category</option>
        ${CATEGORY_OPTIONS.map((option) => `<option value="${option.label}">${option.code} - ${option.label}</option>`).join("")}
    `;

    if (CATEGORY_OPTIONS.some((option) => option.label === selectedFormCategory)) {
        formCategorySelect.value = selectedFormCategory;
    }
    updateCategoryDrivenFields();
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
        tree[item.category] ??= {};
        tree[item.category][item.brand] ??= {};
        tree[item.category][item.brand][item.type] ??= [];
        tree[item.category][item.brand][item.type].push(item);
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
                    .map((item) => `<div class="tree-leaf">${item.size} - ${item.quantity} ${item.unit}</div>`)
                    .join("");

                return `
                    <details>
                        <summary>${itemType}</summary>
                        <div class="tree-children">${sizes}</div>
                    </details>
                `;
            }).join("");

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

function renderTable(items) {
    statusText.textContent = `${items.length} item(s) found`;

    if (items.length === 0) {
        inventoryTableBody.innerHTML = '<tr><td colspan="10" class="empty-state">No inventory available</td></tr>';
        return;
    }

    inventoryTableBody.innerHTML = "";

    items.forEach((item) => {
        const row = rowTemplate.content.firstElementChild.cloneNode(true);
        row.dataset.itemKey = getItemKey(item);

        row.querySelector('[data-field="category"]').textContent = item.category;
        row.querySelector('[data-field="brand"]').textContent = item.brand;
        row.querySelector('[data-field="type"]').textContent = item.type;
        row.querySelector('[data-field="width"]').textContent = item.width || "-";
        row.querySelector('[data-field="height"]').textContent = item.height || "-";
        row.querySelector('[data-field="thickness"]').textContent = item.thickness || "-";
        row.querySelector('[data-field="quantity"]').textContent = item.quantity;
        row.querySelector('[data-field="unit"]').textContent = item.unit;

        inventoryTableBody.appendChild(row);
    });
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

    const markup = logs.map((log) => `
        <div class="log-entry">
            <p><strong>${log.action}</strong> via ${log.source}</p>
            <p>${log.category} / ${log.brand} / ${log.type} / ${log.size}</p>
            <p>${log.quantity_before} -> ${log.quantity_after} ${log.unit}</p>
            <p>${new Date(log.changed_at).toLocaleString()}</p>
        </div>
    `).join("");

    logsList.innerHTML = markup;
    overviewLogs.innerHTML = logs.slice(0, 4).map((log) => `
        <div class="log-entry">
            <p><strong>${log.action}</strong> via ${log.source}</p>
            <p>${log.category} / ${log.brand} / ${log.type} / ${log.size}</p>
            <p>${log.quantity_before} -> ${log.quantity_after} ${log.unit}</p>
            <p>${new Date(log.changed_at).toLocaleString()}</p>
        </div>
    `).join("");
}

async function loadInventory() {
    setMessage(statusText, "Loading inventory...");

    try {
        const params = buildParams();
        const path = params.toString() ? `/inventory?${params.toString()}` : "/inventory";
        const items = await request(path);
        state.inventory = Array.isArray(items) ? items : [];

        renderCategories(state.inventory);
        renderTree(state.inventory);
        renderOverviewStats(state.inventory);
        renderTable(state.inventory);
        populateSelectOptions(state.inventory);
        setMessage(statusText, `${state.inventory.length} item(s) found`);
    } catch (error) {
        state.inventory = [];
        renderCategories([]);
        renderTree([]);
        renderOverviewStats([]);
        renderTable([]);
        setMessage(statusText, error.message, "error");
    }
}

async function loadLogs() {
    try {
        const logs = await request("/stock-logs?limit=12");
        state.logs = Array.isArray(logs) ? logs : [];
        renderLogs(state.logs);
    } catch (error) {
        logsList.innerHTML = `<p class="empty-state">${error.message}</p>`;
        overviewLogs.innerHTML = `<p class="empty-state">${error.message}</p>`;
    }
}

function validateForm(formData) {
    const categoryValue = formCategorySelect.value.trim();
    const unitValue = formUnitSelect.value.trim();
    const categoryRule = getCategoryRule(categoryValue);

    if (!categoryValue) {
        return "category is required";
    }
    if (!unitValue) {
        return "unit is required";
    }

    const requiredFields = ["quantity"];
    if (categoryRule.requiresBrand) {
        requiredFields.push("brand");
    }
    if (categoryRule.requiresType) {
        requiredFields.push("type");
    }
    for (const field of requiredFields) {
        const value = formData.get(field);
        if (typeof value === "string" && !value.trim()) {
            return `${field} is required`;
        }
        if (value === null) {
            return `${field} is required`;
        }
    }

    const quantity = Number(formData.get("quantity"));
    if (!Number.isFinite(quantity) || quantity < 0) {
        return "quantity must be a non-negative number";
    }
    if (!categoryRule.quantityAllowsDecimal && !Number.isInteger(quantity)) {
        return "quantity must be a non-negative integer";
    }

    if (categoryRule.unitLinkedToType) {
        const typeValue = (formData.get("type") || "").trim().toLowerCase();
        if (!typeValue) {
            return "type is required";
        }
        if (!categoryRule.typeOptions || !categoryRule.typeOptions.includes(typeValue)) {
            return `type must be one of: ${(categoryRule.typeOptions || []).join(", ")}`;
        }
        if (unitValue.toLowerCase() !== typeValue) {
            return "unit must match type for this category";
        }
    }

    if (categoryRule.typeIsFormat) {
        const typeValue = (formData.get("type") || "").trim();
        if (!/^\d+(?:\.\d+)?\s*(ltr|l|kg|g|ml)$/i.test(typeValue)) {
            return "type must be a format like 1ltr, 5 ltr, 1kg";
        }
    }

    if (categoryRule.usesDimensions) {
        if (!formWidthInput.value.trim() || !formHeightInput.value.trim()) {
            return "width and height are required for this category";
        }
    }

    if (categoryRule.requiresThickness && !formThicknessInput.value.trim()) {
        return `thickness is required in ${categoryRule.thicknessUnit}`;
    }

    return "";
}

async function handleAddItem(event) {
    event.preventDefault();

    const formData = new FormData(itemForm);
    const validationError = validateForm(formData);
    if (validationError) {
        setMessage(formMessage, validationError, "error");
        return;
    }

    const payload = {
        category: formCategorySelect.value.trim(),
        brand: categoryRule.requiresBrand ? formData.get("brand").trim() : "",
        type: categoryRule.requiresType ? formData.get("type").trim() : "",
        width: formWidthInput.value.trim(),
        height: formHeightInput.value.trim(),
        thickness: formThicknessInput.value.trim(),
        quantity: Number(formData.get("quantity")),
        unit: formUnitSelect.value.trim(),
    };

    try {
        await request("/add-item", {
            method: "POST",
            body: JSON.stringify(payload),
        });

        itemForm.reset();
        updateCategoryDrivenFields();
        setMessage(formMessage, "Item added successfully", "success");
        if (window.location.hash !== "#inventory") {
            window.location.hash = "#inventory";
        }
        await Promise.all([loadInventory(), loadLogs()]);
    } catch (error) {
        setMessage(formMessage, error.message, "error");
    }
}

async function handleExcelUpload(event) {
    event.preventDefault();

    const file = excelFileInput.files[0];
    if (!file) {
        setMessage(excelMessage, "Select an Excel file to upload", "error");
        return;
    }

    const body = new FormData();
    body.append("file", file);

    try {
        const response = await request("/upload-excel", {
            method: "POST",
            body,
        });

        excelForm.reset();
        setMessage(
            excelMessage,
            `Excel uploaded: ${response.inserted} inserted, ${response.updated} updated`,
            "success"
        );
        await Promise.all([loadInventory(), loadLogs()]);
    } catch (error) {
        setMessage(excelMessage, error.message, "error");
    }
}

async function handleExportExcel() {
    try {
        const blob = await request("/export-excel", {
            method: "GET",
            expectBlob: true,
        });

        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "inventory_export.xlsx";
        link.click();
        URL.revokeObjectURL(url);
        setMessage(excelMessage, "Excel export downloaded", "success");
    } catch (error) {
        setMessage(excelMessage, error.message, "error");
    }
}

async function handleTableClick(event) {
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
        const confirmed = window.confirm(`Delete ${item.category} / ${item.brand} / ${item.type} / ${item.size}?`);
        if (!confirmed) {
            return;
        }

        try {
            await request("/delete-item", {
                method: "DELETE",
                body: JSON.stringify(getLookupPayload(item)),
            });
            await Promise.all([loadInventory(), loadLogs()]);
        } catch (error) {
            window.alert(error.message);
        }
    }

    if (button.classList.contains("update-button")) {
        const input = row.querySelector(".delta-input");
        const movementSelect = row.querySelector(".movement-select");
        const movementAmount = Number(input.value);

        if (!Number.isInteger(movementAmount) || movementAmount <= 0) {
            window.alert("stock movement must be a whole number greater than 0");
            return;
        }

        const quantityChange = movementSelect.value === "out"
            ? -movementAmount
            : movementAmount;

        try {
            await request("/update-stock", {
                method: "PUT",
                body: JSON.stringify({
                    ...getLookupPayload(item),
                    quantity_change: quantityChange,
                }),
            });
            input.value = "1";
            movementSelect.value = "in";
            await Promise.all([loadInventory(), loadLogs()]);
        } catch (error) {
            window.alert(error.message);
        }
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

window.addEventListener("hashchange", () => {
    showPage(getCurrentPage());
});

itemForm.addEventListener("submit", handleAddItem);
excelForm.addEventListener("submit", handleExcelUpload);
formCategorySelect.addEventListener("change", updateCategoryDrivenFields);
refreshButton.addEventListener("click", async () => {
    await Promise.all([loadInventory(), loadLogs()]);
});
exportButton.addEventListener("click", handleExportExcel);
excelExportButton.addEventListener("click", handleExportExcel);
inventoryTableBody.addEventListener("click", handleTableClick);
searchInput.addEventListener("input", debouncedLoadInventory);
categoryFilter.addEventListener("change", loadInventory);
brandFilter.addEventListener("change", loadInventory);
thicknessFilter.addEventListener("input", debouncedLoadInventory);
typeFilter.addEventListener("change", loadInventory);
lowStockOnly.addEventListener("change", loadInventory);
lowStockThreshold.addEventListener("input", debouncedLoadInventory);

showPage(getCurrentPage());
updateCategoryDrivenFields();
Promise.all([loadInventory(), loadLogs()]);
