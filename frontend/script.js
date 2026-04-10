const API_BASE_URL = window.location.protocol.startsWith("http")
    ? window.location.origin
    : "http://localhost:5000";

const state = {
    inventory: [],
    logs: [],
};

const itemForm = document.getElementById("itemForm");
const excelForm = document.getElementById("excelForm");
const excelFileInput = document.getElementById("excelFile");
const refreshButton = document.getElementById("refreshButton");
const exportButton = document.getElementById("exportButton");
const inventoryTableBody = document.getElementById("inventoryTableBody");
const categoryList = document.getElementById("categoryList");
const categoryCount = document.getElementById("categoryCount");
const treeView = document.getElementById("treeView");
const logsList = document.getElementById("logsList");
const statusText = document.getElementById("statusText");
const formMessage = document.getElementById("formMessage");
const excelMessage = document.getElementById("excelMessage");
const rowTemplate = document.getElementById("inventoryRowTemplate");

const searchInput = document.getElementById("searchInput");
const categoryFilter = document.getElementById("categoryFilter");
const brandFilter = document.getElementById("brandFilter");
const typeFilter = document.getElementById("typeFilter");
const lowStockOnly = document.getElementById("lowStockOnly");
const lowStockThreshold = document.getElementById("lowStockThreshold");
const formCategorySelect = document.getElementById("formCategorySelect");

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

function setMessage(element, text, tone = "") {
    element.textContent = text || "";
    if (tone) {
        element.dataset.tone = tone;
    } else {
        delete element.dataset.tone;
    }
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
    return `${item.category}|${item.brand}|${item.type}|${item.size}`;
}

function getLookupPayload(item) {
    return {
        category: item.category,
        brand: item.brand,
        type: item.type,
        size: item.size,
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
}

function renderCategories(items) {
    const counts = items.reduce((map, item) => {
        map[item.category] = (map[item.category] || 0) + 1;
        return map;
    }, {});

    const categories = Object.keys(counts).sort();
    categoryCount.textContent = String(categories.length);

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

function renderTable(items) {
    statusText.textContent = `${items.length} item(s) found`;

    if (items.length === 0) {
        inventoryTableBody.innerHTML = '<tr><td colspan="8" class="empty-state">No inventory available</td></tr>';
        return;
    }

    inventoryTableBody.innerHTML = "";

    items.forEach((item) => {
        const row = rowTemplate.content.firstElementChild.cloneNode(true);
        row.dataset.itemKey = getItemKey(item);

        row.querySelector('[data-field="category"]').textContent = item.category;
        row.querySelector('[data-field="brand"]').textContent = item.brand;
        row.querySelector('[data-field="type"]').textContent = item.type;
        row.querySelector('[data-field="size"]').textContent = item.size;
        row.querySelector('[data-field="quantity"]').textContent = item.quantity;
        row.querySelector('[data-field="unit"]').textContent = item.unit;

        inventoryTableBody.appendChild(row);
    });
}

function renderLogs(logs) {
    if (logs.length === 0) {
        logsList.innerHTML = '<p class="empty-state">No stock history yet</p>';
        return;
    }

    logsList.innerHTML = logs.map((log) => `
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
        renderTable(state.inventory);
        populateSelectOptions(state.inventory);
        setMessage(statusText, `${state.inventory.length} item(s) found`);
    } catch (error) {
        state.inventory = [];
        renderCategories([]);
        renderTree([]);
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
    }
}

function validateForm(formData) {
    const categoryValue = formCategorySelect.value.trim();

    if (!categoryValue) {
        return "category is required";
    }

    const requiredFields = ["brand", "type", "size", "quantity", "unit"];
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
    if (!Number.isInteger(quantity) || quantity < 0) {
        return "quantity must be a non-negative integer";
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
        brand: formData.get("brand").trim(),
        type: formData.get("type").trim(),
        size: formData.get("size").trim(),
        quantity: Number(formData.get("quantity")),
        unit: formData.get("unit").trim(),
    };

    try {
        await request("/add-item", {
            method: "POST",
            body: JSON.stringify(payload),
        });

        itemForm.reset();
        setMessage(formMessage, "Item added successfully", "success");
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
        const quantityChange = Number(input.value);

        if (!Number.isInteger(quantityChange)) {
            window.alert("quantity_change must be a whole number");
            return;
        }

        try {
            await request("/update-stock", {
                method: "PUT",
                body: JSON.stringify({
                    ...getLookupPayload(item),
                    quantity_change: quantityChange,
                }),
            });
            input.value = "1";
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

itemForm.addEventListener("submit", handleAddItem);
excelForm.addEventListener("submit", handleExcelUpload);
refreshButton.addEventListener("click", async () => {
    await Promise.all([loadInventory(), loadLogs()]);
});
exportButton.addEventListener("click", handleExportExcel);
inventoryTableBody.addEventListener("click", handleTableClick);
searchInput.addEventListener("input", debouncedLoadInventory);
categoryFilter.addEventListener("change", loadInventory);
brandFilter.addEventListener("change", loadInventory);
typeFilter.addEventListener("change", loadInventory);
lowStockOnly.addEventListener("change", loadInventory);
lowStockThreshold.addEventListener("input", debouncedLoadInventory);

Promise.all([loadInventory(), loadLogs()]);
