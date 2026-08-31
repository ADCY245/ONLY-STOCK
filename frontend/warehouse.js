// Proportional warehouse model derived from all three pages of Bala Warehouse copy.pdf.
// The source drawing has no surveyed dimensions, so every physical value lives here
// and can be calibrated later without changing inventory or interaction code.

const WAREHOUSE_CONFIG = {
    warehouse: { id: "WH-A", name: "Warehouse A" },
    footprint: [
        [-18, -10],
        [12, -10],
        [12, 14],
        [8, 14],
        [8, 8],
        [-18, 8],
    ],
    camera: {
        position: [25, 25, 30],
        target: [-1, 1.8, 1.5],
        minDistance: 8,
        maxDistance: 68,
    },
    wallHeight: 2.25,
    wallThickness: 0.22,
    entrance: {
        wall: "south-leg",
        from: [8, 14],
        to: [12, 14],
        openingFrom: 8.75,
        openingTo: 11.25,
        label: "Entrance",
    },
    zones: {
        blankets: { name: "Blanket Rolls", legend: "Rubber Blankets", color: "#397fba" },
        metalback: { name: "Metalback Blanket + Extra", legend: "Metalback Blanket + Extra", color: "#7c91a7" },
        rules: { name: "Cutting/Creasing Rules", legend: "Cutting/Creasing Rules", color: "#e77c2b" },
        chemicals: { name: "Chemicals (Cartons)", legend: "Chemicals", color: "#c79857" },
        matrix: { name: "Creasing Matrix", legend: "Creasing Matrix", color: "#35b9cf" },
        underpacking: { name: "Underpacking Rolls / M3Z", legend: "M3Z / Underpacking", color: "#e2c43a" },
        plates: { name: "PS Plates", legend: "PS Plates", color: "#d69f20" },
        open: { name: "Configurable / Open", legend: "Walkway / Open Area", color: "#b8c2bc" },
    },
    // Repeated inset plan: low fixtures at the left, open working area in the middle,
    // and the three blue rack runs at south, north, and east.
    fixtures: [
        { id: "UPPER-LEFT", x: -14.6, z: -7.5, width: 5.4, depth: 3.1, height: 0.65, kind: "floor-storage" },
        { id: "CARTON-GRID", x: -12.5, z: -2.9, width: 7.6, depth: 2.1, height: 0.42, kind: "carton-grid" },
        { id: "CENTRAL-GRID-1", x: -13.1, z: 1.2, width: 7.8, depth: 2.7, height: 0.32, kind: "floor-grid" },
        { id: "CENTRAL-GRID-2", x: -13.1, z: 4.8, width: 7.8, depth: 2.7, height: 0.32, kind: "floor-grid" },
        { id: "DIVIDER", x: -7.7, z: -0.4, width: 0.34, depth: 13.8, height: 1.65, kind: "divider" },
        { id: "LEFT-STEPS", x: -17.2, z: 4.0, width: 1.05, depth: 6.2, height: 0.25, kind: "steps" },
    ],
    racks: [
        {
            id: "RB02",
            label: "RB-02",
            sourcePage: 1,
            sourceWall: "South wall elevation",
            x: 2.1,
            z: 7.05,
            rotation: 0,
            bays: 3,
            bayWidth: 3.35,
            depth: 1.25,
            levels: 4,
            levelHeight: 1.24,
            levelZones: [
                ["blankets", "blankets", "blankets"],
                ["metalback", "metalback", "metalback"],
                ["rules", "rules", "rules"],
                ["chemicals", "chemicals", "chemicals"],
            ],
        },
        {
            id: "UP01",
            label: "UP-01",
            sourcePage: 2,
            sourceWall: "North wall elevation",
            x: 3.45,
            z: -9.05,
            rotation: 0,
            bays: 3,
            bayWidth: 2.95,
            depth: 1.25,
            levels: 4,
            levelHeight: 1.24,
            levelZones: [
                ["blankets", "blankets", "blankets"],
                ["open", "underpacking", "underpacking"],
                ["rules", "rules", "rules"],
                ["matrix", "matrix", "matrix"],
            ],
        },
        {
            id: "ER01",
            label: "ER-01",
            sourcePage: 3,
            sourceWall: "East wall elevation",
            x: 11.05,
            z: -2.55,
            rotation: Math.PI / 2,
            bays: 4,
            bayWidth: 3.05,
            depth: 1.25,
            levels: 4,
            levelHeight: 1.24,
            levelZones: [
                ["blankets", "blankets", "blankets", "underpacking"],
                ["underpacking", "underpacking", "underpacking", "underpacking"],
                ["rules", "rules", "underpacking", "underpacking"],
                ["matrix", "matrix", "open", "open"],
            ],
            bayNotes: [null, null, "Blanket Rolls 1080 MM", "Underpacking Rolls 1350 MM"],
        },
        {
            id: "PS01",
            label: "PS-01",
            sourcePage: 3,
            sourceWall: "Dedicated east-side plate rack",
            x: 10.7,
            z: 10.2,
            rotation: Math.PI / 2,
            bays: 1,
            bayWidth: 3.15,
            depth: 1.35,
            levels: 3,
            levelHeight: 1.45,
            levelZones: [["plates"], ["open"], ["plates"]],
        },
    ],
};

const dom = {
    page: document.querySelector('[data-page="warehouse"]'),
    canvasHost: document.getElementById("warehouseCanvasHost"),
    loading: document.getElementById("warehouseLoading"),
    view3d: document.getElementById("warehouse3dViewport"),
    viewTop: document.getElementById("warehouseTopViewport"),
    topSvg: document.getElementById("warehouseTopSvg"),
    miniMap: document.getElementById("warehouseMiniMap"),
    search: document.getElementById("warehouseSearchInput"),
    searchOptions: document.getElementById("warehouseSearchOptions"),
    locate: document.getElementById("warehouseLocateButton"),
    message: document.getElementById("warehouseMessage"),
    button3d: document.getElementById("warehouse3dButton"),
    buttonTop: document.getElementById("warehouseTopButton"),
    reset: document.getElementById("warehouseResetButton"),
    showAll: document.getElementById("warehouseShowAllButton"),
    zoomIn: document.getElementById("warehouseZoomIn"),
    zoomOut: document.getElementById("warehouseZoomOut"),
    viewportReset: document.getElementById("warehouseViewportReset"),
    viewStatus: document.getElementById("warehouseViewStatus"),
    details: document.getElementById("warehouseDetailsContent"),
    badge: document.getElementById("warehouseLocationBadge"),
    stockDetails: document.getElementById("warehouseStockDetailsButton"),
    locateAgain: document.getElementById("warehouseLocateAgainButton"),
    legend: document.getElementById("warehouseLegend"),
    warehouseSelect: document.getElementById("warehouseSelect"),
    zoneSelect: document.getElementById("warehouseZoneSelect"),
    rackSelect: document.getElementById("warehouseRackSelect"),
    levelSelect: document.getElementById("warehouseLevelSelect"),
    positionSelect: document.getElementById("warehousePositionSelect"),
};

const runtime = {
    inventory: [],
    selectedItem: null,
    selectedRackId: null,
    selectedLocationIds: new Set(),
    selectedResolution: null,
    mode: "3d",
    topScale: 1,
    initialized3d: false,
    initializing3d: false,
    threeAvailable: false,
    THREE: null,
    renderer: null,
    scene: null,
    camera: null,
    controls: null,
    raycaster: null,
    pointer: null,
    clickables: [],
    locationMeshes: new Map(),
    rackGroups: new Map(),
    highlightMarker: null,
    cameraTween: null,
    pointerDown: null,
    resizeObserver: null,
};

const SVG_NS = "http://www.w3.org/2000/svg";
const positionLetter = (index) => String.fromCharCode(65 + index);
const locationIdFor = (rack, level, bayIndex) => `${WAREHOUSE_CONFIG.warehouse.id}-${rack.id}-L${level}-${positionLetter(bayIndex)}`;

function locationWorldPosition(rack, level, bayIndex) {
    const localX = (bayIndex - (rack.bays - 1) / 2) * rack.bayWidth;
    const cos = Math.cos(rack.rotation);
    const sin = Math.sin(rack.rotation);
    return {
        x: rack.x + localX * cos,
        y: 0.52 + (level - 1) * rack.levelHeight,
        z: rack.z - localX * sin,
    };
}

const WAREHOUSE_LOCATIONS = {};
WAREHOUSE_CONFIG.racks.forEach((rack) => {
    for (let level = 1; level <= rack.levels; level += 1) {
        for (let bayIndex = 0; bayIndex < rack.bays; bayIndex += 1) {
            const locationId = locationIdFor(rack, level, bayIndex);
            const zoneKey = rack.levelZones[level - 1]?.[bayIndex] || "open";
            WAREHOUSE_LOCATIONS[locationId] = {
                location_id: locationId,
                warehouse: WAREHOUSE_CONFIG.warehouse.name,
                warehouse_id: WAREHOUSE_CONFIG.warehouse.id,
                zone_key: zoneKey,
                zone: WAREHOUSE_CONFIG.zones[zoneKey].name,
                rack: rack.label,
                rack_id: rack.id,
                level: `L${level}`,
                level_number: level,
                position: positionLetter(bayIndex),
                source_page: rack.sourcePage,
                source_wall: rack.sourceWall,
                note: rack.bayNotes?.[bayIndex] || null,
                ...locationWorldPosition(rack, level, bayIndex),
            };
        }
    }
});

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function normalizeText(value) {
    return String(value ?? "").trim().toLowerCase();
}

function isBlankValue(value) {
    return value === null || value === undefined || value === "" || value === "__none__";
}

function displayValue(value, suffix = "") {
    return isBlankValue(value) ? null : `${value}${suffix}`;
}

function formatNumber(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
        return String(value ?? "-");
    }
    return new Intl.NumberFormat("en-IN", { maximumFractionDigits: 4 }).format(number);
}

function formatUnit(unit) {
    const normalized = normalizeText(unit);
    if (["mâ²", "m²", "sq.m", "sqm"].includes(normalized)) {
        return "m²";
    }
    return unit || "";
}

function itemName(item) {
    const primary = [item?.blanket_name, item?.brand].find((value) => !isBlankValue(value));
    const type = !isBlankValue(item?.type) ? item.type : null;
    if (primary && type && !normalizeText(primary).includes(normalizeText(type))) {
        return `${primary} ${type}`;
    }
    return primary || type || item?.category || "Inventory item";
}

function itemSearchLabel(item) {
    const parts = [
        itemName(item),
        item.category,
        item.roll_no,
        item.batch_no,
        item.batch_roll_no,
        item.thickness ? `${item.thickness} ${item.thickness_unit || "mm"}` : null,
    ].filter((value, index, values) => !isBlankValue(value) && values.indexOf(value) === index);
    return parts.join(" · ");
}

function itemSearchText(item) {
    return [
        itemSearchLabel(item),
        item.brand,
        item.type,
        item.width,
        item.height,
        item.length,
        item.actual_width,
        item.nominal_width,
    ].filter(Boolean).join(" ").toLowerCase();
}

function zoneKeyForItem(item) {
    const text = normalizeText([item?.category, item?.brand, item?.type, item?.blanket_name].filter(Boolean).join(" "));
    if (/ctcp|ps plate|strip plate/.test(text)) return "plates";
    if (/creasing matrix/.test(text)) return "matrix";
    if (/underpacking|m3z/.test(text)) return "underpacking";
    if (/metalback|underlay blanket/.test(text)) return "metalback";
    if (/rubber blanket|blanket roll/.test(text)) return "blankets";
    if (/cutting rule|creasing rule|litho perforation/.test(text)) return "rules";
    if (/chemical|washing solution|fountain solution|plate care|roller care|blanket maintenance/.test(text)) return "chemicals";
    return null;
}

function exactLocationIdsForItem(item) {
    const ids = [];
    if (item?.location_id) ids.push(String(item.location_id).toUpperCase());
    if (Array.isArray(item?.locations)) {
        item.locations.forEach((entry) => {
            const id = typeof entry === "string" ? entry : entry?.location_id;
            if (id) ids.push(String(id).toUpperCase());
        });
    }
    return [...new Set(ids)].filter((id) => WAREHOUSE_LOCATIONS[id]);
}

function locationsForZone(zoneKey) {
    return Object.values(WAREHOUSE_LOCATIONS).filter((location) => location.zone_key === zoneKey);
}

function resolveItemLocations(item) {
    const exactIds = exactLocationIdsForItem(item);
    if (exactIds.length) {
        return { ids: exactIds, resolution: "exact", zoneKey: WAREHOUSE_LOCATIONS[exactIds[0]].zone_key };
    }
    const zoneKey = zoneKeyForItem(item);
    return {
        ids: zoneKey ? locationsForZone(zoneKey).map((location) => location.location_id) : [],
        resolution: zoneKey ? "zone" : "unassigned",
        zoneKey,
    };
}

function exactItemsAtLocation(locationId) {
    return runtime.inventory.filter((item) => exactLocationIdsForItem(item).includes(locationId));
}

function unassignedItemsForZone(zoneKey) {
    return runtime.inventory.filter((item) => exactLocationIdsForItem(item).length === 0 && zoneKeyForItem(item) === zoneKey);
}

function itemsForRack(rackId) {
    const zoneKeys = new Set(
        Object.values(WAREHOUSE_LOCATIONS)
            .filter((location) => location.rack_id === rackId && location.zone_key !== "open")
            .map((location) => location.zone_key),
    );
    return runtime.inventory.filter((item) => {
        const exact = exactLocationIdsForItem(item);
        if (exact.some((id) => WAREHOUSE_LOCATIONS[id]?.rack_id === rackId)) return true;
        return exact.length === 0 && zoneKeys.has(zoneKeyForItem(item));
    });
}

function setMessage(message, type = "") {
    if (!dom.message) return;
    dom.message.textContent = message || "";
    dom.message.className = `message warehouse-message${type ? ` ${type}` : ""}`;
}

function svgElement(name, attributes = {}) {
    const node = document.createElementNS(SVG_NS, name);
    Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, String(value)));
    return node;
}

function footprintPoints() {
    return WAREHOUSE_CONFIG.footprint.map(([x, z]) => `${x},${z}`).join(" ");
}

function addPlanFixtures(svg) {
    WAREHOUSE_CONFIG.fixtures.forEach((fixture) => {
        const rect = svgElement("rect", {
            x: fixture.x - fixture.width / 2,
            y: fixture.z - fixture.depth / 2,
            width: fixture.width,
            height: fixture.depth,
            rx: 0.12,
            class: "warehouse-plan-fixture",
        });
        svg.appendChild(rect);

        if (["carton-grid", "floor-grid", "steps"].includes(fixture.kind)) {
            const columns = fixture.kind === "steps" ? 1 : 5;
            const rows = fixture.kind === "steps" ? 7 : 3;
            for (let column = 1; column < columns; column += 1) {
                const x = fixture.x - fixture.width / 2 + (fixture.width / columns) * column;
                svg.appendChild(svgElement("line", {
                    x1: x,
                    y1: fixture.z - fixture.depth / 2,
                    x2: x,
                    y2: fixture.z + fixture.depth / 2,
                    stroke: "#9aa39e",
                    "stroke-width": 0.08,
                }));
            }
            for (let row = 1; row < rows; row += 1) {
                const z = fixture.z - fixture.depth / 2 + (fixture.depth / rows) * row;
                svg.appendChild(svgElement("line", {
                    x1: fixture.x - fixture.width / 2,
                    y1: z,
                    x2: fixture.x + fixture.width / 2,
                    y2: z,
                    stroke: "#9aa39e",
                    "stroke-width": 0.08,
                }));
            }
        }
    });
}

function renderPlan(svg, interactive = false) {
    if (!svg) return;
    svg.replaceChildren();
    svg.appendChild(svgElement("polygon", { points: footprintPoints(), class: "warehouse-plan-floor" }));

    const wallPath = [
        "M -18 8 L -18 -10 L 12 -10 L 12 14",
        `M 12 14 L ${WAREHOUSE_CONFIG.entrance.openingTo} 14`,
        `M ${WAREHOUSE_CONFIG.entrance.openingFrom} 14 L 8 14 L 8 8 L -18 8`,
    ];
    wallPath.forEach((d) => svg.appendChild(svgElement("path", { d, class: "warehouse-plan-wall" })));
    addPlanFixtures(svg);

    const openLabel = svgElement("text", { x: 1.6, y: 1.2, class: "warehouse-plan-open-label" });
    openLabel.textContent = "OPEN WORKING / AISLE AREA";
    svg.appendChild(openLabel);

    WAREHOUSE_CONFIG.racks.forEach((rack) => {
        const width = rack.bays * rack.bayWidth;
        const group = svgElement("g", {
            transform: `translate(${rack.x} ${rack.z}) rotate(${(-rack.rotation * 180) / Math.PI})`,
            "data-rack-id": rack.id,
        });
        const rect = svgElement("rect", {
            x: -width / 2,
            y: -rack.depth / 2,
            width,
            height: rack.depth,
            rx: 0.12,
            fill: "#377dac",
            class: "warehouse-plan-rack",
            "data-rack-id": rack.id,
            tabindex: interactive ? 0 : -1,
            role: interactive ? "button" : "img",
            "aria-label": `${rack.label}, ${rack.sourceWall}`,
        });
        group.appendChild(rect);
        for (let index = 1; index < rack.bays; index += 1) {
            const x = -width / 2 + rack.bayWidth * index;
            group.appendChild(svgElement("line", {
                x1: x,
                y1: -rack.depth / 2,
                x2: x,
                y2: rack.depth / 2,
                stroke: "#f0b04b",
                "stroke-width": 0.16,
            }));
        }
        const label = svgElement("text", { x: 0, y: 0.2, class: "warehouse-plan-rack-label" });
        label.textContent = rack.label;
        group.appendChild(label);
        svg.appendChild(group);
    });

    const doorLeft = svgElement("path", {
        d: `M ${WAREHOUSE_CONFIG.entrance.openingFrom} 14 L ${WAREHOUSE_CONFIG.entrance.openingFrom} 12.75 A 1.25 1.25 0 0 1 10 14`,
        class: "warehouse-plan-door",
    });
    const doorRight = svgElement("path", {
        d: `M ${WAREHOUSE_CONFIG.entrance.openingTo} 14 L ${WAREHOUSE_CONFIG.entrance.openingTo} 12.75 A 1.25 1.25 0 0 0 10 14`,
        class: "warehouse-plan-door",
    });
    svg.append(doorLeft, doorRight);

    const entranceLabel = svgElement("text", { x: 10, y: 15.35, class: "warehouse-plan-open-label" });
    entranceLabel.textContent = "ENTRANCE";
    svg.appendChild(entranceLabel);
    updatePlanHighlight(svg);
}

function updatePlanHighlight(svg) {
    if (!svg) return;
    svg.querySelectorAll(".warehouse-plan-highlight").forEach((node) => node.remove());
    svg.querySelectorAll(".warehouse-plan-rack").forEach((node) => {
        node.classList.toggle("is-selected", node.dataset.rackId === runtime.selectedRackId);
    });

    const highlightedByRack = new Map();
    runtime.selectedLocationIds.forEach((locationId) => {
        const location = WAREHOUSE_LOCATIONS[locationId];
        if (!location) return;
        if (!highlightedByRack.has(location.rack_id)) highlightedByRack.set(location.rack_id, new Set());
        highlightedByRack.get(location.rack_id).add(location.position);
    });

    highlightedByRack.forEach((positions, rackId) => {
        const rack = WAREHOUSE_CONFIG.racks.find((entry) => entry.id === rackId);
        if (!rack) return;
        positions.forEach((position) => {
            const bayIndex = position.charCodeAt(0) - 65;
            const localX = (bayIndex - (rack.bays - 1) / 2) * rack.bayWidth;
            const marker = svgElement("rect", {
                x: localX - rack.bayWidth * 0.43,
                y: -rack.depth * 0.47,
                width: rack.bayWidth * 0.86,
                height: rack.depth * 0.94,
                rx: 0.12,
                class: "warehouse-plan-highlight",
                transform: `translate(${rack.x} ${rack.z}) rotate(${(-rack.rotation * 180) / Math.PI})`,
            });
            svg.appendChild(marker);
        });
    });
}

function updateAllPlanHighlights() {
    updatePlanHighlight(dom.topSvg);
    updatePlanHighlight(dom.miniMap);
}

function renderLegend() {
    dom.legend.innerHTML = Object.values(WAREHOUSE_CONFIG.zones).map((zone) => `
        <span class="warehouse-legend-item">
            <span class="warehouse-legend-swatch" style="background:${zone.color}"></span>
            ${escapeHtml(zone.legend)}
        </span>
    `).join("");
}

function setSelectOptions(select, options, placeholder, selectedValue = "") {
    if (!select) return;
    select.innerHTML = `<option value="">${escapeHtml(placeholder)}</option>${options.map((option) => `
        <option value="${escapeHtml(option.value)}" ${option.value === selectedValue ? "selected" : ""}>${escapeHtml(option.label)}</option>
    `).join("")}`;
}

function populateNavigation() {
    const zoneOptions = Object.entries(WAREHOUSE_CONFIG.zones)
        .filter(([key]) => key !== "open")
        .map(([value, zone]) => ({ value, label: zone.name }));
    setSelectOptions(dom.zoneSelect, zoneOptions, "All zones");
    setSelectOptions(
        dom.rackSelect,
        WAREHOUSE_CONFIG.racks.map((rack) => ({ value: rack.id, label: `${rack.label} · ${rack.sourceWall}` })),
        "Select rack",
    );
    setSelectOptions(dom.levelSelect, [], "Select level");
    setSelectOptions(dom.positionSelect, [], "Select position");
}

function filterRackOptionsForZone(zoneKey) {
    const racks = WAREHOUSE_CONFIG.racks.filter((rack) => !zoneKey || rack.levelZones.flat().includes(zoneKey));
    setSelectOptions(
        dom.rackSelect,
        racks.map((rack) => ({ value: rack.id, label: `${rack.label} · ${rack.sourceWall}` })),
        "Select rack",
    );
    setSelectOptions(dom.levelSelect, [], "Select level");
    setSelectOptions(dom.positionSelect, [], "Select position");
}

function populateLevelOptions(rackId) {
    const rack = WAREHOUSE_CONFIG.racks.find((entry) => entry.id === rackId);
    const levels = rack ? Array.from({ length: rack.levels }, (_, index) => ({
        value: String(index + 1),
        label: `Level ${index + 1}`,
    })) : [];
    setSelectOptions(dom.levelSelect, levels, "Select level");
    setSelectOptions(dom.positionSelect, [], "Select position");
}

function populatePositionOptions(rackId, level) {
    const rack = WAREHOUSE_CONFIG.racks.find((entry) => entry.id === rackId);
    const positions = rack ? Array.from({ length: rack.bays }, (_, index) => {
        const id = locationIdFor(rack, Number(level), index);
        const location = WAREHOUSE_LOCATIONS[id];
        return { value: location.position, label: `Position ${location.position} · ${location.zone}` };
    }) : [];
    setSelectOptions(dom.positionSelect, positions, "Select position");
}

function renderSearchOptions() {
    if (!dom.searchOptions) return;
    dom.searchOptions.innerHTML = runtime.inventory.map((item) => (
        `<option value="${escapeHtml(itemSearchLabel(item))}"></option>`
    )).join("");
}

async function loadInventory() {
    const bridge = window.OnlyStockWarehouseBridge;
    if (!bridge?.getUser?.()) return;
    try {
        const items = await bridge.loadAllInventory();
        runtime.inventory = Array.isArray(items) ? items : [];
        renderSearchOptions();
        setMessage(`${runtime.inventory.length} inventory item(s) available for warehouse search.`);
    } catch (error) {
        runtime.inventory = bridge.getInventory?.() || [];
        renderSearchOptions();
        setMessage(error.message || "Unable to load inventory for warehouse search.", "error");
    }
}

function findInventoryItem(query) {
    const normalizedQuery = normalizeText(query);
    if (!normalizedQuery) return null;
    const exact = runtime.inventory.find((item) => normalizeText(itemSearchLabel(item)) === normalizedQuery);
    if (exact) return exact;
    const matches = runtime.inventory
        .map((item) => {
            const text = itemSearchText(item);
            let score = text.startsWith(normalizedQuery) ? 3 : 0;
            if (normalizeText(itemName(item)).includes(normalizedQuery)) score += 2;
            if (text.includes(normalizedQuery)) score += 1;
            return { item, score };
        })
        .filter((entry) => entry.score > 0)
        .sort((left, right) => right.score - left.score);
    return matches[0]?.item || null;
}

function detailRows(rows) {
    return `<dl class="warehouse-detail-grid">${rows
        .filter(([, value]) => !isBlankValue(value))
        .map(([label, value]) => `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`)
        .join("")}</dl>`;
}

function stockCards(items, heading = "Current inventory") {
    if (!items.length) return `<p class="warehouse-resolution-note">No inventory record is assigned to this exact selection.</p>`;
    return `
        <div class="warehouse-stock-list">
            <strong>${escapeHtml(heading)}</strong>
            ${items.slice(0, 12).map((item) => `
                <div class="warehouse-stock-card">
                    <strong>${escapeHtml(itemName(item))}</strong>
                    <span>${escapeHtml(item.category)} · ${escapeHtml(formatNumber(item.quantity))} ${escapeHtml(formatUnit(item.unit))}</span>
                </div>
            `).join("")}
            ${items.length > 12 ? `<span>${items.length - 12} more item(s)</span>` : ""}
        </div>
    `;
}

function setBadge(text, kind = "") {
    dom.badge.textContent = text;
    dom.badge.className = `warehouse-location-badge${kind ? ` is-${kind}` : ""}`;
}

function renderItemDetails(item, resolution) {
    const firstLocation = resolution.ids.map((id) => WAREHOUSE_LOCATIONS[id]).find(Boolean);
    const unit = formatUnit(item.unit);
    const dimensions = [
        displayValue(item.actual_width || item.width, item.width_unit ? ` ${item.width_unit}` : ""),
        displayValue(item.length || item.height, item.length_unit ? ` ${item.length_unit}` : ""),
    ].filter(Boolean).join(" × ");
    const locationRows = firstLocation ? [
        ["Warehouse", firstLocation.warehouse],
        ["Zone", firstLocation.zone],
        ["Rack", resolution.resolution === "exact" ? firstLocation.rack : "PDF-defined zone"],
        ["Level", resolution.resolution === "exact" ? firstLocation.level : "Not assigned"],
        ["Position", resolution.resolution === "exact" ? firstLocation.position : "Not assigned"],
        ["Location ID", resolution.resolution === "exact" ? firstLocation.location_id : "Pending assignment"],
    ] : [["Location", "Not assigned"]];

    dom.details.innerHTML = `
        <h4 class="warehouse-detail-title">${escapeHtml(itemName(item))}</h4>
        ${detailRows([
            ["Category", item.category],
            ["Storage", item.storage_type],
            ["Thickness", displayValue(item.thickness, item.thickness_unit ? ` ${item.thickness_unit}` : "")],
            ["Dimensions", dimensions],
            ["Roll / Batch", item.roll_no || item.batch_no || item.batch_roll_no],
            ["Stock", `${formatNumber(item.quantity)} ${unit}`.trim()],
            ["Rolls", item.number_of_rolls],
            ["Sheets", item.number_of_sheets],
        ])}
        <h4 class="warehouse-detail-title" style="margin-top:1rem">Location</h4>
        ${detailRows(locationRows)}
        ${resolution.resolution === "zone" ? `
            <p class="warehouse-resolution-note">This existing record has no <code>location_id</code>. The highlighted shelves are the general ${escapeHtml(firstLocation?.zone || "PDF")} zone shown in the source drawing; no exact bay is being claimed.</p>
        ` : ""}
        ${resolution.resolution === "unassigned" ? `
            <p class="warehouse-resolution-note">This category has no exact location and no matching storage zone in the supplied PDF.</p>
        ` : ""}
    `;
    setBadge(
        resolution.resolution === "exact" ? `${resolution.ids.length} exact location${resolution.ids.length === 1 ? "" : "s"}` : resolution.resolution === "zone" ? "Zone mapped" : "Unassigned",
        resolution.resolution === "exact" ? "exact" : resolution.resolution === "zone" ? "zone" : "",
    );
    dom.stockDetails.disabled = false;
    dom.locateAgain.disabled = resolution.ids.length === 0;
}

function renderLocationDetails(locationId) {
    const location = WAREHOUSE_LOCATIONS[locationId];
    if (!location) return;
    const exactItems = exactItemsAtLocation(locationId);
    const zoneItems = unassignedItemsForZone(location.zone_key);
    dom.details.innerHTML = `
        <h4 class="warehouse-detail-title">${escapeHtml(location.rack)} · ${escapeHtml(location.level)} · Position ${escapeHtml(location.position)}</h4>
        ${detailRows([
            ["Location", location.location_id],
            ["Warehouse", location.warehouse],
            ["Zone", location.zone],
            ["Rack", location.rack],
            ["Level", location.level],
            ["Position", location.position],
            ["PDF source", `Page ${location.source_page} · ${location.source_wall}`],
            ["Reference note", location.note],
        ])}
        ${stockCards(exactItems, "Exactly assigned inventory")}
        ${zoneItems.length ? stockCards(zoneItems, "Zone inventory awaiting exact shelf assignment") : ""}
    `;
    setBadge(location.location_id, "exact");
    runtime.selectedItem = exactItems[0] || null;
    dom.stockDetails.disabled = !runtime.selectedItem;
    dom.locateAgain.disabled = false;
}

function renderRackDetails(rackId) {
    const rack = WAREHOUSE_CONFIG.racks.find((entry) => entry.id === rackId);
    if (!rack) return;
    const locations = Object.values(WAREHOUSE_LOCATIONS).filter((location) => location.rack_id === rackId);
    const zones = [...new Set(locations.filter((location) => location.zone_key !== "open").map((location) => location.zone))];
    const items = itemsForRack(rackId);
    dom.details.innerHTML = `
        <h4 class="warehouse-detail-title">Rack ${escapeHtml(rack.label)}</h4>
        ${detailRows([
            ["Zones", zones.join(", ")],
            ["Levels", rack.levels],
            ["Positions / level", rack.bays],
            ["PDF source", `Page ${rack.sourcePage} · ${rack.sourceWall}`],
            ["Location IDs", `${locations.length} configurable positions`],
        ])}
        ${stockCards(items, "Current inventory for rack zones")}
    `;
    setBadge(`${rack.label} selected`);
    runtime.selectedItem = items[0] || null;
    dom.stockDetails.disabled = !runtime.selectedItem;
    dom.locateAgain.disabled = false;
}

function renderZoneDetails(zoneKey) {
    const zone = WAREHOUSE_CONFIG.zones[zoneKey];
    if (!zone) return;
    const locations = locationsForZone(zoneKey);
    const racks = [...new Set(locations.map((location) => location.rack))];
    const exactItems = runtime.inventory.filter((item) => exactLocationIdsForItem(item).some((id) => WAREHOUSE_LOCATIONS[id]?.zone_key === zoneKey));
    const unassigned = unassignedItemsForZone(zoneKey);
    dom.details.innerHTML = `
        <h4 class="warehouse-detail-title">${escapeHtml(zone.name)}</h4>
        ${detailRows([
            ["Warehouse", WAREHOUSE_CONFIG.warehouse.name],
            ["Racks", racks.join(", ")],
            ["Storage positions", locations.length],
            ["Source", "Rack levels identified in the supplied PDF"],
        ])}
        ${stockCards(exactItems, "Exactly assigned inventory")}
        ${unassigned.length ? stockCards(unassigned, "Zone inventory awaiting exact shelf assignment") : ""}
    `;
    setBadge("Zone selected", "zone");
    runtime.selectedItem = exactItems[0] || unassigned[0] || null;
    dom.stockDetails.disabled = !runtime.selectedItem;
    dom.locateAgain.disabled = locations.length === 0;
}

function clear3dHighlights() {
    runtime.locationMeshes.forEach((mesh) => {
        mesh.material.emissive?.setHex(0x000000);
        mesh.material.emissiveIntensity = 0;
        mesh.material.opacity = mesh.userData.baseOpacity;
    });
    if (runtime.highlightMarker) runtime.highlightMarker.visible = false;
}

function update3dHighlights() {
    if (!runtime.threeAvailable) return;
    clear3dHighlights();
    const points = [];
    runtime.selectedLocationIds.forEach((locationId) => {
        const mesh = runtime.locationMeshes.get(locationId);
        const location = WAREHOUSE_LOCATIONS[locationId];
        if (mesh) {
            mesh.material.emissive.setHex(0xffcf35);
            mesh.material.emissiveIntensity = 0.9;
            mesh.material.opacity = Math.max(0.58, mesh.userData.baseOpacity);
        }
        if (location) points.push(location);
    });
    if (points.length && runtime.highlightMarker) {
        const center = averageLocation(points);
        runtime.highlightMarker.position.set(center.x, center.y + 1.1, center.z);
        runtime.highlightMarker.visible = true;
    }
}

function setSelection(locationIds, { rackId = null, resolution = null } = {}) {
    runtime.selectedLocationIds = new Set(locationIds.filter((id) => WAREHOUSE_LOCATIONS[id]));
    runtime.selectedRackId = rackId || WAREHOUSE_LOCATIONS[locationIds[0]]?.rack_id || null;
    runtime.selectedResolution = resolution;
    update3dHighlights();
    updateAllPlanHighlights();
}

function averageLocation(locations) {
    if (!locations.length) return { x: 0, y: 1.5, z: 1 };
    return locations.reduce((center, location) => ({
        x: center.x + location.x / locations.length,
        y: center.y + location.y / locations.length,
        z: center.z + location.z / locations.length,
    }), { x: 0, y: 0, z: 0 });
}

function animateCameraToLocations(locationIds, broad = false) {
    if (!runtime.threeAvailable || !locationIds.length) return;
    const points = locationIds.map((id) => WAREHOUSE_LOCATIONS[id]).filter(Boolean);
    const target = averageLocation(points);
    const distance = broad || points.length > 4 ? 15 : 8.5;
    const direction = runtime.camera.position.clone().sub(runtime.controls.target).normalize();
    if (Math.abs(direction.y) < 0.25) direction.y = 0.42;
    direction.normalize();
    const endTarget = new runtime.THREE.Vector3(target.x, Math.max(1.25, target.y), target.z);
    const endPosition = endTarget.clone().add(direction.multiplyScalar(distance));
    runtime.cameraTween = {
        startedAt: performance.now(),
        duration: 1050,
        fromPosition: runtime.camera.position.clone(),
        fromTarget: runtime.controls.target.clone(),
        endPosition,
        endTarget,
    };
}

function selectLocation(locationId, animate = true) {
    const location = WAREHOUSE_LOCATIONS[locationId];
    if (!location) return;
    runtime.selectedItem = null;
    setSelection([locationId], { rackId: location.rack_id, resolution: "exact" });
    renderLocationDetails(locationId);
    syncNavigationToLocation(location);
    if (animate) {
        setMode("3d");
        animateCameraToLocations([locationId]);
    }
}

function selectRack(rackId, animate = true) {
    const ids = Object.values(WAREHOUSE_LOCATIONS)
        .filter((location) => location.rack_id === rackId)
        .map((location) => location.location_id);
    runtime.selectedItem = null;
    setSelection(ids, { rackId, resolution: "rack" });
    renderRackDetails(rackId);
    if (animate) {
        setMode("3d");
        animateCameraToLocations(ids, true);
    }
}

function selectZone(zoneKey, animate = true) {
    const ids = locationsForZone(zoneKey).map((location) => location.location_id);
    runtime.selectedItem = null;
    setSelection(ids, { resolution: "zone" });
    renderZoneDetails(zoneKey);
    if (animate && ids.length) {
        setMode("3d");
        animateCameraToLocations(ids, true);
    }
}

function locateItem(item) {
    if (!item) return;
    const resolution = resolveItemLocations(item);
    runtime.selectedItem = item;
    setSelection(resolution.ids, { resolution: resolution.resolution });
    renderItemDetails(item, resolution);
    dom.search.value = itemSearchLabel(item);
    if (resolution.ids.length) {
        setMode("3d");
        animateCameraToLocations(resolution.ids, resolution.resolution !== "exact");
        const location = WAREHOUSE_LOCATIONS[resolution.ids[0]];
        if (resolution.resolution === "exact" && location) syncNavigationToLocation(location);
        setMessage(
            resolution.resolution === "exact"
                ? `Moving to ${resolution.ids.join(", ")}.`
                : `Showing the PDF-defined ${WAREHOUSE_CONFIG.zones[resolution.zoneKey].name} zone. Assign location_id for an exact bay.`,
            resolution.resolution === "exact" ? "success" : "",
        );
    } else {
        setMessage("This item has no location_id and its category is not assigned in the supplied warehouse plan.", "error");
    }
}

function syncNavigationToLocation(location) {
    dom.zoneSelect.value = location.zone_key === "open" ? "" : location.zone_key;
    filterRackOptionsForZone(dom.zoneSelect.value);
    dom.rackSelect.value = location.rack_id;
    populateLevelOptions(location.rack_id);
    dom.levelSelect.value = String(location.level_number);
    populatePositionOptions(location.rack_id, location.level_number);
    dom.positionSelect.value = location.position;
}

function resetDetails() {
    runtime.selectedItem = null;
    runtime.selectedRackId = null;
    runtime.selectedLocationIds.clear();
    runtime.selectedResolution = null;
    update3dHighlights();
    updateAllPlanHighlights();
    dom.details.innerHTML = `
        <div class="warehouse-empty-details">
            <strong>Select a rack, shelf, or inventory item</strong>
            <p>The model uses the rack arrangement and category levels shown across all three PDF pages.</p>
        </div>
    `;
    setBadge("No selection");
    dom.stockDetails.disabled = true;
    dom.locateAgain.disabled = true;
}

function setMode(mode) {
    runtime.mode = mode;
    const is3d = mode === "3d";
    dom.view3d.classList.toggle("is-active", is3d);
    dom.viewTop.classList.toggle("is-active", !is3d);
    dom.button3d.setAttribute("aria-pressed", String(is3d));
    dom.buttonTop.setAttribute("aria-pressed", String(!is3d));
    dom.viewStatus.textContent = is3d ? "Complete warehouse · 3D isometric" : "Complete warehouse · 2D top view";
    if (is3d) {
        ensureThreeScene();
        window.setTimeout(resizeRenderer, 0);
    }
}

function resetCamera() {
    runtime.topScale = 1;
    dom.topSvg.style.transform = "scale(1)";
    if (!runtime.threeAvailable) return;
    const { position, target } = WAREHOUSE_CONFIG.camera;
    runtime.cameraTween = {
        startedAt: performance.now(),
        duration: 850,
        fromPosition: runtime.camera.position.clone(),
        fromTarget: runtime.controls.target.clone(),
        endPosition: new runtime.THREE.Vector3(...position),
        endTarget: new runtime.THREE.Vector3(...target),
    };
}

function zoom(factor) {
    if (runtime.mode === "2d") {
        runtime.topScale = Math.min(2.2, Math.max(0.8, runtime.topScale * factor));
        dom.topSvg.style.transform = `scale(${runtime.topScale})`;
        return;
    }
    if (!runtime.threeAvailable) return;
    const offset = runtime.camera.position.clone().sub(runtime.controls.target).multiplyScalar(factor);
    const distance = offset.length();
    if (distance < WAREHOUSE_CONFIG.camera.minDistance || distance > WAREHOUSE_CONFIG.camera.maxDistance) return;
    runtime.camera.position.copy(runtime.controls.target.clone().add(offset));
    runtime.controls.update();
}

function showAll() {
    resetDetails();
    resetCamera();
    setMessage("Showing the complete PDF-derived warehouse footprint.");
}

function geometryKey(width, height, depth) {
    return `${width.toFixed(3)}:${height.toFixed(3)}:${depth.toFixed(3)}`;
}

function createBoxFactory(THREE) {
    const cache = new Map();
    return (width, height, depth) => {
        const key = geometryKey(width, height, depth);
        if (!cache.has(key)) cache.set(key, new THREE.BoxGeometry(width, height, depth));
        return cache.get(key);
    };
}

function addBox(sceneOrGroup, boxGeometry, material, options) {
    const mesh = new runtime.THREE.Mesh(boxGeometry(options.width, options.height, options.depth), material);
    mesh.position.set(options.x, options.y, options.z);
    if (options.rotation) mesh.rotation.y = options.rotation;
    mesh.castShadow = Boolean(options.castShadow);
    mesh.receiveShadow = Boolean(options.receiveShadow);
    if (options.userData) Object.assign(mesh.userData, options.userData);
    sceneOrGroup.add(mesh);
    return mesh;
}

function createLabelSprite(text, color = "#eff9f7") {
    const THREE = runtime.THREE;
    const canvas = document.createElement("canvas");
    canvas.width = 384;
    canvas.height = 92;
    const context = canvas.getContext("2d");
    context.fillStyle = "rgba(7, 30, 38, 0.88)";
    context.roundRect(4, 4, 376, 84, 18);
    context.fill();
    context.strokeStyle = "rgba(255,255,255,0.28)";
    context.lineWidth = 3;
    context.stroke();
    context.fillStyle = color;
    context.font = "700 30px Segoe UI, sans-serif";
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText(text, 192, 47, 350);
    const texture = new THREE.CanvasTexture(canvas);
    texture.colorSpace = THREE.SRGBColorSpace;
    const material = new THREE.SpriteMaterial({ map: texture, transparent: true, depthTest: false });
    const sprite = new THREE.Sprite(material);
    sprite.scale.set(4.2, 1.0, 1);
    sprite.renderOrder = 5;
    return sprite;
}

function buildFloorAndWalls(boxGeometry) {
    const THREE = runtime.THREE;
    const floorShape = new THREE.Shape();
    WAREHOUSE_CONFIG.footprint.forEach(([x, z], index) => {
        if (index === 0) floorShape.moveTo(x, -z);
        else floorShape.lineTo(x, -z);
    });
    floorShape.closePath();
    const floor = new THREE.Mesh(
        new THREE.ShapeGeometry(floorShape),
        new THREE.MeshStandardMaterial({ color: 0xcfd6d0, roughness: 0.96, metalness: 0 }),
    );
    floor.rotation.x = -Math.PI / 2;
    floor.position.y = -0.04;
    floor.receiveShadow = true;
    runtime.scene.add(floor);

    const wallMaterial = new THREE.MeshStandardMaterial({ color: 0xe9e3d7, roughness: 0.86 });
    const wallSegments = [
        [[-18, -10], [12, -10]],
        [[12, -10], [12, 14]],
        [[12, 14], [WAREHOUSE_CONFIG.entrance.openingTo, 14]],
        [[WAREHOUSE_CONFIG.entrance.openingFrom, 14], [8, 14]],
        [[8, 14], [8, 8]],
        [[8, 8], [-18, 8]],
        [[-18, 8], [-18, -10]],
    ];
    wallSegments.forEach(([[x1, z1], [x2, z2]]) => {
        const dx = x2 - x1;
        const dz = z2 - z1;
        const length = Math.hypot(dx, dz);
        addBox(runtime.scene, boxGeometry, wallMaterial, {
            width: length,
            height: WAREHOUSE_CONFIG.wallHeight,
            depth: WAREHOUSE_CONFIG.wallThickness,
            x: (x1 + x2) / 2,
            y: WAREHOUSE_CONFIG.wallHeight / 2,
            z: (z1 + z2) / 2,
            rotation: -Math.atan2(dz, dx),
            receiveShadow: true,
        });
    });

    const threshold = addBox(runtime.scene, boxGeometry, new THREE.MeshStandardMaterial({ color: 0x0d6c74 }), {
        width: WAREHOUSE_CONFIG.entrance.openingTo - WAREHOUSE_CONFIG.entrance.openingFrom,
        height: 0.06,
        depth: 0.36,
        x: 10,
        y: 0.01,
        z: 14,
    });
    threshold.receiveShadow = true;
    const entranceLabel = createLabelSprite("ENTRANCE", "#c7fbf3");
    entranceLabel.position.set(10, 1.65, 13.75);
    entranceLabel.scale.set(3.3, 0.78, 1);
    runtime.scene.add(entranceLabel);
}

function buildFixtures(boxGeometry) {
    const THREE = runtime.THREE;
    const fixtureMaterial = new THREE.MeshStandardMaterial({ color: 0xf0e7d7, roughness: 0.93 });
    const dividerMaterial = new THREE.MeshStandardMaterial({ color: 0x7f8d8c, roughness: 0.9 });
    WAREHOUSE_CONFIG.fixtures.forEach((fixture) => {
        addBox(runtime.scene, boxGeometry, fixture.kind === "divider" ? dividerMaterial : fixtureMaterial, {
            width: fixture.width,
            height: fixture.height,
            depth: fixture.depth,
            x: fixture.x,
            y: fixture.height / 2,
            z: fixture.z,
            receiveShadow: true,
        });
    });

    const openLabel = createLabelSprite("OPEN WORKING / AISLE AREA", "#d8e8e4");
    openLabel.position.set(1.1, 0.16, 1.4);
    openLabel.rotation.x = -Math.PI / 2;
    openLabel.scale.set(6.2, 1.1, 1);
    runtime.scene.add(openLabel);
}

function createRack(rack, boxGeometry, materials) {
    const THREE = runtime.THREE;
    const group = new THREE.Group();
    group.position.set(rack.x, 0, rack.z);
    group.rotation.y = rack.rotation;
    group.userData = { type: "rack", rackId: rack.id };
    const totalWidth = rack.bays * rack.bayWidth;
    const rackHeight = 0.42 + rack.levels * rack.levelHeight;
    const postWidth = 0.13;
    const shelfY = (level) => 0.28 + (level - 1) * rack.levelHeight;

    for (let boundary = 0; boundary <= rack.bays; boundary += 1) {
        const x = -totalWidth / 2 + boundary * rack.bayWidth;
        [-rack.depth / 2, rack.depth / 2].forEach((z) => {
            const post = addBox(group, boxGeometry, materials.upright, {
                width: postWidth,
                height: rackHeight,
                depth: postWidth,
                x,
                y: rackHeight / 2,
                z,
                castShadow: true,
                userData: { type: "rack", rackId: rack.id },
            });
            runtime.clickables.push(post);
        });
    }

    for (let level = 1; level <= rack.levels; level += 1) {
        const y = shelfY(level);
        [-rack.depth / 2, rack.depth / 2].forEach((z) => {
            const beam = addBox(group, boxGeometry, materials.beam, {
                width: totalWidth + 0.12,
                height: 0.12,
                depth: 0.1,
                x: 0,
                y,
                z,
                castShadow: true,
                userData: { type: "rack", rackId: rack.id },
            });
            runtime.clickables.push(beam);
        });
        addBox(group, boxGeometry, materials.shelf, {
            width: totalWidth,
            height: 0.055,
            depth: rack.depth,
            x: 0,
            y: y + 0.03,
            z: 0,
            receiveShadow: true,
            userData: { type: "rack", rackId: rack.id },
        });

        for (let bayIndex = 0; bayIndex < rack.bays; bayIndex += 1) {
            const locationId = locationIdFor(rack, level, bayIndex);
            const location = WAREHOUSE_LOCATIONS[locationId];
            const zone = WAREHOUSE_CONFIG.zones[location.zone_key];
            const baseOpacity = location.zone_key === "open" ? 0.07 : 0.29;
            const material = new THREE.MeshStandardMaterial({
                color: zone.color,
                transparent: true,
                opacity: baseOpacity,
                roughness: 0.68,
                metalness: 0.04,
                emissive: 0x000000,
                emissiveIntensity: 0,
                depthWrite: false,
            });
            const x = (bayIndex - (rack.bays - 1) / 2) * rack.bayWidth;
            const volume = addBox(group, boxGeometry, material, {
                width: rack.bayWidth * 0.82,
                height: rack.levelHeight * 0.62,
                depth: rack.depth * 0.7,
                x,
                y: y + rack.levelHeight * 0.34,
                z: 0,
                userData: {
                    type: "location",
                    locationId,
                    rackId: rack.id,
                    baseOpacity,
                },
            });
            runtime.locationMeshes.set(locationId, volume);
            runtime.clickables.push(volume);
        }
    }

    const label = createLabelSprite(`${rack.label} · PDF P${rack.sourcePage}`);
    label.position.set(0, rackHeight + 0.72, 0);
    group.add(label);
    runtime.rackGroups.set(rack.id, group);
    runtime.scene.add(group);
}

async function ensureThreeScene() {
    if (runtime.initialized3d || runtime.initializing3d || !dom.canvasHost) return;
    runtime.initializing3d = true;
    try {
        const THREE = await import("three");
        const { OrbitControls } = await import("three/addons/controls/OrbitControls.js");
        runtime.THREE = THREE;
        runtime.scene = new THREE.Scene();
        runtime.scene.background = new THREE.Color(0x0c222c);
        runtime.scene.fog = new THREE.Fog(0x0c222c, 38, 78);
        runtime.camera = new THREE.PerspectiveCamera(42, 1, 0.1, 140);
        runtime.camera.position.set(...WAREHOUSE_CONFIG.camera.position);
        runtime.renderer = new THREE.WebGLRenderer({ antialias: true, powerPreference: "high-performance", alpha: false });
        runtime.renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.7));
        runtime.renderer.outputColorSpace = THREE.SRGBColorSpace;
        runtime.renderer.shadowMap.enabled = true;
        runtime.renderer.shadowMap.type = THREE.PCFSoftShadowMap;
        runtime.renderer.domElement.setAttribute("aria-label", "Drag to orbit, right-drag to pan, and scroll to zoom the warehouse");
        dom.canvasHost.replaceChildren(runtime.renderer.domElement);

        runtime.controls = new OrbitControls(runtime.camera, runtime.renderer.domElement);
        runtime.controls.target.set(...WAREHOUSE_CONFIG.camera.target);
        runtime.controls.enableDamping = true;
        runtime.controls.dampingFactor = 0.075;
        runtime.controls.enablePan = true;
        runtime.controls.screenSpacePanning = true;
        runtime.controls.minDistance = WAREHOUSE_CONFIG.camera.minDistance;
        runtime.controls.maxDistance = WAREHOUSE_CONFIG.camera.maxDistance;
        runtime.controls.maxPolarAngle = Math.PI * 0.49;
        runtime.controls.update();

        runtime.scene.add(new THREE.HemisphereLight(0xd9f5f1, 0x24363c, 2.05));
        const keyLight = new THREE.DirectionalLight(0xfff0d7, 2.25);
        keyLight.position.set(-10, 24, 17);
        keyLight.castShadow = true;
        keyLight.shadow.mapSize.set(1024, 1024);
        keyLight.shadow.camera.left = -26;
        keyLight.shadow.camera.right = 26;
        keyLight.shadow.camera.top = 25;
        keyLight.shadow.camera.bottom = -25;
        runtime.scene.add(keyLight);

        const boxGeometry = createBoxFactory(THREE);
        const materials = {
            upright: new THREE.MeshStandardMaterial({ color: 0x285e91, roughness: 0.52, metalness: 0.42 }),
            beam: new THREE.MeshStandardMaterial({ color: 0xb94a3b, roughness: 0.58, metalness: 0.28 }),
            shelf: new THREE.MeshStandardMaterial({ color: 0x78909a, roughness: 0.65, metalness: 0.32 }),
        };
        buildFloorAndWalls(boxGeometry);
        buildFixtures(boxGeometry);
        WAREHOUSE_CONFIG.racks.forEach((rack) => createRack(rack, boxGeometry, materials));

        runtime.highlightMarker = new THREE.Group();
        const ring = new THREE.Mesh(
            new THREE.TorusGeometry(0.62, 0.08, 10, 32),
            new THREE.MeshBasicMaterial({ color: 0xffdc4a, transparent: true, opacity: 0.92 }),
        );
        ring.rotation.x = Math.PI / 2;
        const beacon = new THREE.Mesh(
            new THREE.SphereGeometry(0.16, 14, 10),
            new THREE.MeshBasicMaterial({ color: 0xfff2a0 }),
        );
        runtime.highlightMarker.add(ring, beacon);
        runtime.highlightMarker.visible = false;
        runtime.scene.add(runtime.highlightMarker);

        runtime.raycaster = new THREE.Raycaster();
        runtime.pointer = new THREE.Vector2();
        runtime.renderer.domElement.addEventListener("pointerdown", handleCanvasPointerDown);
        runtime.renderer.domElement.addEventListener("pointerup", handleCanvasPointerUp);
        runtime.resizeObserver = new ResizeObserver(resizeRenderer);
        runtime.resizeObserver.observe(dom.canvasHost);
        runtime.threeAvailable = true;
        runtime.initialized3d = true;
        runtime.initializing3d = false;
        resizeRenderer();
        update3dHighlights();
        runtime.renderer.setAnimationLoop(renderFrame);
    } catch (error) {
        runtime.initializing3d = false;
        runtime.threeAvailable = false;
        dom.canvasHost.innerHTML = `<div class="warehouse-webgl-error">The 3D library could not load. The PDF-derived 2D Top View remains available.<br>${escapeHtml(error.message || "Three.js unavailable")}</div>`;
        setMode("2d");
        setMessage("3D could not load from the pinned CDN. Use the 2D Top View or check the network connection.", "error");
    }
}

function resizeRenderer() {
    if (!runtime.renderer || !runtime.camera || !dom.canvasHost) return;
    const width = Math.max(1, dom.canvasHost.clientWidth);
    const height = Math.max(1, dom.canvasHost.clientHeight);
    runtime.renderer.setSize(width, height, false);
    runtime.camera.aspect = width / height;
    runtime.camera.updateProjectionMatrix();
}

function handleCanvasPointerDown(event) {
    runtime.pointerDown = { x: event.clientX, y: event.clientY };
}

function handleCanvasPointerUp(event) {
    if (!runtime.pointerDown || !runtime.raycaster) return;
    const distance = Math.hypot(event.clientX - runtime.pointerDown.x, event.clientY - runtime.pointerDown.y);
    runtime.pointerDown = null;
    if (distance > 5) return;
    const rect = runtime.renderer.domElement.getBoundingClientRect();
    runtime.pointer.x = ((event.clientX - rect.left) / rect.width) * 2 - 1;
    runtime.pointer.y = -((event.clientY - rect.top) / rect.height) * 2 + 1;
    runtime.raycaster.setFromCamera(runtime.pointer, runtime.camera);
    const hit = runtime.raycaster.intersectObjects(runtime.clickables, false)[0]?.object;
    if (!hit) return;
    if (hit.userData.type === "location") selectLocation(hit.userData.locationId, false);
    else if (hit.userData.rackId) selectRack(hit.userData.rackId, false);
}

function renderFrame(now) {
    if (!runtime.renderer || !runtime.scene || !runtime.camera) return;
    if (runtime.cameraTween) {
        const elapsed = (now - runtime.cameraTween.startedAt) / runtime.cameraTween.duration;
        const t = Math.min(1, Math.max(0, elapsed));
        const eased = t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
        runtime.camera.position.lerpVectors(runtime.cameraTween.fromPosition, runtime.cameraTween.endPosition, eased);
        runtime.controls.target.lerpVectors(runtime.cameraTween.fromTarget, runtime.cameraTween.endTarget, eased);
        if (t >= 1) runtime.cameraTween = null;
    }
    if (runtime.highlightMarker?.visible) {
        const pulse = 1 + Math.sin(now * 0.006) * 0.16;
        runtime.highlightMarker.scale.setScalar(pulse);
        runtime.locationMeshes.forEach((mesh, locationId) => {
            if (runtime.selectedLocationIds.has(locationId)) {
                mesh.material.emissiveIntensity = 0.75 + Math.sin(now * 0.006) * 0.25;
            }
        });
    }
    runtime.controls.update();
    runtime.renderer.render(runtime.scene, runtime.camera);
}

function handleLocateClick() {
    const item = findInventoryItem(dom.search.value);
    if (!item) {
        setMessage("Choose an existing inventory item from the search suggestions or enter a matching name.", "error");
        return;
    }
    locateItem(item);
}

function handlePlanInteraction(event) {
    const rackNode = event.target.closest?.("[data-rack-id]");
    if (rackNode?.dataset.rackId) selectRack(rackNode.dataset.rackId);
}

function bindEvents() {
    dom.locate.addEventListener("click", handleLocateClick);
    dom.search.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
            event.preventDefault();
            handleLocateClick();
        }
    });
    dom.button3d.addEventListener("click", () => setMode("3d"));
    dom.buttonTop.addEventListener("click", () => setMode("2d"));
    dom.reset.addEventListener("click", resetCamera);
    dom.viewportReset.addEventListener("click", resetCamera);
    dom.showAll.addEventListener("click", showAll);
    dom.zoomIn.addEventListener("click", () => zoom(0.82));
    dom.zoomOut.addEventListener("click", () => zoom(1.22));
    dom.topSvg.addEventListener("click", handlePlanInteraction);
    dom.topSvg.addEventListener("keydown", (event) => {
        if (["Enter", " "].includes(event.key)) handlePlanInteraction(event);
    });
    dom.stockDetails.addEventListener("click", () => {
        if (runtime.selectedItem) window.OnlyStockWarehouseBridge?.viewInventoryItem(runtime.selectedItem.id);
    });
    dom.locateAgain.addEventListener("click", () => {
        if (runtime.selectedItem) locateItem(runtime.selectedItem);
        else if (runtime.selectedLocationIds.size) animateCameraToLocations([...runtime.selectedLocationIds], runtime.selectedLocationIds.size > 4);
    });
    dom.zoneSelect.addEventListener("change", () => {
        const zoneKey = dom.zoneSelect.value;
        filterRackOptionsForZone(zoneKey);
        if (zoneKey) selectZone(zoneKey);
        else showAll();
    });
    dom.rackSelect.addEventListener("change", () => {
        const rackId = dom.rackSelect.value;
        populateLevelOptions(rackId);
        if (rackId) selectRack(rackId);
    });
    dom.levelSelect.addEventListener("change", () => {
        const rackId = dom.rackSelect.value;
        const level = Number(dom.levelSelect.value);
        populatePositionOptions(rackId, level);
        if (!rackId || !level) return;
        const ids = Object.values(WAREHOUSE_LOCATIONS)
            .filter((location) => location.rack_id === rackId && location.level_number === level)
            .map((location) => location.location_id);
        setSelection(ids, { rackId, resolution: "level" });
        renderRackDetails(rackId);
        setMode("3d");
        animateCameraToLocations(ids);
    });
    dom.positionSelect.addEventListener("change", () => {
        const rack = WAREHOUSE_CONFIG.racks.find((entry) => entry.id === dom.rackSelect.value);
        const level = Number(dom.levelSelect.value);
        const bayIndex = dom.positionSelect.value.charCodeAt(0) - 65;
        if (rack && level && bayIndex >= 0) selectLocation(locationIdFor(rack, level, bayIndex));
    });
    window.addEventListener("onlystock:page-changed", (event) => {
        if (event.detail?.page === "warehouse") {
            ensureThreeScene();
            loadInventory();
        }
    });
    window.addEventListener("onlystock:inventory-updated", (event) => {
        if (!Array.isArray(event.detail?.items)) return;
        runtime.inventory = event.detail.items;
        renderSearchOptions();
    });
}

function initializeWarehousePage() {
    if (!dom.page) return;
    renderPlan(dom.topSvg, true);
    renderPlan(dom.miniMap, false);
    renderLegend();
    populateNavigation();
    bindEvents();
    if (window.location.hash === "#warehouse" && window.OnlyStockWarehouseBridge?.getUser?.()) {
        ensureThreeScene();
        loadInventory();
    }
}

initializeWarehousePage();

export { WAREHOUSE_CONFIG, WAREHOUSE_LOCATIONS };
