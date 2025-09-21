/* 
*  Power BI Visual CLI
*  MIT License
*/
"use strict";

import powerbi from "powerbi-visuals-api";
import * as d3 from "d3";
import "./../style/visual.less";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IViewport = powerbi.IViewport;
import PrimitiveValue = powerbi.PrimitiveValue;
import DataView = powerbi.DataView;
import DataViewTable = powerbi.DataViewTable;
import DataViewTableRow = powerbi.DataViewTableRow;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import VisualSelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;

interface FieldValue {
    label: string;
    value: string;
    role: "detail" | "metric";
}

interface OrgChartDatum {
    id: string;
    identity: VisualSelectionId;
    parentId?: string;
    displayName: string;
    title?: string;
    division?: string;          // renamed from “department” for clarity in the UI
    avatarUrl?: string;
    details: FieldValue[];
    metrics: FieldValue[];
    highlight?: boolean;
}

interface ChartNode {
    id: string;
    payload: OrgChartDatum | null;
    children: ChartNode[];
    totalChildCount: number; // direct children count
}

interface RenderNode {
    payload: OrgChartDatum;
    hasChildren: boolean;
    isCollapsed: boolean;
    x: number;
    y: number;
    centerX: number;
    centerY: number;
    width: number;
    height: number;
    node: d3.HierarchyPointNode<ChartNode>;
}

interface RenderLink {
    source: RenderNode;
    target: RenderNode;
    key: string;
}

interface NodePosition {
    x: number;
    y: number;
    centerX: number;
    centerY: number;
    absX: number;
    absY: number;
    absCenterX: number;
    absCenterY: number;
    width: number;
    height: number;
}

interface TransformResult {
    tree: ChartNode | null;
    nodesById: Map<string, ChartNode>;
    warnings: string[];
}

type Orientation = "vertical" | "horizontal";
type ToolbarIcon = "menu" | "layout" | "expand" | "collapse" | "fit" | "reset";

const DEFAULT_LAYOUT: Orientation = "vertical";
const ZOOM_MIN = 0.25;
const ZOOM_MAX = 4;

export class Visual implements IVisual {
    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;

    private root: HTMLDivElement;

    // Dense rail-only toolbar
    private toolbar: HTMLDivElement;
    private toolbarRail: HTMLDivElement;
    private toolbarToggle: HTMLButtonElement;
    private controlsVisible: boolean = true;

    private canvas: HTMLDivElement;
    private svgElement: SVGSVGElement;
    private svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private zoomRoot: d3.Selection<SVGGElement, unknown, null, undefined>;
    private contentRoot: d3.Selection<SVGGElement, unknown, null, undefined>;
    private linksGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    private nodesGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    private zoomBehavior: d3.ZoomBehavior<SVGSVGElement, unknown>;

    private viewport: IViewport = { width: 0, height: 0 };
    private orientation: Orientation = DEFAULT_LAYOUT;
    private allowZoom: boolean = true;
    private showToolbar: boolean = true;
    private fullTree: ChartNode | null = null;
    private nodesById: Map<string, ChartNode> = new Map();
    private collapsedNodes: Set<string> = new Set();
    private nodePositions: Map<string, NodePosition> = new Map();
    private layoutBounds: { minX: number; maxX: number; minY: number; maxY: number; baseX: number; baseY: number; padding: number; } | null = null;
    private currentTransform: d3.ZoomTransform = d3.zoomIdentity;
    private selectedKeys: Set<string> = new Set();
    private identityToNodeId: Map<string, string> = new Map();
    private nameToNodeId: Map<string, string> = new Map();

    private emptyState: HTMLDivElement;

    // init flags
    private didInitialAutoCollapse = false;
    private didInitialFit = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();

        this.root = document.createElement("div");
        this.root.className = "orgchart-visual";
        this.root.dataset.toolbarVisible = "1";

        // ===== Toolbar (dense) =====
        this.toolbar = document.createElement("div");
        this.toolbar.className = "orgchart__toolbar";
        this.root.appendChild(this.toolbar);

        this.toolbarRail = document.createElement("div");
        this.toolbarRail.className = "orgchart__toolbar-rail";
        Object.assign(this.toolbarRail.style, {
            width: "30px",
            padding: "2px",
            borderRadius: "12px",
            boxShadow: "0 4px 12px rgba(15,23,42,0.10)",
            backdropFilter: "blur(4px)",
            lineHeight: "0"
        } as CSSStyleDeclaration);
        this.toolbar.appendChild(this.toolbarRail);

        // ===== Canvas + SVG =====
        this.canvas = document.createElement("div");
        this.canvas.className = "orgchart__canvas";
        this.root.appendChild(this.canvas);

        this.svgElement = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        this.svgElement.classList.add("orgchart__svg");
        this.canvas.appendChild(this.svgElement);

        this.svg = d3.select(this.svgElement);
        this.zoomBehavior = d3.zoom<SVGSVGElement, unknown>()
            .scaleExtent([ZOOM_MIN, ZOOM_MAX])
            .on("zoom", (event) => this.onZoom(event))
            .filter((event) => {
                if (event.type === 'wheel') return true;
                return !event.ctrlKey && !event.button;
            });

        this.zoomRoot = this.svg.append("g").attr("class", "orgchart__zoom-root");
        this.contentRoot = this.zoomRoot.append("g").attr("class", "orgchart__content");
        this.linksGroup = this.contentRoot.append("g").attr("class", "orgchart__links");
        this.nodesGroup = this.contentRoot.append("g").attr("class", "orgchart__nodes");

        this.svg.on("click", (event) => {
            if (event.target === this.svgElement) {
                this.selectionManager.clear().catch(() => undefined);
                this.selectedKeys.clear();
                this.updateSelectionVisuals();
            }
        });

        this.emptyState = document.createElement("div");
        this.emptyState.className = "orgchart__empty";
        this.emptyState.textContent = "Add employee and manager fields to build the organisation chart.";
        this.canvas.appendChild(this.emptyState);

        this.buildToolbar();
        this.updateToolbarRailState();

        options.element.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        this.viewport = options.viewport;

        const dataView: DataView | undefined = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);

        this.applyFormatting();

        const transformResult = this.transform(dataView);
        this.fullTree = transformResult.tree;
        this.nodesById = transformResult.nodesById;
if (!this.fullTree) {
    // clear scene and show empty state
    this.linksGroup.selectAll("path").remove();
    this.nodesGroup.selectAll("g").remove();
    if (this.emptyState) this.emptyState.style.display = "flex";
    // reset the initial-fit flag so we will fit when data becomes complete
    this.didInitialFit = false;
    return;
}        
        this.pruneCollapsedNodes();

        // Default: only levels 1 & 2 expanded (parents at depth==2 collapsed)
        if (this.fullTree && (!this.didInitialAutoCollapse || this.collapsedNodes.size === 0)) {
            this.collapsedNodes = this.computeDefaultCollapsedSet(2);
            this.didInitialAutoCollapse = true;
        }

        this.updateHighlights(dataView);
        this.render();

        if (!this.didInitialFit) {
            this.fitToViewport(0);
            this.didInitialFit = true;
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ================= Toolbar (rail-only, dense) =================

    private buildToolbar(): void {
        if (!this.toolbarRail) return;
        while (this.toolbarRail.firstChild) this.toolbarRail.removeChild(this.toolbarRail.firstChild);

        // Toggle
        this.toolbarToggle = this.createToolbarRailButton("menu", "Show/Hide controls", () => {
            this.controlsVisible = !this.controlsVisible;
            this.updateToolbarRailState();
        }, 24, 1);
        this.toolbarToggle.classList.add("orgchart__toolbar-btn--toggle");
        this.toolbarRail.appendChild(this.toolbarToggle);

        // Actions
        const buttons: Array<{ title: string; icon: ToolbarIcon; onClick: () => void; }> = [
            { title: "Toggle vertical / horizontal", icon: "layout", onClick: () => this.toggleOrientation() },
            { title: "Expand all nodes", icon: "expand", onClick: () => this.expandAll() },
            { title: "Collapse to root", icon: "collapse", onClick: () => this.collapseAll() },
            { title: "Fit to view", icon: "fit", onClick: () => this.fitToViewport(100) },
            //{ title: "Reset zoom/pan", icon: "reset", onClick: () => this.resetZoom() }
        ];

        buttons.forEach((config) => {
            const btn = this.createToolbarRailButton(config.icon, config.title, config.onClick, 24, 1);
            btn.classList.add("orgchart__toolbar-btn--action");
            this.toolbarRail.appendChild(btn);
        });
    }

    private updateToolbarRailState(): void {
        if (!this.toolbarRail) return;
        const children = Array.from(this.toolbarRail.children) as HTMLElement[];
        children.forEach((el) => {
            if (el === this.toolbarToggle) return;
            el.style.display = this.controlsVisible ? "inline-flex" : "none";
        });

        const label = this.controlsVisible ? "Hide controls" : "Show controls";
        this.toolbarToggle.setAttribute("aria-expanded", this.controlsVisible ? "true" : "false");
        this.toolbarToggle.setAttribute("aria-label", label);
        this.toolbarToggle.title = label;

        if (this.root) this.root.dataset.toolbarMenu = this.controlsVisible ? "1" : "0";
    }

    private createToolbarRailButton(
        icon: ToolbarIcon,
        title: string,
        onClick: () => void,
        size: number = 24,
        margin: number = 1
    ): HTMLButtonElement {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "orgchart__toolbar-btn";
        button.title = title;
        button.setAttribute("aria-label", title);

        Object.assign(button.style, {
            width: `${size}px`,
            height: `${size}px`,
            borderRadius: "8px",
            margin: `${margin}px`,
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            padding: "0"
        } as CSSStyleDeclaration);

        const iconElement = this.createToolbarIcon(icon);
        iconElement.classList.add("orgchart__toolbar-icon");
        const iconSize = Math.max(14, size - 6);
        iconElement.setAttribute("width", `${iconSize}`);
        iconElement.setAttribute("height", `${iconSize}`);
        button.appendChild(iconElement);

        button.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            onClick();
        });
        return button;
    }

    private createToolbarIcon(icon: ToolbarIcon): SVGSVGElement {
        const svgNamespace = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(svgNamespace, "svg");
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("focusable", "false");
        svg.setAttribute("aria-hidden", "true");

        const applyStroke = <T extends SVGElement>(element: T): T => {
            element.setAttribute("stroke", "currentColor");
            element.setAttribute("stroke-width", "1.4");
            element.setAttribute("stroke-linecap", "round");
            element.setAttribute("stroke-linejoin", "round");
            return element;
        };

        const append = <T extends SVGElement>(element: T): T => { svg.appendChild(element); return element; };

        const createLine = (x1: number, y1: number, x2: number, y2: number): SVGLineElement => {
            const line = document.createElementNS(svgNamespace, "line");
            line.setAttribute("x1", x1.toString());
            line.setAttribute("y1", y1.toString());
            line.setAttribute("x2", x2.toString());
            line.setAttribute("y2", y2.toString());
            return append(applyStroke(line));
        };

        const createCircle = (cx: number, cy: number, r: number): SVGCircleElement => {
            const circle = document.createElementNS(svgNamespace, "circle");
            circle.setAttribute("cx", cx.toString());
            circle.setAttribute("cy", cy.toString());
            circle.setAttribute("r", r.toString());
            circle.setAttribute("fill", "none");
            return append(applyStroke(circle));
        };

        const createRect = (x: number, y: number, width: number, height: number, radius: number = 2): SVGRectElement => {
            const rect = document.createElementNS(svgNamespace, "rect");
            rect.setAttribute("x", x.toString());
            rect.setAttribute("y", y.toString());
            rect.setAttribute("width", width.toString());
            rect.setAttribute("height", height.toString());
            rect.setAttribute("rx", radius.toString());
            rect.setAttribute("fill", "none");
            return append(applyStroke(rect));
        };

        const createPath = (d: string): SVGPathElement => {
            const path = document.createElementNS(svgNamespace, "path");
            path.setAttribute("d", d);
            path.setAttribute("fill", "none");
            return append(applyStroke(path));
        };

        const createPolyline = (points: string): SVGPolylineElement => {
            const polyline = document.createElementNS(svgNamespace, "polyline");
            polyline.setAttribute("points", points);
            polyline.setAttribute("fill", "none");
            return append(applyStroke(polyline));
        };

        switch (icon) {
            case "menu": { [7, 12, 17].forEach((y) => createLine(5, y, 19, y)); break; }
            case "layout": { createRect(4.5, 4, 7.5, 16); createRect(13.5, 8, 7.5, 12); break; }
            case "expand": { createCircle(12, 12, 8); createLine(12, 8.5, 12, 15.5); createLine(8.5, 12, 15.5, 12); break; }
            case "collapse": { createCircle(12, 12, 8); createLine(8.5, 12, 15.5, 12); break; }
            case "fit": { createCircle(11, 11, 5.5); createLine(15, 15, 20, 20); break; }
            case "reset": { createPath("M20 12a8 8 0 1 1-2.34-5.66"); createPolyline("20,6 20,12 14,12"); break; }
        }
        return svg;
    }

    // ================= Formatting / data =================

    private applyFormatting(): void {
        if (!this.formattingSettings) return;

        const layout = this.formattingSettings.layout;
        const card = this.formattingSettings.card;
        const labels = this.formattingSettings.labels;
        const links = this.formattingSettings.links;

        this.allowZoom = layout.enableZoom.value;
        this.showToolbar = layout.showToolbar.value;

        const nameSize = this.clampNumber(labels.nameTextSize.value, 8, 32, 16);
        const titleSize = this.clampNumber(labels.titleTextSize.value, 8, 28, 12);
        const detailSize = this.clampNumber(labels.detailTextSize.value, 6, 24, 11);
        const linkWidth = this.clampNumber(links.linkWidth.value, 0.5, 6, 1.4);
        const linkOpacity = this.clampNumber(links.linkOpacity.value, 0.1, 1, 0.9);
        const borderWidth = this.clampNumber(card.borderWidth.value, 0, 8, 1);
        const borderRadius = this.clampNumber(card.borderRadius.value, 0, 50, 18);
        const alignment = (card.cardAlignment.value?.value as string) || "start";

        this.root.style.setProperty("--org-card-bg", card.backgroundColor.value.value || "#ffffff");
        this.root.style.setProperty("--org-card-accent", card.accentColor.value.value || "#2563eb");
        this.root.style.setProperty("--org-card-border", card.borderColor.value.value || "#cbd5f5");
        this.root.style.setProperty("--org-card-text", card.textColor.value.value || "#0f172a");
        this.root.style.setProperty("--org-link-color", links.linkColor.value.value || "#94a3b8");
        this.root.style.setProperty("--org-link-width", `${linkWidth}`);
        this.root.style.setProperty("--org-link-opacity", `${linkOpacity}`);
        this.root.style.setProperty("--org-name-size", `${nameSize}px`);
        this.root.style.setProperty("--org-title-size", `${titleSize}px`);
        this.root.style.setProperty("--org-detail-size", `${detailSize}px`);
        this.root.style.setProperty("--org-border-width", `${borderWidth}px`);
        this.root.style.setProperty("--org-border-radius", `${borderRadius}px`);
        this.root.dataset.cardShadow = card.showShadow.value ? "1" : "0";
        this.root.dataset.showImages = card.showImage.value ? "1" : "0";
        this.root.dataset.cardAlign = alignment;
        this.root.style.setProperty("--org-show-images", card.showImage.value ? "1" : "0");

        this.toolbar.style.display = this.showToolbar ? "block" : "none";
        this.root.dataset.toolbarVisible = this.showToolbar ? "1" : "0";

        if (this.allowZoom) {
            this.svg.call(this.zoomBehavior);
        } else {
            this.svg.on(".zoom", null);
            this.zoomRoot.attr("transform", "translate(0,0) scale(1)");
            this.currentTransform = d3.zoomIdentity;
        }
    }

    private transform(dataView?: DataView): TransformResult {
        this.identityToNodeId.clear();
        this.nameToNodeId.clear();
        if (!dataView || !dataView.table) return { tree: null, nodesById: new Map(), warnings: ["No data"] };

        const table: DataViewTable = dataView.table;
        const columns = table.columns || [];

        const roleIndex = (role: string): number | undefined => {
            const idx = columns.findIndex((column) => column.roles && column.roles[role]);
            return idx >= 0 ? idx : undefined;
        };

        const roleIndexes = {
            employee: roleIndex("employeeId"),
            manager: roleIndex("managerId"),
            name: roleIndex("displayName"),
            title: roleIndex("title"),
            department: roleIndex("department"),
            image: roleIndex("imageUrl")
        };

        if (roleIndexes.employee === undefined) {
            return { tree: null, nodesById: new Map(), warnings: ["Employee Id field is required"] };
        }
        // NEW: Manager is also mandatory to start the chart
        if (roleIndexes.manager === undefined) {
            return { tree: null, nodesById: new Map(), warnings: ["Manager Id field is required"] };
        }

        const detailIndexes: number[] = [];
        const metricIndexes: number[] = [];

        columns.forEach((column, index) => {
            if (!column.roles) return;
            if (column.roles.details) detailIndexes.push(index);
            if (column.roles.metric) metricIndexes.push(index);
        });

        const nodes: OrgChartDatum[] = [];
        const nodesById: Map<string, ChartNode> = new Map();

        table.rows.forEach((row: DataViewTableRow, rowIndex: number) => {
            const employeeValue = row[roleIndexes.employee!];
            const employeeId = this.toKey(employeeValue);
            if (!employeeId) return;

            const managerValue = roleIndexes.manager !== undefined ? row[roleIndexes.manager] : undefined;
            const parentId = this.toKey(managerValue);

            const displayName = roleIndexes.name !== undefined ? this.toText(row[roleIndexes.name]) : undefined;
            const title = roleIndexes.title !== undefined ? this.toText(row[roleIndexes.title]) : undefined;
            const division = roleIndexes.department !== undefined ? this.toText(row[roleIndexes.department]) : undefined;
            const avatarUrl = roleIndexes.image !== undefined ? this.toText(row[roleIndexes.image]) : undefined;

            const details: FieldValue[] = [];
            detailIndexes.forEach((index) => {
                const value = this.toText(row[index]);
                if (value) details.push({ label: columns[index].displayName, value, role: "detail" });
            });

            const metrics: FieldValue[] = [];
            metricIndexes.forEach((index) => {
                const value = this.formatMetric(row[index], columns[index]);
                if (value) metrics.push({ label: columns[index].displayName, value, role: "metric" });
            });

            const identity = this.host.createSelectionIdBuilder()
                .withTable(table, rowIndex)
                .createSelectionId() as VisualSelectionId;

            const datum: OrgChartDatum = {
                id: employeeId,
                identity,
                parentId,
                displayName: displayName || employeeId,
                title,
                division,
                avatarUrl,
                details,
                metrics
            };

            nodes.push(datum);
            this.identityToNodeId.set(datum.identity.getKey(), datum.id);
            this.nameToNodeId.set(datum.displayName.toLowerCase(), datum.id);
        });

        if (!nodes.length) return { tree: null, nodesById: new Map(), warnings: ["No valid rows"] };

        const virtualRoot: ChartNode = { id: "__virtual__", payload: null, children: [], totalChildCount: 0 };
        const chartNodes: Map<string, ChartNode> = new Map();

        nodes.forEach((datum) => {
            const chartNode: ChartNode = { id: datum.id, payload: datum, children: [], totalChildCount: 0 };
            chartNodes.set(datum.id, chartNode);
        });

        chartNodes.forEach((node) => {
            const parentId = node.payload?.parentId;
            if (parentId && parentId !== node.id) {
                const parent = chartNodes.get(parentId);
                if (parent) { parent.children.push(node); return; }
            }
            virtualRoot.children.push(node);
        });

        const assignTotals = (node: ChartNode): void => {
            node.totalChildCount = node.children.length; // direct children
            node.children.forEach(assignTotals);
        };
        assignTotals(virtualRoot);

        nodesById.clear();
        chartNodes.forEach((node, id) => nodesById.set(id, node));

        return { tree: virtualRoot, nodesById, warnings: [] };
    }

    private pruneCollapsedNodes(): void {
        if (!this.fullTree) { this.collapsedNodes.clear(); return; }
        const validIds = new Set<string>();
        this.nodesById.forEach((_node, id) => validIds.add(id));
        Array.from(this.collapsedNodes).forEach((id) => { if (!validIds.has(id)) this.collapsedNodes.delete(id); });
    }

    // Collapse parents at depth == maxVisibleDepth
    private computeDefaultCollapsedSet(maxVisibleDepth: number): Set<string> {
        const set = new Set<string>();
        if (!this.fullTree) return set;
        const h = d3.hierarchy<ChartNode>(this.fullTree, (n) => n.children);
        h.each((d) => {
            // depth: 0 virtual root, 1 level-1, 2 level-2
            if (d.depth === maxVisibleDepth && d.data.payload) set.add(d.data.payload.id);
        });
        return set;
    }

    // ================= Render =================

    private render(): void {
        const layout = this.formattingSettings?.layout;
        const nodeWidth = this.clampNumber(layout?.nodeWidth.value, 160, 420, 260);
        const nodeHeight = this.clampNumber(layout?.nodeHeight.value, 80, 280, 140);
        const horizontalSpacing = this.clampNumber(layout?.horizontalSpacing.value, 20, 200, 60);
        const verticalSpacing = this.clampNumber(layout?.verticalSpacing.value, 20, 220, 80);

        this.svg.attr("width", this.viewport.width);
        this.svg.attr("height", this.viewport.height);

        const visibleTree = this.buildVisibleTree();
        const hasData = visibleTree && visibleTree.children.length > 0;
        this.emptyState.style.display = hasData ? "none" : "flex";
        if (!hasData) {
            this.linksGroup.selectAll("path").remove();
            this.nodesGroup.selectAll("g").remove();
            return;
        }

        const root = d3.hierarchy<ChartNode>(visibleTree!, (node) => node.children);
        const treeLayout = d3.tree<ChartNode>()
            .nodeSize([nodeWidth + horizontalSpacing, nodeHeight + verticalSpacing])
            .separation((a, b) => (a.parent === b.parent ? 1 : 1.3));
        const layoutRoot = treeLayout(root);

        const orientation = this.orientation;
        const nodeMap = new Map<string, RenderNode>();
        const renderNodes: RenderNode[] = [];

        layoutRoot.descendants().forEach((node) => {
            if (!node.data.payload) return;
            const payload = node.data.payload;
            const centerX = orientation === "vertical" ? node.x : node.y;
            const centerY = orientation === "vertical" ? node.y : node.x;
            const x = centerX - nodeWidth / 2;
            const y = centerY - nodeHeight / 2;
            const renderNode: RenderNode = {
                payload,
                hasChildren: node.data.totalChildCount > 0,
                isCollapsed: this.collapsedNodes.has(payload.id),
                x, y, centerX, centerY, width: nodeWidth, height: nodeHeight, node
            };
            nodeMap.set(payload.id, renderNode);
            renderNodes.push(renderNode);
        });

        // bounds
        let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
        renderNodes.forEach((rn) => {
            minX = Math.min(minX, rn.x);
            maxX = Math.max(maxX, rn.x + rn.width);
            minY = Math.min(minY, rn.y);
            maxY = Math.max(maxY, rn.y + rn.height);
        });
        const padding = 60;
        const baseX = padding - minX;
        const baseY = padding - minY;
        this.contentRoot.attr("transform", `translate(${baseX},${baseY})`);
        this.layoutBounds = { minX, maxX, minY, maxY, baseX, baseY, padding };

        this.nodePositions.clear();
        renderNodes.forEach((rn) => {
            this.nodePositions.set(rn.payload.id, {
                x: rn.x, y: rn.y, centerX: rn.centerX, centerY: rn.centerY,
                absX: rn.x + baseX, absY: rn.y + baseY,
                absCenterX: rn.centerX + baseX, absCenterY: rn.centerY + baseY,
                width: rn.width, height: rn.height
            });
        });

        // links
        const linkData: RenderLink[] = [];
        layoutRoot.links().forEach((link) => {
            const sp = link.source.data.payload, tp = link.target.data.payload;
            if (!sp || !tp) return; // skip virtual root
            const s = nodeMap.get(sp.id), t = nodeMap.get(tp.id);
            if (s && t) linkData.push({ source: s, target: t, key: `${s.payload.id}-${t.payload.id}` });
        });

        const linkSelection = this.linksGroup.selectAll<SVGPathElement, RenderLink>("path.orgchart__link")
            .data(linkData, (d: RenderLink) => d.key);

        linkSelection.enter()
            .append("path")
            .attr("class", "orgchart__link")
            .attr("d", (d) => this.linkPath(d))
            .attr("stroke", "var(--org-link-color)")
            .attr("stroke-width", () => `${this.formattingSettings.links.linkWidth.value}`)
            .attr("stroke-opacity", () => `${this.formattingSettings.links.linkOpacity.value}`)
            .attr("fill", "none");

        linkSelection
            .attr("d", (d) => this.linkPath(d))
            .attr("stroke-width", () => `${this.formattingSettings.links.linkWidth.value}`)
            .attr("stroke-opacity", () => `${this.formattingSettings.links.linkOpacity.value}`);

        linkSelection.exit().remove();

        // nodes
        const nodeSelection = this.nodesGroup.selectAll<SVGGElement, RenderNode>("g.orgchart__node")
            .data(renderNodes, (d: RenderNode) => d.payload.id);

        // ENTER — prevent paint until fully placed to avoid “top pop”
const nodeEnter = nodeSelection.enter()
  .append("g")
  .attr("class", "orgchart__node")
  .attr("transform", (d) => `translate(${d.x}, ${d.y})`)
  // hard kill any intermediate paint
  .style("display", "none")
  .on("click", (event, d) => this.handleNodeClick(event as MouseEvent, d));
        // seed new nodes at parent connector
        nodeEnter.attr("transform", (d) => {
            const parentId = d.node.parent?.data.payload?.id;
            if (parentId) {
                const p = this.nodePositions.get(parentId);
                if (p) {
                    const px = this.orientation === "vertical" ? p.centerX - d.width / 2 : p.x + p.width;
                    const py = this.orientation === "vertical" ? p.y + p.height : p.centerY - d.height / 2;
                    return `translate(${px}, ${py})`;
                }
            }
            return `translate(${d.x}, ${d.y})`;
        });

        // card bg
        nodeEnter.append("rect")
            .attr("class", "orgchart__node-bg")
            .attr("width", (d) => d.width)
            .attr("height", (d) => d.height)
            .attr("rx", this.clampNumber(this.formattingSettings.card.borderRadius.value, 0, 50, 18))
            .attr("ry", this.clampNumber(this.formattingSettings.card.borderRadius.value, 0, 50, 18));

        // card content (Name + Title + Division on separate lines)
        nodeEnter.each((d, index, groups) => {
            const group = groups[index];
            const foreignObject = d3.select(group)
                .append("foreignObject")
                .attr("width", d.width)
                .attr("height", d.height);

            const card = foreignObject
                .append("xhtml:div")
                .attr("class", "orgcard")
                .attr("data-card-align", this.formattingSettings.card.cardAlignment.value?.value || "start");

            const header = card.append("div").attr("class", "orgcard__header");
            header.append("div").attr("class", "orgcard__avatar").attr("data-id", d.payload.id);

            const titles = header.append("div").attr("class", "orgcard__titles");
            titles.append("div").attr("class", "orgcard__name");      // line 1
            titles.append("div").attr("class", "orgcard__title");     // line 2
            titles.append("div").attr("class", "orgcard__division");  // line 3

            card.append("div").attr("class", "orgcard__details");
            card.append("div").attr("class", "orgcard__metrics");
        });

        // toggle & child-count block near connector origin
        const toggleEnter = nodeEnter.append("g")
            .attr("class", "orgchart__node-toggle")
            .style("cursor", "pointer")
            .on("click", (event, d) => {
                event.preventDefault();
                event.stopPropagation();
                if (d.hasChildren) this.toggleNodeCollapseAnchoredOneLevel(d.payload.id);
            });

        toggleEnter.append("text")
            .attr("class", "orgchart__child-count")
            .attr("text-anchor", "middle")
            .attr("dy", "-12")
            .style("font-size", "11px")
            .style("fill", "var(--org-link-color)");

        toggleEnter.append("circle")
            .attr("r", 8)
            .attr("class", "orgchart__node-toggle-bg")
            .style("pointer-events", "all");

        toggleEnter.append("line").attr("class", "orgchart__toggle-line-h").style("pointer-events", "none");
        toggleEnter.append("line").attr("class", "orgchart__toggle-line-v").style("pointer-events", "none");

        // UPDATE (shared)
        this.updateNodeContent(nodeSelection.merge(nodeEnter));
        this.updateNodeToggles(nodeSelection.merge(nodeEnter));

// UPDATE + ENTER
const merged = nodeSelection.merge(nodeEnter)
  .attr("transform", (d) => `translate(${d.x}, ${d.y})`)
  .classed("has-children", (d) => d.hasChildren)
  .classed("is-collapsed", (d) => d.isCollapsed)
  .classed("is-highlighted", (d) => !!d.payload.highlight)
  .classed("is-selected", (d) => this.selectedKeys.has(d.payload.identity.getKey()));

// DOUBLE rAF REVEAL — guarantees transforms/children are committed before first paint
requestAnimationFrame(() => {
  requestAnimationFrame(() => {
    nodeEnter
      .style("display", "inline")    // reveal atomically, no transition
      .style("opacity", "1")         // (optional) if you had opacity styles elsewhere
      .style("visibility", "visible");
  });
});
        nodeSelection.exit().remove();

        if (this.selectedKeys.size > 0) {
            const selectedNodeId = Array.from(this.identityToNodeId.entries())
                .find(([key]) => this.selectedKeys.has(key))?.[1];
            if (selectedNodeId) this.highlightPath(selectedNodeId);
        }
    }

    private updateNodeContent(selection: d3.Selection<SVGGElement, RenderNode, any, unknown>): void {
        selection.each((d, index, groups) => {
            const group = d3.select(groups[index]);
            group.select("rect.orgchart__node-bg")
                .attr("width", d.width)
                .attr("height", d.height);

            const foreign = group.select<SVGForeignObjectElement>("foreignObject")
                .attr("width", d.width)
                .attr("height", d.height);

            const card = foreign.select<HTMLDivElement>("div.orgcard");
            const avatar = card.select<HTMLDivElement>(".orgcard__avatar");
            const nameEl = card.select<HTMLDivElement>(".orgcard__name");
            const titleEl = card.select<HTMLDivElement>(".orgcard__title");
            const divisionEl = card.select<HTMLDivElement>(".orgcard__division");
            const detailsEl = card.select<HTMLDivElement>(".orgcard__details");
            const metricsEl = card.select<HTMLDivElement>(".orgcard__metrics");

            const payload = d.payload;

            nameEl.text(payload.displayName);
            titleEl.text(payload.title || "");           // line 2
            divisionEl.text(payload.division || "");     // line 3

            const showImages = this.formattingSettings.card.showImage.value;
            if (showImages && payload.avatarUrl) {
                avatar.style("background-image", `url(${payload.avatarUrl})`);
                avatar.classed("orgcard__avatar--initials", false);
                avatar.text("");
            } else {
                avatar.style("background-image", "none");
                const initials = this.toInitials(payload.displayName);
                avatar.classed("orgcard__avatar--initials", true);
                avatar.text(initials);
            }

            // details – exclude title/division here (we already show them as separate lines)
            const detailNode = detailsEl.node();
            if (detailNode) {
                while (detailNode.firstChild) detailNode.removeChild(detailNode.firstChild);
                payload.details.forEach((item) => {
                    if (!item.value) return;
                    const line = detailNode.ownerDocument!.createElement("div");
                    line.className = "orgcard__detail-line";
                    if (item.label) {
                        const label = detailNode.ownerDocument!.createElement("span");
                        label.className = "orgcard__detail-label";
                        label.textContent = item.label;
                        line.appendChild(label);
                    }
                    const value = detailNode.ownerDocument!.createElement("span");
                    value.className = "orgcard__detail-value";
                    value.textContent = item.value;
                    line.appendChild(value);
                    detailNode.appendChild(line);
                });
            }

            const metricNode = metricsEl.node();
            if (metricNode) {
                while (metricNode.firstChild) metricNode.removeChild(metricNode.firstChild);
                payload.metrics.forEach((metric) => {
                    if (!metric.value) return;
                    const badge = metricNode.ownerDocument!.createElement("div");
                    badge.className = "orgcard__metric";
                    if (metric.label) {
                        const label = metricNode.ownerDocument!.createElement("span");
                        label.className = "orgcard__metric-label";
                        label.textContent = metric.label;
                        badge.appendChild(label);
                    }
                    const value = metricNode.ownerDocument!.createElement("span");
                    value.className = "orgcard__metric-value";
                    value.textContent = metric.value;
                    badge.appendChild(value);
                    metricNode.appendChild(badge);
                });
            }
        });
    }

    // Position + glyphs + direct child count for toggle
    private updateNodeToggles(selection: d3.Selection<SVGGElement, RenderNode, any, unknown>): void {
        selection.each((d, i, groups) => {
            const g = d3.select(groups[i]).select<SVGGElement>("g.orgchart__node-toggle");
            if (g.empty()) return;

            const chartNode = this.nodesById.get(d.payload.id);
            const directCount = chartNode ? chartNode.totalChildCount : 0;

            g.style("display", d.hasChildren ? "inline" : "none");

            // anchor position at connector origin
            let tx = 0, ty = 0;
            if (this.orientation === "vertical") { tx = d.width / 2; ty = d.height + 4; }
            else { tx = d.width + 4; ty = d.height / 2; }
            g.attr("transform", `translate(${tx}, ${ty})`);

            // child count label (above the circle)
            g.select<SVGTextElement>("text.orgchart__child-count")
                .text(d.hasChildren ? `(${directCount})` : "")
                .style("display", d.hasChildren ? "inline" : "none")
                .attr("x", 0).attr("y", 0).attr("dy", "-12");

            // toggle circle
            g.select<SVGCircleElement>("circle.orgchart__node-toggle-bg")
                .attr("r", 8).attr("fill", "#ffffff")
                .attr("stroke", "var(--org-link-color)")
                .attr("stroke-width", 1.2);

            // minus line
            g.select<SVGLineElement>("line.orgchart__toggle-line-h")
                .attr("x1", -4.5).attr("y1", 0)
                .attr("x2", 4.5).attr("y2", 0)
                .attr("stroke", "var(--org-link-color)")
                .attr("stroke-width", 1.4)
                .attr("stroke-linecap", "round");

            // plus vertical line (only when collapsed)
            const showPlus = d.isCollapsed;
            g.select<SVGLineElement>("line.orgchart__toggle-line-v")
                .attr("x1", 0).attr("y1", -4.5)
                .attr("x2", 0).attr("y2", 4.5)
                .attr("stroke", "var(--org-link-color)")
                .attr("stroke-width", 1.4)
                .attr("stroke-linecap", "round")
                .style("display", showPlus ? "inline" : "none");
        });
    }

    private linkPath(link: RenderLink): string {
        const s = link.source, t = link.target;
        if (this.orientation === "vertical") {
            const sx = s.centerX, sy = s.y + s.height;
            const tx = t.centerX, ty = t.y;
            const midY = sy + (ty - sy) / 2;
            return `M${sx},${sy} C${sx},${midY} ${tx},${midY} ${tx},${ty}`;
        }
        const sx = s.x + s.width, sy = s.centerY;
        const tx = t.x, ty = t.centerY;
        const midX = sx + (tx - sx) / 2;
        return `M${sx},${sy} C${midX},${sy} ${midX},${ty} ${tx},${ty}`;
    }

    // ================= Interactions / zoom / selection =================

    private handleNodeClick(event: MouseEvent, node: RenderNode): void {
        const multiSelect = event.ctrlKey || event.metaKey;
        this.selectionManager.select(node.payload.identity, multiSelect)
            .then((ids) => {
                this.selectedKeys = new Set(ids.map((id) => (id as VisualSelectionId).getKey()));
                this.updateSelectionVisuals();
                if (ids.length > 0) this.highlightPath(node.payload.id);
                else this.clearHighlights();
            })
            .catch(() => undefined);
        event.preventDefault();
        event.stopPropagation();
    }

    private updateSelectionVisuals(): void {
        this.nodesGroup.selectAll<SVGGElement, RenderNode>("g.orgchart__node")
            .classed("is-selected", (d) => this.selectedKeys.has(d.payload.identity.getKey()));
    }

    private toggleOrientation(): void {
        this.orientation = this.orientation === "vertical" ? "horizontal" : "vertical";
        this.render();
    }

    private expandAll(): void {
        this.collapsedNodes.clear();
        this.render();
    }

    private collapseAll(): void {
        if (!this.nodesById.size) return;
        this.collapsedNodes = new Set(Array.from(this.nodesById.keys()));
        this.render();
    }

    // Expand/collapse with anchor preservation; expand ONE level
    private toggleNodeCollapseAnchoredOneLevel(nodeId: string): void {
        const beforePos = this.nodePositions.get(nodeId);
        const wasCollapsed = this.collapsedNodes.has(nodeId);

        const k = this.currentTransform.k;
        const screenX = beforePos ? beforePos.absCenterX * k + this.currentTransform.x : 0;
        const screenY = beforePos ? beforePos.absCenterY * k + this.currentTransform.y : 0;

        if (wasCollapsed) this.expandOneLevel(nodeId);
        else this.collapsedNodes.add(nodeId); // collapse whole subtree implicitly

        this.render();

        const afterPos = this.nodePositions.get(nodeId);
        if (beforePos && afterPos) {
            const newX = screenX - afterPos.absCenterX * k;
            const newY = screenY - afterPos.absCenterY * k;
            const t = d3.zoomIdentity.translate(newX, newY).scale(k);
            this.svg.transition().duration(0).call(this.zoomBehavior.transform, t);
        }
    }

    // Only reveal direct children; keep grandchildren collapsed
    private expandOneLevel(nodeId: string): void {
        this.collapsedNodes.delete(nodeId);
        const node = this.nodesById.get(nodeId);
        if (node) node.children.forEach((c) => { if (c.payload) this.collapsedNodes.add(c.payload.id); });
    }

    private clampNumber(value: number | undefined, min: number, max: number, fallback: number): number {
        if (typeof value !== "number" || Number.isNaN(value)) return fallback;
        if (value < min) return min;
        if (value > max) return max;
        return value;
    }

    private focusOnNode(nodeId: string): void {
        if (!this.allowZoom) return;
        const position = this.nodePositions.get(nodeId);
        if (!position) return;
        const width = Math.max(this.viewport.width, 1);
        const height = Math.max(this.viewport.height, 1);
        const nodeWidth = position.width + 120;
        const nodeHeight = position.height + 120;
        const scale = Math.min(Math.min(width / nodeWidth, height / nodeHeight), 2.5);
        const translateX = width / 2 - position.absCenterX * scale;
        const translateY = height / 2 - position.absCenterY * scale;
        const transform = d3.zoomIdentity.translate(translateX, translateY).scale(scale);
        this.svg.transition().duration(0).call(this.zoomBehavior.transform, transform);
    }

    private fitToViewport(duration: number = 400): void {
        if (!this.allowZoom || !this.layoutBounds) return;
        const b = this.layoutBounds;
        const width = Math.max(this.viewport.width, 1);
        const height = Math.max(this.viewport.height, 1);
        const contentWidth = b.maxX - b.minX + b.padding * 2;
        const contentHeight = b.maxY - b.minY + b.padding * 2;
        const scale = Math.min(width / contentWidth, height / contentHeight);
        const translateX = (width - contentWidth * scale) / 2;
        const translateY = (height - contentHeight * scale) / 2;
        const transform = d3.zoomIdentity.translate(translateX, translateY).scale(scale);
        this.svg.transition().duration(duration).call(this.zoomBehavior.transform, transform);
    }

    private resetZoom(): void {
        if (!this.allowZoom) return;
        this.svg.transition().duration(0).call(this.zoomBehavior.transform, d3.zoomIdentity);
    }

    private onZoom(event: d3.D3ZoomEvent<SVGSVGElement, unknown>): void {
        if (!this.allowZoom) return;
        this.currentTransform = event.transform;
        this.zoomRoot.attr("transform", event.transform.toString());
    }

    private updateHighlights(dataView?: DataView): void {
        if (!dataView || !dataView.table) return;

        this.nodesById.forEach((node) => { if (node.payload) node.payload.highlight = false; });

        const selections = this.selectionManager.getSelectionIds();
        if (selections && selections.length > 0) {
            selections.forEach((selection) => {
                const nodeId = this.identityToNodeId.get((selection as VisualSelectionId).getKey());
                if (nodeId) {
                    const node = this.nodesById.get(nodeId);
                    if (node && node.payload) {
                        node.payload.highlight = true;
                        this.ensureNodeVisible(nodeId);
                        this.highlightPath(nodeId);
                    }
                }
            });
        } else {
            this.clearHighlights();
        }
    }

    private ensureNodeVisible(nodeId: string): void {
        const visited = new Set<string>();
        let current = this.nodesById.get(nodeId);
        while (current && current.payload) {
            if (visited.has(current.id)) break;
            visited.add(current.id);
            this.collapsedNodes.delete(current.id);
            const parentId = current.payload.parentId;
            current = parentId ? this.nodesById.get(parentId) : undefined;
        }
    }

    private buildVisibleTree(): ChartNode | null {
        if (!this.fullTree) return null;

        const cloneNode = (node: ChartNode): ChartNode => {
            const payloadId = node.payload?.id;
            const isCollapsed = payloadId ? this.collapsedNodes.has(payloadId) : false;
            const children = isCollapsed ? [] : node.children.map((child) => cloneNode(child));
            return { id: node.id, payload: node.payload, children, totalChildCount: node.totalChildCount };
        };

        return cloneNode(this.fullTree);
    }

    private getPathToRoot(nodeId: string): string[] {
        const path: string[] = [];
        let current = this.nodesById.get(nodeId);
        while (current && current.payload) {
            path.push(current.id);
            const parentId = current.payload.parentId;
            current = parentId ? this.nodesById.get(parentId) : undefined;
        }
        return path;
    }

    private highlightPath(nodeId: string): void {
        const pathIds = this.getPathToRoot(nodeId);

        this.nodesGroup.selectAll<SVGGElement, RenderNode>("g.orgchart__node")
            .classed("is-dimmed", (d) => !pathIds.includes(d.payload.id))
            .classed("is-highlighted", (d) => pathIds.includes(d.payload.id));

        this.linksGroup.selectAll<SVGPathElement, RenderLink>("path.orgchart__link")
            .classed("is-dimmed", (d) => {
                const s = pathIds.includes(d.source.payload.id);
                const t = pathIds.includes(d.target.payload.id);
                return !(s && t);
            })
            .classed("is-highlighted", (d) => {
                const s = pathIds.includes(d.source.payload.id);
                const t = pathIds.includes(d.target.payload.id);
                return s && t;
            });
    }

    private clearHighlights(): void {
        this.nodesGroup.selectAll<SVGGElement, RenderNode>("g.orgchart__node")
            .classed("is-dimmed", false)
            .classed("is-highlighted", false);

        this.linksGroup.selectAll<SVGPathElement, RenderLink>("path.orgchart__link")
            .classed("is-dimmed", false)
            .classed("is-highlighted", false);
    }

    private toKey(value: PrimitiveValue): string | undefined {
        if (value == null || value === "") return undefined;
        return String(value).trim();
    }

    private toText(value: PrimitiveValue): string | undefined {
        if (value == null || value === "") return undefined;
        return String(value).trim();
    }

    private formatMetric(value: PrimitiveValue, _column: DataViewMetadataColumn): string | undefined {
        if (value == null || value === "") return undefined;
        if (typeof value === "number") return value.toLocaleString();
        return String(value);
    }

    private toInitials(name: string): string {
        if (!name) return "";
        const words = name.trim().split(/\s+/);
        if (words.length === 1) return words[0].charAt(0).toUpperCase();
        return (words[0].charAt(0) + words[words.length - 1].charAt(0)).toUpperCase();
    }
}
