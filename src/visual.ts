/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the "Software"), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
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
    role: "detail" | "metric" | "department" | "title";
}

interface OrgChartDatum {
    id: string;
    identity: VisualSelectionId;
    parentId?: string;
    displayName: string;
    title?: string;
    department?: string;
    avatarUrl?: string;
    details: FieldValue[];
    metrics: FieldValue[];
    highlight?: boolean;
}

interface ChartNode {
    id: string;
    payload: OrgChartDatum | null;
    children: ChartNode[];
    totalChildCount: number;
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
    private toolbar: HTMLDivElement;
    private toolbarRail: HTMLDivElement;
    private toolbarToggle: HTMLButtonElement;
    private toolbarToggleLabel: HTMLSpanElement;
    private toolbarMenu: HTMLDivElement;
    private toolbarMenuHeader: HTMLButtonElement;
    private toolbarMenuItems: HTMLDivElement;
    private toolbarMenuOpen: boolean = true;
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
    private filterNodeIds: Set<string> = new Set();
    private identityToNodeId: Map<string, string> = new Map();
    private nameToNodeId: Map<string, string> = new Map();

    private emptyState: HTMLDivElement;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = this.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();

        this.root = document.createElement("div");
        this.root.className = "orgchart-visual";
        this.root.dataset.toolbarVisible = "1";

        this.toolbar = document.createElement("div");
        this.toolbar.className = "orgchart__toolbar";
        this.root.appendChild(this.toolbar);

        this.toolbarRail = document.createElement("div");
        this.toolbarRail.className = "orgchart__toolbar-rail";
        this.toolbar.appendChild(this.toolbarRail);

        this.toolbarMenu = document.createElement("div");
        this.toolbarMenu.className = "orgchart__menu";
        this.root.appendChild(this.toolbarMenu);

        const menuCard = document.createElement("div");
        menuCard.className = "orgchart__menu-card";
        this.toolbarMenu.appendChild(menuCard);

        this.toolbarMenuHeader = document.createElement("button");
        this.toolbarMenuHeader.type = "button";
        this.toolbarMenuHeader.className = "orgchart__menu-header";
        this.toolbarMenuHeader.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            this.toolbarMenuOpen = !this.toolbarMenuOpen;
            this.updateToolbarMenuState();
        });
        const headerIcon = this.createToolbarIcon("menu");
        headerIcon.classList.add("orgchart__toolbar-icon");
        this.toolbarMenuHeader.appendChild(headerIcon);
        this.toolbarToggleLabel = document.createElement("span");
        this.toolbarToggleLabel.className = "orgchart__menu-header-label";
        this.toolbarMenuHeader.appendChild(this.toolbarToggleLabel);
        menuCard.appendChild(this.toolbarMenuHeader);

        this.toolbarMenuItems = document.createElement("div");
        this.toolbarMenuItems.className = "orgchart__menu-items";
        menuCard.appendChild(this.toolbarMenuItems);

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
                // Allow mouse wheel zoom and pan, prevent conflicts
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
        this.updateToolbarMenuState();

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
        this.pruneCollapsedNodes();
        this.updateHighlights(dataView);

        this.render();
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private buildToolbar(): void {
        if (!this.toolbarRail || !this.toolbarMenuItems) {
            return;
        }

        while (this.toolbarRail.firstChild) {
            this.toolbarRail.removeChild(this.toolbarRail.firstChild);
        }
        while (this.toolbarMenuItems.firstChild) {
            this.toolbarMenuItems.removeChild(this.toolbarMenuItems.firstChild);
        }

        this.toolbarToggle = this.createToolbarRailButton("menu", "Hide controls", () => {
            this.toolbarMenuOpen = !this.toolbarMenuOpen;
            this.updateToolbarMenuState();
        });
        this.toolbarToggle.classList.add("orgchart__toolbar-btn--toggle");
        this.toolbarRail.appendChild(this.toolbarToggle);

        const buttons: Array<{ label: string; title: string; icon: ToolbarIcon; onClick: () => void; }> = [
            {
                label: "Switch layout",
                title: "Toggle between vertical and horizontal layouts",
                icon: "layout",
                onClick: () => this.toggleOrientation()
            },
            {
                label: "Expand all",
                title: "Expand all nodes",
                icon: "expand",
                onClick: () => this.expandAll()
            },
            {
                label: "Collapse all",
                title: "Collapse to root nodes",
                icon: "collapse",
                onClick: () => this.collapseAll()
            },
            {
                label: "Fit",
                title: "Fit chart to view",
                icon: "fit",
                onClick: () => this.fitToViewport()
            },
            {
                label: "Reset",
                title: "Reset zoom and pan",
                icon: "reset",
                onClick: () => this.resetZoom()
            }
        ];

        buttons.forEach((config) => {
            const railButton = this.createToolbarRailButton(config.icon, config.title, config.onClick);
            this.toolbarRail.appendChild(railButton);

            const menuButton = this.createToolbarMenuButton(config.label, config.title, config.icon, config.onClick);
            this.toolbarMenuItems.appendChild(menuButton);
        });
    }

    private createToolbarRailButton(icon: ToolbarIcon, title: string, onClick: () => void): HTMLButtonElement {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "orgchart__toolbar-btn";
        button.title = title;
        button.setAttribute("aria-label", title);

        const iconElement = this.createToolbarIcon(icon);
        iconElement.classList.add("orgchart__toolbar-icon");
        button.appendChild(iconElement);

        button.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            onClick();
        });
        return button;
    }

    private createToolbarMenuButton(label: string, title: string, icon: ToolbarIcon, onClick: () => void): HTMLButtonElement {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "orgchart__menu-item";
        button.title = title;
        button.setAttribute("aria-label", title);

        const iconElement = this.createToolbarIcon(icon);
        iconElement.classList.add("orgchart__toolbar-icon");
        button.appendChild(iconElement);

        const textElement = document.createElement("span");
        textElement.className = "orgchart__menu-item-label";
        textElement.textContent = label;
        button.appendChild(textElement);

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
            element.setAttribute("stroke-width", "1.8");
            element.setAttribute("stroke-linecap", "round");
            element.setAttribute("stroke-linejoin", "round");
            return element;
        };

        const append = <T extends SVGElement>(element: T): T => {
            svg.appendChild(element);
            return element;
        };

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
            case "menu": {
                [7, 12, 17].forEach((y) => createLine(5, y, 19, y));
                break;
            }
            case "layout": {
                createRect(4.5, 4, 7.5, 16);
                createRect(13.5, 8, 7.5, 12);
                break;
            }
            case "expand": {
                createCircle(12, 12, 8);
                createLine(12, 8.5, 12, 15.5);
                createLine(8.5, 12, 15.5, 12);
                break;
            }
            case "collapse": {
                createCircle(12, 12, 8);
                createLine(8.5, 12, 15.5, 12);
                break;
            }
            case "fit": {
                createCircle(11, 11, 5.5);
                createLine(15, 15, 20, 20);
                break;
            }
            case "reset": {
                createPath("M20 12a8 8 0 1 1-2.34-5.66");
                createPolyline("20,6 20,12 14,12");
                break;
            }
        }

        return svg;
    }

    private updateToolbarMenuState(): void {
        if (!this.toolbar || !this.toolbarToggle || !this.toolbarMenu) {
            return;
        }

        const isVisible = this.showToolbar;
        const isOpen = isVisible && this.toolbarMenuOpen;

        this.toolbarMenu.classList.toggle("is-open", isOpen);
        this.toolbarMenu.style.display = isVisible ? "block" : "none";

        const label = isOpen ? "Hide controls" : "Show controls";
        this.toolbarToggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
        this.toolbarToggle.setAttribute("aria-label", label);
        this.toolbarToggle.title = label;

        if (this.toolbarMenuHeader) {
            this.toolbarMenuHeader.setAttribute("aria-expanded", isOpen ? "true" : "false");
            this.toolbarMenuHeader.setAttribute("aria-label", label);
            this.toolbarMenuHeader.title = label;
        }

        if (this.toolbarToggleLabel) {
            this.toolbarToggleLabel.textContent = label;
        }

        if (this.root) {
            this.root.dataset.toolbarMenu = isOpen ? "1" : "0";
        }
    }

    private applyFormatting(): void {
        if (!this.formattingSettings) {
            return;
        }

        const layout = this.formattingSettings.layout;
        const card = this.formattingSettings.card;
        const labels = this.formattingSettings.labels;
        const links = this.formattingSettings.links;

        this.allowZoom = layout.enableZoom.value;
        this.showToolbar = layout.showToolbar.value;

        const nameSize = this.clampNumber(labels.nameTextSize.value, 8, 32, 16);
        const titleSize = this.clampNumber(labels.titleTextSize.value, 8, 28, 12);
        const detailSize = this.clampNumber(labels.detailTextSize.value, 6, 24, 11);
        const linkWidth = this.clampNumber(links.linkWidth.value, 0.5, 6, 1.8);
        const linkOpacity = this.clampNumber(links.linkOpacity.value, 0.1, 1, 0.8);
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
        this.updateToolbarMenuState();

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
        if (!dataView || !dataView.table) {
            return { tree: null, nodesById: new Map(), warnings: ["No data"] };
        }

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

        const detailIndexes: number[] = [];
        const metricIndexes: number[] = [];

        columns.forEach((column, index) => {
            if (!column.roles) {
                return;
            }
            if (column.roles.details) {
                detailIndexes.push(index);
            }
            if (column.roles.metric) {
                metricIndexes.push(index);
            }
        });

        const nodes: OrgChartDatum[] = [];
        const nodesById: Map<string, ChartNode> = new Map();

        table.rows.forEach((row: DataViewTableRow, rowIndex: number) => {
            const employeeValue = row[roleIndexes.employee!];
            const employeeId = this.toKey(employeeValue);
            if (!employeeId) {
                return;
            }

            const managerValue = roleIndexes.manager !== undefined ? row[roleIndexes.manager] : undefined;
            const parentId = this.toKey(managerValue);

            const displayName = roleIndexes.name !== undefined ? this.toText(row[roleIndexes.name]) : undefined;
            const title = roleIndexes.title !== undefined ? this.toText(row[roleIndexes.title]) : undefined;
            const department = roleIndexes.department !== undefined ? this.toText(row[roleIndexes.department]) : undefined;
            const avatarUrl = roleIndexes.image !== undefined ? this.toText(row[roleIndexes.image]) : undefined;

            const details: FieldValue[] = [];
            detailIndexes.forEach((index) => {
                const value = this.toText(row[index]);
                if (value) {
                    details.push({
                        label: columns[index].displayName,
                        value,
                        role: "detail"
                    });
                }
            });

            const metrics: FieldValue[] = [];
            metricIndexes.forEach((index) => {
                const value = this.formatMetric(row[index], columns[index]);
                if (value) {
                    metrics.push({
                        label: columns[index].displayName,
                        value,
                        role: "metric"
                    });
                }
            });

            if (department) {
                details.unshift({ label: "", value: department, role: "department" });
            }
            if (title) {
                details.unshift({ label: "", value: title, role: "title" });
            }

            const identity = this.host.createSelectionIdBuilder()
                .withTable(table, rowIndex)
                .createSelectionId() as VisualSelectionId;

            const datum: OrgChartDatum = {
                id: employeeId,
                identity,
                parentId,
                displayName: displayName || employeeId,
                title,
                department,
                avatarUrl,
                details,
                metrics
            };

            nodes.push(datum);
            this.identityToNodeId.set(datum.identity.getKey(), datum.id);
            this.nameToNodeId.set(datum.displayName.toLowerCase(), datum.id);
        });

        if (!nodes.length) {
            return { tree: null, nodesById: new Map(), warnings: ["No valid rows"] };
        }

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
                if (parent) {
                    parent.children.push(node);
                    return;
                }
            }
            virtualRoot.children.push(node);
        });

        const assignTotals = (node: ChartNode): void => {
            node.totalChildCount = node.children.length;
            node.children.forEach(assignTotals);
        };
        assignTotals(virtualRoot);

        nodesById.clear();
        chartNodes.forEach((node, id) => nodesById.set(id, node));

        return { tree: virtualRoot, nodesById, warnings: [] };
    }

    private pruneCollapsedNodes(): void {
        if (!this.fullTree) {
            this.collapsedNodes.clear();
            return;
        }
        const validIds = new Set<string>();
        this.nodesById.forEach((_node, id) => validIds.add(id));
        Array.from(this.collapsedNodes).forEach((id) => {
            if (!validIds.has(id)) {
                this.collapsedNodes.delete(id);
            }
        });
    }

    private buildVisibleTree(): ChartNode | null {
        if (!this.fullTree) {
            return null;
        }

        const cloneNode = (node: ChartNode): ChartNode => {
            const payloadId = node.payload?.id;
            const isCollapsed = payloadId ? this.collapsedNodes.has(payloadId) : false;
            const children = isCollapsed ? [] : node.children.map((child) => cloneNode(child));
            return {
                id: node.id,
                payload: node.payload,
                children,
                totalChildCount: node.totalChildCount
            };
        };

        return cloneNode(this.fullTree);
    }

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
        const treeLayout = d3.tree<ChartNode>().nodeSize([nodeWidth + horizontalSpacing, nodeHeight + verticalSpacing])
            .separation((a, b) => (a.parent === b.parent ? 1 : 1.3));
        const layoutRoot = treeLayout(root);

        const orientation = this.orientation;
        const nodeMap = new Map<string, RenderNode>();
        const renderNodes: RenderNode[] = [];

        layoutRoot.descendants().forEach((node) => {
            if (!node.data.payload) {
                return;
            }
            const payload = node.data.payload;
            const centerX = orientation === "vertical" ? node.x : node.y;
            const centerY = orientation === "vertical" ? node.y : node.x;
            const x = centerX - nodeWidth / 2;
            const y = centerY - nodeHeight / 2;
            const renderNode: RenderNode = {
                payload,
                hasChildren: node.data.totalChildCount > 0,
                isCollapsed: this.collapsedNodes.has(payload.id),
                x,
                y,
                centerX,
                centerY,
                width: nodeWidth,
                height: nodeHeight,
                node
            };
            nodeMap.set(payload.id, renderNode);
            renderNodes.push(renderNode);
        });

        let minX = Infinity;
        let maxX = -Infinity;
        let minY = Infinity;
        let maxY = -Infinity;
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
                x: rn.x,
                y: rn.y,
                centerX: rn.centerX,
                centerY: rn.centerY,
                absX: rn.x + baseX,
                absY: rn.y + baseY,
                absCenterX: rn.centerX + baseX,
                absCenterY: rn.centerY + baseY,
                width: rn.width,
                height: rn.height
            });
        });

        const linkData: RenderLink[] = [];
        layoutRoot.links().forEach((link) => {
            const sourcePayload = link.source.data.payload;
            const targetPayload = link.target.data.payload;
            if (!sourcePayload || !targetPayload) {
                return;
            }
            const sourceNode = nodeMap.get(sourcePayload.id);
            const targetNode = nodeMap.get(targetPayload.id);
            if (sourceNode && targetNode) {
                linkData.push({ 
                    source: sourceNode, 
                    target: targetNode, 
                    key: `${sourceNode.payload.id}-${targetNode.payload.id}` 
                });
            }
        });

        const linkSelection = this.linksGroup.selectAll<SVGPathElement, RenderLink>("path.orgchart__link")
            .data(linkData, (d: RenderLink) => `${d.source.payload.id}-${d.target.payload.id}`);

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

        const nodeSelection = this.nodesGroup.selectAll<SVGGElement, RenderNode>("g.orgchart__node")
            .data(renderNodes, (d: RenderNode) => d.payload.id);

        const nodeEnter = nodeSelection.enter()
            .append("g")
            .attr("class", "orgchart__node")
            .attr("transform", (d) => `translate(${d.x}, ${d.y})`)
            .on("click", (event, d) => this.handleNodeClick(event as MouseEvent, d));

        nodeEnter.append("rect")
            .attr("class", "orgchart__node-bg")
            .attr("width", (d) => d.width)
            .attr("height", (d) => d.height)
            .attr("rx", this.clampNumber(this.formattingSettings.card.borderRadius.value, 0, 50, 18))
            .attr("ry", this.clampNumber(this.formattingSettings.card.borderRadius.value, 0, 50, 18));

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
            const avatar = header.append("div").attr("class", "orgcard__avatar");
            avatar.attr("data-id", d.payload.id);

            const titles = header.append("div").attr("class", "orgcard__titles");
            titles.append("div").attr("class", "orgcard__name");
            titles.append("div").attr("class", "orgcard__title");

            const toggle = header.append("button")
                .attr("type", "button")
                .attr("class", "orgcard__toggle")
                .text(d.hasChildren ? (d.isCollapsed ? "+" : "−") : "")
                .on("click", (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    if (d.hasChildren) {
                        this.toggleNodeCollapse(d.payload.id);
                    }
                });

            card.append("div").attr("class", "orgcard__details");
            card.append("div").attr("class", "orgcard__metrics");
        });

        this.updateNodeContent(nodeSelection.merge(nodeEnter));

        nodeSelection.merge(nodeEnter)
            .attr("transform", (d) => `translate(${d.x}, ${d.y})`)
            .classed("has-children", (d) => d.hasChildren)
            .classed("is-collapsed", (d) => d.isCollapsed)
            .classed("is-highlighted", (d) => !!d.payload.highlight)
            .classed("is-selected", (d) => this.selectedKeys.has(d.payload.identity.getKey()));

        nodeSelection.exit().remove();

        // Apply highlighting effects after rendering if there are selections
        if (this.selectedKeys.size > 0) {
            const selectedNodeId = Array.from(this.identityToNodeId.entries())
                .find(([key]) => this.selectedKeys.has(key))?.[1];
            if (selectedNodeId) {
                this.highlightPath(selectedNodeId);
            }
        }
    }

    private updateNodeContent(selection: d3.Selection<SVGGElement, RenderNode, any, unknown>): void {
        selection.each((d, index, groups) => {
            const group = d3.select(groups[index]);
            group.select("rect.orgchart__node-bg")
                .attr("width", d.width)
                .attr("height", d.height);

            const foreign = group.select<SVGForeignObjectElement>("foreignObject");
            foreign
                .attr("width", d.width)
                .attr("height", d.height);

            const card = foreign.select<HTMLDivElement>("div.orgcard");
            const avatar = card.select<HTMLDivElement>(".orgcard__avatar");
            const nameEl = card.select<HTMLDivElement>(".orgcard__name");
            const titleEl = card.select<HTMLDivElement>(".orgcard__title");
            const detailsEl = card.select<HTMLDivElement>(".orgcard__details");
            const metricsEl = card.select<HTMLDivElement>(".orgcard__metrics");
            const toggleButton = card.select<HTMLButtonElement>(".orgcard__toggle");

            const payload = d.payload;

            nameEl.text(payload.displayName);
            titleEl.text(payload.title || payload.department || "");

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

            const detailNode = detailsEl.node();
            if (detailNode) {
                while (detailNode.firstChild) {
                    detailNode.removeChild(detailNode.firstChild);
                }
                payload.details.forEach((item) => {
                    if (!item.value || item.role === "title" || item.role === "department") {
                        return;
                    }
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
                while (metricNode.firstChild) {
                    metricNode.removeChild(metricNode.firstChild);
                }
                payload.metrics.forEach((metric) => {
                    if (!metric.value) {
                        return;
                    }
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

            toggleButton.text(d.hasChildren ? (d.isCollapsed ? "+" : "−") : "");
        });
    }

    private linkPath(link: RenderLink): string {
        const source = link.source;
        const target = link.target;
        if (this.orientation === "vertical") {
            const sx = source.centerX;
            const sy = source.y + source.height;
            const tx = target.centerX;
            const ty = target.y;
            const midY = sy + (ty - sy) / 2;
            return `M${sx},${sy} C${sx},${midY} ${tx},${midY} ${tx},${ty}`;
        }

        const sx = source.x + source.width;
        const sy = source.centerY;
        const tx = target.x;
        const ty = target.centerY;
        const midX = sx + (tx - sx) / 2;
        return `M${sx},${sy} C${midX},${sy} ${midX},${ty} ${tx},${ty}`;
    }

    private handleNodeClick(event: MouseEvent, node: RenderNode): void {
        const multiSelect = event.ctrlKey || event.metaKey;
        this.selectionManager.select(node.payload.identity, multiSelect)
            .then((ids) => {
                this.selectedKeys = new Set(ids.map((id) => (id as VisualSelectionId).getKey()));
                this.updateSelectionVisuals();
                
                // Apply highlighting to the selected node path
                if (ids.length > 0) {
                    this.highlightPath(node.payload.id);
                } else {
                    this.clearHighlights();
                }
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
        if (!this.nodesById.size) {
            return;
        }
        this.collapsedNodes = new Set(Array.from(this.nodesById.keys()));
        this.render();
    }

    private toggleNodeCollapse(nodeId: string): void {
        if (this.collapsedNodes.has(nodeId)) {
            this.collapsedNodes.delete(nodeId);
        } else {
            this.collapsedNodes.add(nodeId);
        }
        this.render();
    }



    private ensureNodeVisible(nodeId: string): void {
        const visited = new Set<string>();
        let current = this.nodesById.get(nodeId);
        while (current && current.payload) {
            if (visited.has(current.id)) {
                break;
            }
            visited.add(current.id);
            this.collapsedNodes.delete(current.id);
            const parentId = current.payload.parentId;
            current = parentId ? this.nodesById.get(parentId) : undefined;
        }
    }

    private clampNumber(value: number | undefined, min: number, max: number, fallback: number): number {
        if (typeof value !== "number" || Number.isNaN(value)) {
            return fallback;
        }
        if (value < min) {
            return min;
        }
        if (value > max) {
            return max;
        }
        return value;
    }

    private focusOnNode(nodeId: string): void {
        if (!this.allowZoom) {
            return;
        }
        const position = this.nodePositions.get(nodeId);
        if (!position) {
            return;
        }
        const width = Math.max(this.viewport.width, 1);
        const height = Math.max(this.viewport.height, 1);
        const nodeWidth = position.width + 120;
        const nodeHeight = position.height + 120;
        const scale = Math.min(Math.min(width / nodeWidth, height / nodeHeight), 2.5);
        const translateX = width / 2 - position.absCenterX * scale;
        const translateY = height / 2 - position.absCenterY * scale;
        const transform = d3.zoomIdentity.translate(translateX, translateY).scale(scale);
        this.svg.transition().duration(480).call(this.zoomBehavior.transform, transform);
    }

    private fitToViewport(): void {
        if (!this.allowZoom || !this.layoutBounds) {
            return;
        }
        const bounds = this.layoutBounds;
        const width = Math.max(this.viewport.width, 1);
        const height = Math.max(this.viewport.height, 1);
        const contentWidth = bounds.maxX - bounds.minX + bounds.padding * 2;
        const contentHeight = bounds.maxY - bounds.minY + bounds.padding * 2;
        const scale = Math.min(width / contentWidth, height / contentHeight);
        const translateX = (width - contentWidth * scale) / 2;
        const translateY = (height - contentHeight * scale) / 2;
        const transform = d3.zoomIdentity.translate(translateX, translateY).scale(scale);
        this.svg.transition().duration(400).call(this.zoomBehavior.transform, transform);
    }

    private resetZoom(): void {
        if (!this.allowZoom) {
            return;
        }
        this.svg.transition().duration(300).call(this.zoomBehavior.transform, d3.zoomIdentity);
    }

    private onZoom(event: d3.D3ZoomEvent<SVGSVGElement, unknown>): void {
        if (!this.allowZoom) {
            return;
        }
        this.currentTransform = event.transform;
        this.zoomRoot.attr("transform", event.transform.toString());
    }

    private updateHighlights(dataView?: DataView): void {
        if (!dataView || !dataView.table) {
            return;
        }

        // Clear existing highlights
        this.nodesById.forEach((node) => {
            if (node.payload) {
                node.payload.highlight = false;
            }
        });

        // Check for cross-highlighting from other visuals
        const selections = this.selectionManager.getSelectionIds();
        if (selections && selections.length > 0) {
            selections.forEach((selection) => {
                const nodeId = this.identityToNodeId.get((selection as VisualSelectionId).getKey());
                if (nodeId) {
                    const node = this.nodesById.get(nodeId);
                    if (node && node.payload) {
                        node.payload.highlight = true;
                        this.ensureNodeVisible(nodeId);
                        // Highlight path to root
                        this.highlightPath(nodeId);
                    }
                }
            });
        } else {
            // Clear highlights when no selections
            this.clearHighlights();
        }
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
        
        // Update visual state for nodes
        this.nodesGroup.selectAll<SVGGElement, RenderNode>("g.orgchart__node")
            .classed("is-dimmed", (d) => !pathIds.includes(d.payload.id))
            .classed("is-highlighted", (d) => pathIds.includes(d.payload.id));

        // Update visual state for links
        this.linksGroup.selectAll<SVGPathElement, RenderLink>("path.orgchart__link")
            .classed("is-dimmed", (d) => {
                const sourceInPath = pathIds.includes(d.source.payload.id);
                const targetInPath = pathIds.includes(d.target.payload.id);
                return !(sourceInPath && targetInPath);
            })
            .classed("is-highlighted", (d) => {
                const sourceInPath = pathIds.includes(d.source.payload.id);
                const targetInPath = pathIds.includes(d.target.payload.id);
                return sourceInPath && targetInPath;
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
        if (value == null || value === "") {
            return undefined;
        }
        return String(value).trim();
    }

    private toText(value: PrimitiveValue): string | undefined {
        if (value == null || value === "") {
            return undefined;
        }
        return String(value).trim();
    }

    private formatMetric(value: PrimitiveValue, column: DataViewMetadataColumn): string | undefined {
        if (value == null || value === "") {
            return undefined;
        }
        if (typeof value === "number") {
            return value.toLocaleString();
        }
        return String(value);
    }

    private toInitials(name: string): string {
        if (!name) {
            return "";
        }
        const words = name.trim().split(/\s+/);
        if (words.length === 1) {
            return words[0].charAt(0).toUpperCase();
        }
        return (words[0].charAt(0) + words[words.length - 1].charAt(0)).toUpperCase();
    }
}
