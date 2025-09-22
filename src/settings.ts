/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
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

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsDropdown = formattingSettings.ItemDropdown;

class LayoutCardSettings extends FormattingSettingsCard {
    nodeWidth = new formattingSettings.NumUpDown({
        name: "nodeWidth",
        displayName: "Node width",
        value: 260
    });

    nodeHeight = new formattingSettings.NumUpDown({
        name: "nodeHeight",
        displayName: "Node height",
        value: 140
    });

    horizontalSpacing = new formattingSettings.NumUpDown({
        name: "horizontalSpacing",
        displayName: "Horizontal spacing",
        value: 60
    });

    verticalSpacing = new formattingSettings.NumUpDown({
        name: "verticalSpacing",
        displayName: "Vertical spacing",
        value: 80
    });

    initialExpandedLevels = new FormattingSettingsDropdown({
        name: "initialExpandedLevels",
        displayName: "Initial expanded levels",
        value: { value: "2", displayName: "2" },
        items: [
            { value: "0", displayName: "None" },
            { value: "1", displayName: "1" },
            { value: "2", displayName: "2" },
            { value: "3", displayName: "3" },
            { value: "4", displayName: "4" },
            { value: "5", displayName: "5" },
            { value: "6", displayName: "6" }
        ]
    });

    enableZoom = new formattingSettings.ToggleSwitch({
        name: "enableZoom",
        displayName: "Enable zoom & pan",
        value: true
    });

    showToolbar = new formattingSettings.ToggleSwitch({
        name: "showToolbar",
        displayName: "Show toolbar",
        value: true
    });

    name: string = "layout";
    displayName: string = "Layout";
    slices: FormattingSettingsSlice[] = [
        this.nodeWidth,
        this.nodeHeight,
        this.horizontalSpacing,
        this.verticalSpacing,
        this.initialExpandedLevels,
        this.enableZoom,
        this.showToolbar
    ];
}

class CardStyleSettings extends FormattingSettingsCard {
    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background",
        value: { value: "#ffffff" }
    });

    accentColor = new formattingSettings.ColorPicker({
        name: "accentColor",
        displayName: "Accent",
        value: { value: "#2563eb" }
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Border",
        value: { value: "#cbd5f5" }
    });

    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Primary text",
        value: { value: "#0f172a" }
    });

    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth",
        displayName: "Border width",
        value: 1
    });

    borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Corner radius",
        value: 18
    });

    showShadow = new formattingSettings.ToggleSwitch({
        name: "showShadow",
        displayName: "Drop shadow",
        value: true
    });

    showImage = new formattingSettings.ToggleSwitch({
        name: "showImage",
        displayName: "Show avatars",
        value: true
    });

    cardAlignment = new FormattingSettingsDropdown({
        name: "cardAlignment",
        displayName: "Content alignment",
        value: { value: "start", displayName: "Left" },
        items: [
            { value: "start", displayName: "Left" },
            { value: "center", displayName: "Center" },
            { value: "end", displayName: "Right" }
        ]
    });

    name: string = "card";
    displayName: string = "Card";
    slices: FormattingSettingsSlice[] = [
        this.backgroundColor,
        this.accentColor,
        this.borderColor,
        this.textColor,
        this.borderWidth,
        this.borderRadius,
        this.showShadow,
        this.showImage,
        this.cardAlignment
    ];
}

class LabelSettings extends FormattingSettingsCard {
    nameTextSize = new formattingSettings.NumUpDown({
        name: "nameTextSize",
        displayName: "Name size",
        value: 16
    });

    titleTextSize = new formattingSettings.NumUpDown({
        name: "titleTextSize",
        displayName: "Title size",
        value: 12
    });

    detailTextSize = new formattingSettings.NumUpDown({
        name: "detailTextSize",
        displayName: "Detail size",
        value: 11
    });

    name: string = "labels";
    displayName: string = "Labels";
    slices: FormattingSettingsSlice[] = [
        this.nameTextSize,
        this.titleTextSize,
        this.detailTextSize
    ];
}

class LinkSettings extends FormattingSettingsCard {
    linkColor = new formattingSettings.ColorPicker({
        name: "linkColor",
        displayName: "Link color",
        value: { value: "#94a3b8" }
    });

    linkWidth = new formattingSettings.NumUpDown({
        name: "linkWidth",
        displayName: "Link width",
        value: 1.8
    });

    linkOpacity = new formattingSettings.NumUpDown({
        name: "linkOpacity",
        displayName: "Link opacity",
        value: 0.8
    });

    name: string = "links";
    displayName: string = "Links";
    slices: FormattingSettingsSlice[] = [
        this.linkColor,
        this.linkWidth,
        this.linkOpacity
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    layout = new LayoutCardSettings();
    card = new CardStyleSettings();
    labels = new LabelSettings();
    links = new LinkSettings();

    cards = [this.layout, this.card, this.labels, this.links];
}
