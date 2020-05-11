/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software""), to deal
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

// powerbi.extensibility.utils.color
import { ColorHelper } from "powerbi-visuals-utils-colorutils";

// powerbi.extensibility.utils.chart.dataLabel
import { dataLabelUtils } from "powerbi-visuals-utils-chartutils";

// powerbi.extensibility.utils.dataview
import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";
import DataViewObjectsParser = dataViewObjectsParser.DataViewObjectsParser;

// powerbi
import powerbiVisualsApi from "powerbi-visuals-api";
import DataView = powerbiVisualsApi.DataView;

// powerbi.extensibility
import IColorPalette = powerbiVisualsApi.extensibility.IColorPalette;

export class AxisSettings {
    public show: boolean = true;
    public color: string = "#212121";
}

export class DataPointSettings {
    public showAllDataPoints: boolean = false;
    public defaultColor: string = null;
}

export class LabelsSettings {
    public show: boolean = true;
    public color: string = dataLabelUtils.defaultLabelColor;
    public fontSize: number = dataLabelUtils.DefaultFontSizeInPt;
}

export class ChordSettings {
    public strokeColor: string = "#000000";
    public strokeWidth: number = 0.5;
    public strokeWidthMin: number = 0.5;
    public strokeWidthMax: number = 1;
}

export class Settings extends DataViewObjectsParser {
    public axis: AxisSettings = new AxisSettings();
    public dataPoint: DataPointSettings = new DataPointSettings();
    public labels: LabelsSettings = new LabelsSettings();
    public chord: ChordSettings = new ChordSettings();

    public static PARSE_SETTINGS(dataView: DataView, colorPalette?: IColorPalette): Settings {
        const settings: Settings = this.parse<Settings>(dataView);

        const colorHelper: ColorHelper = new ColorHelper(colorPalette);

        settings.axis.color = colorHelper.getHighContrastColor(
            "foreground",
            settings.axis.color
        );

        settings.dataPoint.defaultColor = colorHelper.getHighContrastColor(
            "background",
            settings.dataPoint.defaultColor
        );

        settings.labels.color = colorHelper.getHighContrastColor(
            "foreground",
            settings.labels.color
        );

        settings.chord.strokeColor = colorHelper.getHighContrastColor(
            "foreground",
            settings.chord.strokeColor
        );

        if (colorPalette && colorHelper.isHighContrast) {
            settings.chord.strokeWidth = settings.chord.strokeWidthMax;
        } else {
            settings.chord.strokeWidth = settings.chord.strokeWidthMin;
        }

        return settings;
    }
}
