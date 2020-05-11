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
import "./../style/chordChart.less";
import "@babel/polyfill";

// d3
import * as d3 from "d3";
import Arc = d3.Arc;
import ChordLayout = d3.ChordLayout;
import Selection = d3.Selection;
import ChordGroup = d3.ChordGroup;

// powerbi
import powerbiVisualsApi from "powerbi-visuals-api";
import DataView = powerbiVisualsApi.DataView;
import IViewport = powerbiVisualsApi.IViewport;
import DataViewObjects = powerbiVisualsApi.DataViewObjects;
import DataViewValueColumn = powerbiVisualsApi.DataViewValueColumn;
import VisualObjectInstance = powerbiVisualsApi.VisualObjectInstance;
import DataViewMetadataColumn = powerbiVisualsApi.DataViewMetadataColumn;
import EnumerateVisualObjectInstancesOptions = powerbiVisualsApi.EnumerateVisualObjectInstancesOptions;
import DataViewValueColumnGroup = powerbiVisualsApi.DataViewValueColumnGroup;
import PrimitiveValue = powerbiVisualsApi.PrimitiveValue;
import VisualObjectInstanceEnumeration = powerbiVisualsApi.VisualObjectInstanceEnumeration;
import ISelectionId = powerbiVisualsApi.visuals.ISelectionId;

// powerbi.extensibility
import IColorPalette = powerbiVisualsApi.extensibility.IColorPalette;
import VisualTooltipDataItem = powerbiVisualsApi.extensibility.VisualTooltipDataItem;
import IVisual = powerbiVisualsApi.extensibility.IVisual;
import IVisualHost = powerbiVisualsApi.extensibility.visual.IVisualHost;
import ILocalizationManager = powerbiVisualsApi.extensibility.ILocalizationManager;
import VisualConstructorOptions = powerbiVisualsApi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbiVisualsApi.extensibility.visual.VisualUpdateOptions;

// powerbi.extensibility.utils.dataview
import { converterHelper as ConverterHelper } from "powerbi-visuals-utils-dataviewutils";
import converterHelper = ConverterHelper.converterHelper;

// powerbi.extensibility.utils.svg
import { manipulation, IMargin, IRect, shapes, CssConstants, shapesInterfaces } from "powerbi-visuals-utils-svgutils";
import translate = manipulation.translate;
import ClassAndSelector = CssConstants.ClassAndSelector;
import createClassAndSelector = CssConstants.createClassAndSelector;
import translateAndRotate = manipulation.translateAndRotate;

// powerbi.extensibility.utils.color
import { ColorHelper } from "powerbi-visuals-utils-colorutils";

// powerbi.extensibility.utils.chart
import { dataLabelInterfaces, legendInterfaces, DataLabelManager, DataLabelArrangeGrid, dataLabelUtils } from "powerbi-visuals-utils-chartutils";
import LabelEnabledDataPoint = dataLabelInterfaces.LabelEnabledDataPoint;
import LegendData = legendInterfaces.LegendData;
import ILabelLayout = dataLabelInterfaces.ILabelLayout;
import IDataLabelInfo = dataLabelInterfaces.IDataLabelInfo;

// powerbi.extensibility.utils.formatting
import { valueFormatter as ValueFormatter } from "powerbi-visuals-utils-formattingutils";
import IValueFormatter = ValueFormatter.IValueFormatter;
import create = ValueFormatter.create;
import getFormatStringByColumn = ValueFormatter.getFormatStringByColumn;

// powerbi.extensibility.utils.type
import { pixelConverter as PixelConverter, double as TypeUtilsDouble } from "powerbi-visuals-utils-typeutils";

import lessWithPrecision = TypeUtilsDouble.lessWithPrecision;

// powerbi.extensibility.utils.interactivity
import { interactivitySelectionService, interactivityBaseService } from "powerbi-visuals-utils-interactivityutils";
import SelectableDataPoint = interactivitySelectionService.SelectableDataPoint;
import IInteractivityService = interactivityBaseService.IInteractivityService;
import createInteractivitySelectionService = interactivitySelectionService.createInteractivitySelectionService;

// powerbi.extensibility.utils.tooltip
import { TooltipEventArgs, ITooltipServiceWrapper, createTooltipServiceWrapper } from "powerbi-visuals-utils-tooltiputils";

import { Settings, AxisSettings, DataPointSettings, LabelsSettings } from "./settings";
import { ChordArcDescriptor, ChordArcLabelData } from "./interfaces";
import { VisualLayout } from "./visualLayout";
import { InteractiveBehavior, BehaviorOptions } from "./interactiveBehavior";
import { ChordChartColumns, ChordChartCategoricalColumns } from "./columns";
import { createTooltipInfo } from "./tooltipBuilder";
import { ChordChartHelpers } from "./chordChartHelpers";

import {
    mapValues as lodashMapValues,
    invert as lodashInvert,
    isEmpty as lodashIsEmpty,
    reduce as lodashReduce,
    max as lodashMax,
    map as lodashMap,
    forEach as lodashForEach,
    isArray as lodashIsArray
} from "lodash";
import { merge } from "d3";

export interface ChordChartData {
    settings: Settings;
    dataView: DataView;
    dataMatrix: number[][];
    tooltipData: ChordTooltipData[][];
    sliceTooltipData: ChordTooltipData[];
    tickUnit: number;
    differentFromTo: boolean;
    defaultDataPointColor?: string;
    prevAxisVisible: boolean;

    groups: ChordArcDescriptor[];
    chords: any[];
}

export type ChordChartCategoricalDict = {};

export interface ChordLabelEnabledDataPoint extends LabelEnabledDataPoint {
    data?: ChordArcLabelData;
}

export interface ChordTooltipData {
    tooltipInfo: VisualTooltipDataItem[];
}

export class ChordChart implements IVisual {
    private static ChordLayoutPadding: number = 0.1;
    private static DefaultMargin: IMargin = { left: 10, right: 10, top: 10, bottom: 10 };
    private static DefaultViewPort: IViewport = { width: 150, height: 150 };
    private static InnerArcRadiusRatio = 0.8;
    private static InnerLinePointMultiplier: number = 2.05;
    private static LabelMargin = 10;
    private static MaxValue: number = 1000;
    private static OuterArcRadiusRatio = 0.9;
    private static PolylineOpacity = 0.5;
    private static TicksFontSize = 12;
    private static VisualClassName = "chordChart";
    private static DefaultDY: string = ".35em";
    private static DefaultTickShiftX: number = 8;
    private static MaxUnitSize: number = 5;
    private static DefaultFormatValue: string = "0.##";
    private static DefaultTickLineColorValue: string = "#000";

    private static chordClass: ClassAndSelector = createClassAndSelector("chord");
    private static chordsClass: ClassAndSelector = createClassAndSelector("chords");
    private static labelGraphicsContextClass: ClassAndSelector = createClassAndSelector("labels");
    private static labelsClass: ClassAndSelector = createClassAndSelector("data-labels");
    private static lineClass: ClassAndSelector = createClassAndSelector("line-label");
    private static linesGraphicsContextClass: ClassAndSelector = createClassAndSelector("lines");
    private static polylineClass: ClassAndSelector = createClassAndSelector("polyline");
    private static sliceClass: ClassAndSelector = createClassAndSelector("slice");
    private static sliceTicksClass: ClassAndSelector = createClassAndSelector("slice-ticks");
    private static tickLineClass: ClassAndSelector = createClassAndSelector("tick-line");
    private static tickPairClass: ClassAndSelector = createClassAndSelector("tick-pair");
    private static tickTextClass: ClassAndSelector = createClassAndSelector("tick-text");
    private static ticksClass: ClassAndSelector = createClassAndSelector("ticks");

    private labels: Selection<d3.BaseType, any, any, any>;
    private lines: Selection<d3.BaseType, any, any, any>;
    private mainGraphicsContext: Selection<d3.BaseType, any, any, any>;
    private slices: Selection<d3.BaseType, any, any, any>;
    private svg: Selection<d3.BaseType, any, any, any>;

    private colors: IColorPalette;
    private data: ChordChartData;
    private layout: VisualLayout;

    private duration: number = 250;

    private tooltipServiceWrapper: ITooltipServiceWrapper;

    private host: IVisualHost;

    private interactivityService: IInteractivityService<ChordArcDescriptor>;
    private interactiveBehavior: InteractiveBehavior;

    private localizationManager: ILocalizationManager;

    private get settings(): Settings {
        return this.data && this.data.settings;
    }

    private radius: number;
    private get innerRadius(): number {
        return this.radius * ChordChart.InnerArcRadiusRatio;
    }
    private get outerRadius(): number {
        return this.radius * ChordChart.OuterArcRadiusRatio;
    }

    private static convertCategoricalToArray(values: any[]): ChordChartCategoricalDict {
        return lodashMapValues(lodashInvert(values), (d: string) => parseFloat(d));
    }

    public static defaultValue1: number = 1;

    /**
     * @param dataView DataView object
     * @param host PBI Host object
     * @param colors Color Palette from PBI
     * @param prevAxisVisible Indicates if the previous axis is visible
     * @param localizationManager Localization Manager
     */
    // tslint:disable-next-line: max-func-body-length
    public static CONVERTER(
        dataView: DataView,
        host: IVisualHost,
        colors: IColorPalette,
        localizationManager: ILocalizationManager
    ): ChordChartData {
        const settings: Settings = Settings.PARSE_SETTINGS(dataView, colors);
        const columns: ChordChartColumns<ChordChartCategoricalColumns> = ChordChartColumns.GET_CATEGORICAL_COLUMNS(dataView);
        const sources: ChordChartColumns<DataViewMetadataColumn> = ChordChartColumns.GET_COLUMN_SOURCES(dataView);
        const categoricalValues: ChordChartColumns<any> = ChordChartColumns.GET_CATEGORICAL_VALUES(dataView);
        const prevAxisVisible = settings.axis.show;

        if (!categoricalValues || lodashIsEmpty(categoricalValues.Category)) {
            return null;
        }

        if (!categoricalValues.Y) {
            categoricalValues.Y = [];
            for (let i: number = 0; i < categoricalValues.Category.length; i++) {
                categoricalValues.Series.push(sources.Series.displayName + i);
                categoricalValues.Y.push(ChordChart.defaultValue1);
            }
        }

        categoricalValues.Series = categoricalValues.Series || ChordChartColumns.GET_SERIES_VALUES(dataView);
        let grouped: DataViewValueColumnGroup[] = null;
        if (columns.Series) {
            grouped = columns.Series.grouped();
        }

        let dataMatrix: number[][] = [];
        let renderingDataMatrix: number[][] = [];
        let toolTipData: ChordTooltipData[][] = [];
        let sliceTooltipData: ChordTooltipData[] = [];
        let max: number = ChordChart.MaxValue;
        let seriesIndex: ChordChartCategoricalDict = this.convertCategoricalToArray(categoricalValues.Series); // series index array
        let catIndex: ChordChartCategoricalDict = this.convertCategoricalToArray(categoricalValues.Category); // index array for category names
        let isDiffFromTo: boolean = false;  // boolean variable indicates that From and To are different
        let labelData: ChordArcLabelData[] = [];    // label data: !important

        const colorHelper: ColorHelper = new ColorHelper(
            colors,
            { objectName: "dataPoint", propertyName: "fill" },
            settings.dataPoint.defaultColor
        );

        let totalFields: any[] = this.union_arrays(categoricalValues.Category, categoricalValues.Series).reverse();

        if (ChordChart.getValidArrayLength(totalFields) ===
            ChordChart.getValidArrayLength(categoricalValues.Category) + ChordChart.getValidArrayLength(categoricalValues.Series)) {
            isDiffFromTo = true;
        }

        let categoryColumnFormatter: IValueFormatter = create({
            format: getFormatStringByColumn(sources.Category, true)
                || sources.Category.format
        });
        let seriesColumnFormatter: IValueFormatter = create({
            format: sources.Series && (getFormatStringByColumn(sources.Series, true)
                || sources.Series.format)
        });
        let valueColumnFormatter: IValueFormatter = create({
            format: sources.Y ? getFormatStringByColumn(sources.Y, true)
                || sources.Y.format : "0"
        });

        const selectionIds: ISelectionId[] = [];

        for (let i: number = 0, iLength: number = totalFields.length; i < iLength; i++) {
            let selectionId: ISelectionId = null;
            let barFillColor: string = "";
            let isCategory: boolean = false;
            let index: number;

            const label: string = (sources.Series && i < categoricalValues.Series.length)
                ? seriesColumnFormatter.format(totalFields[i])
                : categoryColumnFormatter.format(totalFields[i]);

            if ((index = catIndex[totalFields[i]]) !== undefined) {
                selectionId = host.createSelectionIdBuilder()
                    .withCategory(columns.Category, index)
                    .createSelectionId();

                isCategory = true;

                let thisCategoryObjects: DataViewObjects = columns.Category.objects
                    ? columns.Category.objects[index]
                    : undefined;

                barFillColor = colorHelper.getColorForSeriesValue(thisCategoryObjects, categoricalValues.Category[index]);
            } else if ((index = seriesIndex[totalFields[i]]) !== undefined) {
                let seriesObjects: DataViewObjects = (grouped) ? grouped[index].objects : null;

                let seriesData: DataViewValueColumn = columns.Y ? columns.Y[index] : {
                    objects: null,
                    source: {
                        displayName: "Value",
                        queryName: "Value",
                        groupName: "Value",
                    },
                    values: [ChordChart.defaultValue1]
                };

                let seriesNameStr: PrimitiveValue = seriesData ? converterHelper.getSeriesName(seriesData.source) : "Value";

                selectionId = host.createSelectionIdBuilder()
                    .withSeries(columns.Series, (grouped) ? grouped[index] : null)
                    .withMeasure(seriesNameStr ? seriesNameStr.toString() : null)
                    .createSelectionId();
                isCategory = false;

                barFillColor = colorHelper.getColorForSeriesValue(seriesObjects, seriesNameStr ? seriesNameStr : `${ChordChart.defaultValue1}`);
            }

            const barStrokeColor: string = colorHelper.getHighContrastColor("foreground", barFillColor);

            selectionIds.push(selectionId);

            labelData.push({
                label,
                isCategory,
                barFillColor,
                barStrokeColor,
                labelColor: settings.labels.color,
                isGrouped: !!grouped,
                labelFontSize: PixelConverter.fromPointToPixel(settings.labels.fontSize),
            });

            renderingDataMatrix.push([]);
            dataMatrix.push([]);
            toolTipData.push([]);

            for (let j: number = 0, jLength: number = totalFields.length; j < jLength; j++) {
                let elementValue: number = 0;
                let tooltipInfo: VisualTooltipDataItem[] = [];

                if (catIndex[totalFields[i]] !== undefined &&
                    seriesIndex[totalFields[j]] !== undefined) {
                    let row: number = catIndex[totalFields[i]];
                    let col: number = seriesIndex[totalFields[j]];

                    if (columns.Y && columns.Y[col].values[row] !== null) {
                        elementValue = <number>columns.Y[col].values[row];

                        if (elementValue > max) {
                            max = elementValue;
                        }

                        tooltipInfo = createTooltipInfo(
                            dataView.categorical,
                            label,
                            valueColumnFormatter.format(elementValue),
                            col,
                            row,
                            localizationManager);
                    } else if (!columns.Y) {
                        max = ChordChart.defaultValue1;
                        elementValue = ChordChart.defaultValue1;
                        tooltipInfo = createTooltipInfo(
                            dataView.categorical,
                            label,
                            valueColumnFormatter.format(`${ChordChart.defaultValue1}`),
                            col,
                            row,
                            localizationManager);
                    }

                } else if (isDiffFromTo
                    && catIndex[totalFields[j]] !== undefined
                    && seriesIndex[totalFields[i]] !== undefined
                ) {
                    let row: number = catIndex[totalFields[j]];
                    let col: number = seriesIndex[totalFields[i]];

                    if (columns.Y && columns.Y[col].values[row] !== null) {
                        elementValue = <number>columns.Y[col].values[row];
                    } else if (!columns.Y) {
                        elementValue = ChordChart.defaultValue1;
                    }
                }

                renderingDataMatrix[i].push(Math.max(elementValue || 0, 0));
                dataMatrix[i].push(elementValue || 0);
                toolTipData[i].push({
                    tooltipInfo: tooltipInfo
                });
            }

            let totalSum: number = d3.sum(dataMatrix[i]);

            sliceTooltipData.push({
                tooltipInfo: [{
                    displayName: label,
                    value: valueColumnFormatter.format(totalSum)
                }]
            });
        }

        let chordLayout: ChordLayout = d3.chord();
        chordLayout.padAngle(ChordChart.ChordLayoutPadding);
        let chords: d3.Chords = chordLayout(renderingDataMatrix);

        const groups: ChordArcDescriptor[] = ChordChart.getChordArcDescriptors(
            ChordChart.COPY_ARC_DESCRIPTORS_WITHOUT_NAN_VALUES(chords.groups),
            labelData,
            selectionIds
        );

        const unitLength: number = Math.round(max / ChordChart.MaxUnitSize).toString().length - 1;

        return {
            dataMatrix: dataMatrix,
            dataView: dataView,
            settings: settings,
            tooltipData: toolTipData,
            sliceTooltipData: sliceTooltipData,
            tickUnit: Math.pow(10, unitLength),
            differentFromTo: isDiffFromTo,
            prevAxisVisible: prevAxisVisible === undefined
                ? settings.axis.show
                : prevAxisVisible,
            groups: groups,
            chords: chords
        };
    }

    // Check every element of the array and returns the count of elements which are valid(not undefined)
    private static getValidArrayLength(array: any[]): number {
        return lodashReduce(array, (total, value) => {
            return (value === undefined) ? total : total + 1;
        }, 0);
    }

    private static getChordArcDescriptors(
        groups: ChordGroup[],
        datum: ChordArcLabelData[],
        selectionIds: ISelectionId[]
    ): ChordArcDescriptor[] {
        groups.forEach((arcDescriptor: ChordArcDescriptor, index: number) => {
            arcDescriptor.data = datum[index];
            arcDescriptor.identity = selectionIds[index];
        });

        return <ChordArcDescriptor[]>groups;
    }

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;

        this.interactivityService = createInteractivitySelectionService(this.host);
        this.interactiveBehavior = new InteractiveBehavior();

        this.localizationManager = this.host.createLocalizationManager();

        this.tooltipServiceWrapper = createTooltipServiceWrapper(
            this.host.tooltipService,
            options.element);

        this.layout = new VisualLayout(ChordChart.DefaultViewPort, ChordChart.DefaultMargin);
        this.layout.minViewport = ChordChart.DefaultViewPort;

        this.svg = d3.select(options.element)
            .append("svg")
            .style("position", "absolute")
            .classed(ChordChart.VisualClassName, true);

        let svgSelection: Selection<d3.BaseType, any, any, any> = this.mainGraphicsContext = this.svg
            .append("g");

        svgSelection
            .append("g")
            .classed("chords", true);

        this.slices = svgSelection
            .append("g")
            .classed("slices", true);

        svgSelection
            .append("g")
            .classed(ChordChart.ticksClass.className, true);

        this.labels = svgSelection
            .append("g")
            .classed(ChordChart.labelGraphicsContextClass.className, true);

        this.lines = svgSelection
            .append("g")
            .classed(ChordChart.linesGraphicsContextClass.className, true);

        this.colors = options.host.colorPalette;
    }

    // Called for data, size, formatting changes
    public update(options: VisualUpdateOptions): void {
        this.host.eventService.renderingStarted(options);
        try {
            // assert dataView
            if (!options.dataViews || !options.dataViews[0]) {
                return;
            }

            this.layout.viewport = options.viewport;

            this.layout.viewport = options.viewport;

            this.data = ChordChart.CONVERTER(
                options.dataViews[0],
                this.host,
                this.colors,
                this.localizationManager);

            if (!this.data) {
                this.clear();

                return;
            }

            this.layout.resetMargin();
            this.layout.margin.top
                = this.layout.margin.bottom
                = PixelConverter.fromPointToPixel(this.settings.labels.fontSize) / 2;

            this.render();
            this.host.eventService.renderingFinished(options);
        }
        catch (e) {
            this.host.eventService.renderingFailed(options, e);
        }
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
        if (!this.data || !this.settings) {
            return [];
        }

        switch (options.objectName) {
            case "axis": {
                return ChordChart.enumerateAxis(this.settings, this.localizationManager);
            }
            case "dataPoint": {
                return ChordChart.enumerateDataPoint(this.settings, this.data.groups, this.localizationManager);
            }
            case "labels": {
                return ChordChart.enumerateLabels(this.settings, this.localizationManager);
            }
            default: {
                return [];
            }
        }
    }

    private static enumerateAxis(settings: Settings, localManager: ILocalizationManager): VisualObjectInstance[] {
        let axisSettings: AxisSettings = settings.axis,
            instances: VisualObjectInstance[] = [{
                objectName: "axis",
                displayName: localManager.getDisplayName("Visual_Axis"),
                selector: null,
                properties: {
                    show: axisSettings.show
                }
            }];
        return instances;
    }

    private static enumerateDataPoint(
        settings: Settings,
        labelDataPoints: ChordArcDescriptor[],
        localManager: ILocalizationManager
    ): VisualObjectInstance[] {
        const dataPointSettings: DataPointSettings = settings.dataPoint;

        let instances: VisualObjectInstance[] = [{
            objectName: "datapoint",
            displayName: localManager.getDisplayName("Visual_Data_Colors"),
            selector: null,
            properties: {
                defaultColor: dataPointSettings.defaultColor,
                showAllDataPoints: dataPointSettings.showAllDataPoints
            }
        }];

        if (!dataPointSettings.showAllDataPoints) {
            return instances;
        }

        for (let labelDataPoint of labelDataPoints) {
            let data: ChordArcLabelData = labelDataPoint.data;

            let colorInstance: VisualObjectInstance = {
                objectName: "dataPoint",
                displayName: data.label,
                selector: ColorHelper.normalizeSelector((<ISelectionId>labelDataPoint.identity).getSelector()),
                properties: {
                    fill: { solid: { color: data.barFillColor } }
                }
            };

            if (data.isCategory || data.isGrouped) {
                instances.push(colorInstance);
            }
        }

        return instances;
    }

    private static enumerateLabels(settings: Settings, localManager: ILocalizationManager): VisualObjectInstance[] {
        const labelSettings: LabelsSettings = settings.labels;

        return [{
            objectName: "labels",
            displayName: localManager.getDisplayName("Visual_Labels"),
            selector: null,
            properties: {
                show: labelSettings.show,
                color: labelSettings.color,
                fontSize: labelSettings.fontSize
            }
        }];
    }

    // Calculate radius
    private calculateRadius(): number {
        if (this.settings.labels.show) {
            // if we have category or data labels, use a sigmoid to blend the desired denominator from 2 to 3.
            // if we are taller than we are wide, we need to use a larger denominator to leave horizontal room for the labels.
            let hw: number = this.layout.viewportIn.height / this.layout.viewportIn.width;
            let denom: number = 2 + (1 / (1 + Math.exp(-5 * (hw - 1))));
            return Math.min(this.layout.viewportIn.height, this.layout.viewportIn.width) / denom;
        }

        // no labels
        return Math.min(this.layout.viewportIn.height, this.layout.viewportIn.width) / 2;
    }

    private drawCategoryLabels(): void {
        // Multiplier to place the end point of the reference line at 0.05 * radius away from the outer edge of the chord/pie.
        let arc: Arc<any, d3.DefaultArcObject> = d3.arc()
            .innerRadius(0)
            .outerRadius(this.innerRadius);

        let outerArc: Arc<any, d3.DefaultArcObject> = d3.arc()
            .innerRadius(this.outerRadius)
            .outerRadius(this.outerRadius);

        if (this.settings.labels.show) {
            let labelLayout: ILabelLayout = this.getChordChartLabelLayout(outerArc);
            let filteredData: LabelEnabledDataPoint[] = this.getDataLabelManager().hideCollidedLabels(
                this.layout.viewportIn,
                this.data.groups,
                labelLayout,
                    /* addTransform */ true);

            this.renderLabels(filteredData, labelLayout, true);
            this.renderLines(filteredData, arc, outerArc);
        }
        else {
            dataLabelUtils.cleanDataLabels(this.labels);
            dataLabelUtils.cleanDataLabels(this.lines, true);
        }
    }

    private getDataLabelManager(): DataLabelManager {
        let dataLabelManager = new DataLabelManager();
        (<any>dataLabelManager).hasCollisions = hasCollisions.bind(dataLabelManager);
        return dataLabelManager;

        function hasCollisions(arrangeGrid: DataLabelArrangeGrid, info: IDataLabelInfo, position: IRect, size: shapesInterfaces.ISize) {
            if (arrangeGrid.hasConflict(position)) {
                return true;
            }

            let intersection = { left: 0, top: position.height / 2, width: size.width, height: size.height };
            intersection = shapes.Rect.inflate(intersection, {
                left: DataLabelManager.InflateAmount,
                top: 0,
                right: DataLabelManager.InflateAmount,
                bottom: 0
            });
            intersection = shapes.Rect.intersect(intersection, position);

            if (shapes.Rect.isEmpty(intersection)) {
                return true;
            }

            return lessWithPrecision(intersection.height, position.height / 2);
        }
    }

    private render(): void {
        this.radius = this.calculateRadius();

        let arc: Arc<any, d3.DefaultArcObject> = d3.arc()
            .innerRadius(this.radius)
            .outerRadius(this.innerRadius);

        this.svg
            .attr("width", this.layout.viewport.width)
            .attr("height", this.layout.viewport.height);

        this.mainGraphicsContext
            .attr("transform", translate(this.layout.viewport.width / 2, this.layout.viewport.height / 2));

        let sliceShapes: Selection<d3.BaseType, ChordArcDescriptor, any, any> = this.slices
            .selectAll("path" + ChordChart.sliceClass.selectorName)
            .data(this.getChordTicksArcDescriptors());

        sliceShapes.exit().remove();

        sliceShapes = sliceShapes.merge(sliceShapes
            .enter()
            .append("path")
            .classed(ChordChart.sliceClass.className, true));

        sliceShapes
            .style("fill", (d) => d.data.barFillColor)
            .style("stroke", (d) => d.data.barStrokeColor)
            .attr("d", d => arc(<any>d));


        this.tooltipServiceWrapper.addTooltip(
            sliceShapes,
            (tooltipEvent: TooltipEventArgs<ChordArcDescriptor>) => {
                return this.data.sliceTooltipData[tooltipEvent.data.index].tooltipInfo;
            });

        let path: any = d3.ribbon().radius(this.radius);

        let chordShapes: Selection<d3.BaseType, any, any, any> = this.svg
            .select(ChordChart.chordsClass.selectorName)
            .selectAll(ChordChart.chordClass.selectorName)
            .data(this.data.chords);

        chordShapes.exit().remove();

        chordShapes = chordShapes.merge(chordShapes
            .enter()
            .append("path")
            .classed(ChordChart.chordClass.className, true));

        chordShapes
            .style("fill", (chordLink: any) => {
                return this.data.groups[chordLink.target.index].data.barFillColor;
            })
            .style("stroke", this.settings.chord.strokeColor)
            .style("stroke-width", PixelConverter.toString(this.settings.chord.strokeWidth))
            .attr("d", path);

        this.drawTicks();
        this.drawCategoryLabels();

        if (this.interactivityService) {
            this.interactivityService.applySelectionStateToData(this.data.groups);

            const behaviorOptions: BehaviorOptions = {
                clearCatcher: this.svg,
                arcSelection: sliceShapes,
                chordSelection: chordShapes,
                dataPoints: this.data.groups,
                behavior: this.interactiveBehavior
            };

            this.interactivityService.bind(behaviorOptions);
        }

        this.tooltipServiceWrapper.addTooltip(
            chordShapes,
            (tooltipEvent: TooltipEventArgs<any>) => {
                let tooltipInfo: VisualTooltipDataItem[] = [];

                if (this.data.differentFromTo) {
                    const { index, subindex } = tooltipEvent.data.source;

                    tooltipInfo = this.data.tooltipData[subindex][index].tooltipInfo;
                } else {
                    tooltipInfo.push(ChordChart.createTooltipInfo(
                        this.data.groups,
                        this.data.dataMatrix,
                        tooltipEvent.data.source));

                    tooltipInfo.push(ChordChart.createTooltipInfo(
                        this.data.groups,
                        this.data.dataMatrix,
                        tooltipEvent.data.target));
                }

                return tooltipInfo;
            });
    }

    private static createTooltipInfo(labelDataPoints: ChordArcDescriptor[], dataMatrix: number[][], source: any) {
        return {
            displayName: labelDataPoints[source.index].data.label
                + "->" + labelDataPoints[source.subindex].data.label,
            value: dataMatrix[source.index][source.subindex].toString()
        };
    }

    private clear(): void {
        this.clearNodes([
            ChordChart.chordClass,
            ChordChart.sliceClass,
            ChordChart.sliceTicksClass,
            ChordChart.labelsClass,
            ChordChart.lineClass
        ]);
    }

    private clearTicks(): void {
        this.clearNodes([
            ChordChart.tickLineClass,
            ChordChart.tickPairClass,
            ChordChart.tickTextClass,
            ChordChart.sliceTicksClass
        ]);
    }

    private clearNodes(selectors: ClassAndSelector | ClassAndSelector[]): void {
        selectors = lodashIsArray(selectors) ? selectors : <ClassAndSelector[]>[selectors];
        lodashForEach(selectors, (d: ClassAndSelector) => ChordChart.clearNode(this.mainGraphicsContext, d));
    }

    private static clearNode(selector: Selection<d3.BaseType, any, any, any>, d: ClassAndSelector): void {
        const empty: any[] = [];
        let selectors: Selection<d3.BaseType, any, any, any> = selector
            .selectAll(d.selectorName)
            .data(empty);

        selectors
            .exit()
            .remove();
    }

    private getChordTicksArcDescriptors(): ChordArcDescriptor[] {
        let groups: ChordGroup[] = this.data.groups;

        let maxValue: number = !lodashIsEmpty(groups) && lodashMax(lodashMap(groups, (x: ChordGroup) => x.value)) || 0;
        let minValue: number = !lodashIsEmpty(groups) && lodashMax(lodashMap(groups, (x: ChordGroup) => x.value)) || 0;

        let radiusCoeff: number = this.radius / Math.abs(maxValue - minValue) * 1.25;

        let formatter: IValueFormatter = create({
            format: ChordChart.DefaultFormatValue,
            value: maxValue
        });

        groups.forEach((x: ChordArcDescriptor) => {
            let k: number = (x.endAngle - x.startAngle) / x.value,
                absValue: number = Math.abs(x.value),
                range: number[] = d3.range(0, absValue, absValue - 1 < 0.15 ? 0.15 : absValue - 1);

            if (x.value < 0) {
                range = range.map(x => x * -1).reverse();
            }

            for (let i: number = 1; i < range.length; i++) {
                let gapSize: number = Math.abs(range[i] - range[i - 1]) * radiusCoeff;

                if (gapSize < ChordChart.TicksFontSize) {
                    if (range.length > 2 && i === range.length - 1) {
                        range.splice(--i, 1);
                    } else {
                        range.splice(i--, 1);
                    }
                }
            }

            x.angleLabels = range.map((v, i) => <any>{ angle: v * k + x.startAngle, label: formatter.format(v) });
        });

        return <ChordArcDescriptor[]>groups;
    }

    public static COPY_ARC_DESCRIPTORS_WITHOUT_NAN_VALUES(arcDescriptors: ChordGroup[]): ChordGroup[] {
        if (lodashIsEmpty(arcDescriptors)) {
            return arcDescriptors;
        }

        return arcDescriptors.map((sourceArcDescriptor: ChordGroup) => {
            let targetArcDescriptor: ChordGroup = <ChordGroup>{};

            for (let propertyName of Object.keys(sourceArcDescriptor)) {
                if (!sourceArcDescriptor[propertyName] && isNaN(sourceArcDescriptor[propertyName])) {
                    targetArcDescriptor[propertyName] = 0;
                } else {
                    targetArcDescriptor[propertyName] = sourceArcDescriptor[propertyName];
                }
            }

            return targetArcDescriptor;
        });
    }

    // Draw axis(ticks) around the arc
    private drawTicks(): void {
        if (this.settings.axis.show) {
            let animDuration: number = (this.data.prevAxisVisible === this.settings.axis.show)
                ? this.duration
                : 0;

            let tickShapes: Selection<d3.BaseType, any, any, any> = this.mainGraphicsContext
                .select(ChordChart.ticksClass.selectorName)
                .selectAll("g" + ChordChart.sliceTicksClass.selectorName)
                .data(this.data.groups);

            tickShapes
                .exit()
                .remove();

            tickShapes = tickShapes.merge(tickShapes
                .enter()
                .append("g")
                .classed(ChordChart.sliceTicksClass.className, true));

            let tickPairs = tickShapes
                .selectAll("g" + ChordChart.tickPairClass.selectorName)
                .data((d: ChordArcDescriptor) => d.angleLabels);

            tickPairs
                .exit()
                .remove();

            tickPairs = tickPairs.merge(tickPairs
                .enter()
                .append("g")
                .classed(ChordChart.tickPairClass.className, true));

            tickPairs
                .attr("transform", (d) => translateAndRotate(
                    this.innerRadius,
                    0,
                    -this.innerRadius,
                    0,
                    d.angle * 180 / Math.PI - 90)
                );

            let tickLines = tickPairs
                .selectAll("line" + ChordChart.tickLineClass.selectorName)
                .data((d) => [d]);

            tickLines
                .exit()
                .remove();
            tickLines = tickLines.merge(tickLines
                .enter()
                .append("line")
                .classed(ChordChart.tickLineClass.className, true));

            tickLines
                .style("stroke", ChordChart.DefaultTickLineColorValue)
                .attr("x1", 1)
                .attr("y1", 0)
                .attr("x2", 5)
                .attr("y2", 0)
                .merge(tickLines);

            let tickText = tickPairs
                .selectAll("text" + ChordChart.tickTextClass.selectorName)
                .data((d) => [d]);

            tickText
                .exit()
                .remove();

            tickText = tickText.merge(tickText
                .enter()
                .append("text"));

            tickText
                .classed(ChordChart.tickTextClass.className, true)
                .attr("x", ChordChart.DefaultTickShiftX)
                .attr("dy", ChordChart.DefaultDY)
                .text(d => (<any>d).label)
                .style("text-anchor", d => (<any>d).angle > Math.PI ? "end" : null)
                .style("fill", this.settings.axis.color)
                .attr("transform", d => (<any>d).angle > Math.PI ? "rotate(180)translate(-16)" : null);
        } else {
            this.clearTicks();
        }
    }

    private renderLabels(
        filteredData: ChordLabelEnabledDataPoint[],
        layout: ILabelLayout,
        isDonut: boolean = false,
        forAnimation: boolean = false): void {

        // Check for a case where resizing leaves no labels - then we need to remove the labels "g"
        if (filteredData.length === 0) {
            dataLabelUtils.cleanDataLabels(this.labels, true);

            return null;
        }

        // line chart ViewModel has a special "key" property for point identification since the "identity" field is set to the series identity
        let hasKey: boolean = (<any>filteredData)[0].key !== null;
        let hasDataPointIdentity: boolean = (<any>filteredData)[0].identity !== null;

        let dataLabels: Selection<d3.BaseType, ChordLabelEnabledDataPoint, any, any> = this.labels
            .selectAll(ChordChart.labelsClass.selectorName)
            .data(filteredData);

        dataLabels
            .exit()
            .remove();

        dataLabels = dataLabels.merge(dataLabels
            .enter()
            .append("text")
            .classed(ChordChart.labelsClass.className, true));

        let newLabels = dataLabels;

        if (forAnimation) {
            newLabels.style("opacity", 0);
        }

        dataLabels
            .attr("x", (d: LabelEnabledDataPoint) => d.labelX)
            .attr("y", (d: LabelEnabledDataPoint) => d.labelY)
            .attr("dy", ChordChart.DefaultDY)
            .text((d: LabelEnabledDataPoint) => d.labeltext);

        Object.keys(layout.style).forEach(x => dataLabels.style(x, layout.style[x]));
    }

    private renderLines(filteredData: ChordLabelEnabledDataPoint[], arc: Arc<any, d3.DefaultArcObject>, outerArc: Arc<any, d3.DefaultArcObject>): void {
        let lines: Selection<d3.BaseType, ChordLabelEnabledDataPoint, any, any> = this.lines
            .selectAll("polyline")
            .data(filteredData);

        let midAngle = (d: ChordArcDescriptor) => d.startAngle + (d.endAngle - d.startAngle) / 2;

        lines
            .exit()
            .remove();

        lines = lines.merge(lines
            .enter()
            .append("polyline")
            .classed(ChordChart.lineClass.className, true));

        lines
            .attr("points", (d: ChordArcDescriptor): any => {
                let textPoint: [number, number] = outerArc.centroid(<any>d);

                textPoint[0] = (this.radius + ChordChart.LabelMargin / 2) * (midAngle(d) < Math.PI ? 1 : -1);

                let midPoint: [number, number] = outerArc.centroid(<any>d);
                let chartPoint: [number, number] = arc.centroid(<any>d);

                chartPoint[0] *= ChordChart.InnerLinePointMultiplier;
                chartPoint[1] *= ChordChart.InnerLinePointMultiplier;

                return [
                    chartPoint,
                    midPoint,
                    textPoint
                ];
            })
            .style("opacity", ChordChart.PolylineOpacity)
            .style("stroke", (d: ChordArcDescriptor) => d.data.labelColor)
            .style("pointer-events", "none");
    }

    // Get label layout
    private getChordChartLabelLayout(outerArc: Arc<any, d3.DefaultArcObject>): ILabelLayout {
        let midAngle = (d: ChordArcDescriptor) => d.startAngle + (d.endAngle - d.startAngle) / 2;
        let maxLabelWidth: number = (this.layout.viewportIn.width - this.radius * 2 - ChordChart.LabelMargin * 2) / 1.6;

        return {
            labelText: (d: ChordLabelEnabledDataPoint) => {
                // show only category label
                return dataLabelUtils.getLabelFormattedText({
                    label: d.data.label,
                    maxWidth: maxLabelWidth,
                    fontSize: PixelConverter.fromPointToPixel(this.settings.labels.fontSize),
                });
            },
            labelLayout: {
                x: (d: ChordArcDescriptor) =>
                    (this.radius + ChordChart.LabelMargin) * (midAngle(d) < Math.PI ? 1 : -1),
                y: (d: ChordArcDescriptor) => {
                    let pos: [number, number] = outerArc.centroid(<any>d);
                    return pos[1];
                },
            },
            filter: (d: ChordArcDescriptor) => (d !== null && d.data !== null && d.data.label !== null),
            style: {
                "fill": (d: ChordArcDescriptor) => d.data.labelColor,
                "text-anchor": (d: ChordArcDescriptor) => midAngle(d) < Math.PI ? "start" : "end",
                "font-size": (d: ChordArcDescriptor) => PixelConverter.fromPoint(this.settings.labels.fontSize),
            },
        };
    }

    // Utility function for union two arrays without duplicates
    private static union_arrays(x: any[], y: any[]): any[] {
        let obj: Object = {};

        for (let i: number = 0; i < x.length; i++) {
            obj[x[i]] = x[i];
        }

        for (let i: number = 0; i < y.length; i++) {
            obj[y[i]] = y[i];
        }

        let res: string[] = [];

        for (let k of Object.keys(obj)) {
            if (obj.hasOwnProperty(k)) {  // <-- optional
                res.push(obj[k]);
            }
        }
        return res;
    }
}
