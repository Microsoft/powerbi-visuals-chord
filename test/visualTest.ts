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

// d3
import { select as d3select } from "d3";
import ChordGroup = d3.ChordGroup;

import powerbiVisualsApi from "powerbi-visuals-api";
import DataView = powerbiVisualsApi.DataView;
import DataViewValueColumn = powerbiVisualsApi.DataViewValueColumn;
import PrimitiveValue = powerbiVisualsApi.PrimitiveValue;
import DataViewCategoryColumn = powerbiVisualsApi.DataViewCategoryColumn;
import DataViewObjects = powerbiVisualsApi.DataViewObjects;
import ISelectionId = powerbiVisualsApi.visuals.ISelectionId;

// powerbi.extensibility.utils.interactivity
import { interactivitySelectionService } from "powerbi-visuals-utils-interactivityutils";
import SelectableDataPoint = interactivitySelectionService.SelectableDataPoint;

// powerbi.extensibility.utils.test
import { assertColorsMatch } from "powerbi-visuals-utils-testutils";

import { ChordChart, ChordChartData as ChordChartDataInterface } from "../src/chordChart";

import { ChordChartData } from "./chordChartData";
import { ChordChartBuilder } from "./visualBuilder";
import {
    areColorsEqual,
    IsInRangeFunction,
    getRandomUniqueHexColors,
    isSomeTextElementOverlapped,
    isTextElementInOrOutElement,
    getSolidColorStructuralObject
} from "./helpers/helpers";

import {
    sum as lodashSum,
    range as lodashRange,
    isNumber as lodashIsNumber
} from "lodash";

describe("ChordChart", () => {
    let visualBuilder: ChordChartBuilder,
        defaultDataViewBuilder: ChordChartData,
        dataView: DataView;

    beforeEach(() => {
        visualBuilder = new ChordChartBuilder(1000, 500);
        defaultDataViewBuilder = new ChordChartData();

        dataView = defaultDataViewBuilder.getDataView();
    });

    describe("DOM tests", () => {
        it("svg element created", () => {
            expect(visualBuilder.mainElement[0]).toBeInDOM();
        });

        it("update", (done) => {
            visualBuilder.updateRenderTimeout(dataView, () => {
                const valuesLength: number = lodashSum(dataView
                    .categorical
                    .values
                    .map((column: DataViewValueColumn) => {
                        const notEmptyValues: PrimitiveValue[] = column.values.filter((value: number) => {
                            return !isNaN(value) && value !== null;
                        });

                        return notEmptyValues.length;
                    }));

                const categoriesLength: number = dataView.categorical.values.length
                    + dataView.categorical.categories[0].values.length;

                expect(visualBuilder.mainElement.children("g.chords").children("path").length)
                    .toBe(valuesLength);

                expect(visualBuilder.mainElement.children("g.ticks").children("g.slice-ticks").length)
                    .toBe(categoriesLength);

                expect(visualBuilder.mainElement.children("g.slices").children("path.slice").length)
                    .toBe(categoriesLength);

                expect(visualBuilder.element.find(".chordChart").attr("height"))
                    .toBe(visualBuilder.viewport.height.toString());

                expect(visualBuilder.element.find(".chordChart").attr("width"))
                    .toBe(visualBuilder.viewport.width.toString());

                done();
            });
        });

        it("labels shouldn't be overlapped", (done) => {
            dataView.metadata.objects = {
                labels: {
                    show: true,
                    fontSize: 40
                }
            };

            visualBuilder.viewport.height = 100;
            visualBuilder.viewport.width = 1000;

            visualBuilder.updateRenderTimeout(dataView, () => {
                const isInRange: IsInRangeFunction = (
                    value: number,
                    min: number,
                    max: number) => {

                    return value >= min && value <= max;
                };

                expect(isSomeTextElementOverlapped(
                    visualBuilder.dataLabels.toArray(),
                    isInRange)).toBeFalsy();

                done();
            }, 50);
        });

        it("shouldn't throw any unexpected exceptions when category value is null", () => {

            defaultDataViewBuilder.valuesCategoryGroup[5][0] = null;
            expect(() => {
                ChordChart.CONVERTER(
                    defaultDataViewBuilder.getDataView(),
                    visualBuilder.visualHost,
                    visualBuilder.visualHost.colorPalette,
                    null
                );
            }).not.toThrow();

        });

        it("labels shouldn't be cut off", (done) => {
            visualBuilder.viewport.height = 200;
            visualBuilder.viewport.width = 200;

            defaultDataViewBuilder.valuesValue = lodashRange(1, defaultDataViewBuilder.valuesCategoryGroup.length);

            dataView = defaultDataViewBuilder.getDataView();

            dataView.metadata.objects = {
                labels: {
                    show: true
                }
            };

            visualBuilder.updateRenderTimeout(dataView, () => {
                expect(visualBuilder.mainElement.children("g.labels")[0].getBoundingClientRect().left)
                    .toBeGreaterThan(0);

                visualBuilder.dataLabels
                    .toArray()
                    .forEach((element: Element) => {
                        expect(isTextElementInOrOutElement(
                            visualBuilder.mainElement[0],
                            element,
                            (firstValue: number, secondValue: number) => firstValue >= secondValue)).toBeTruthy();
                    });

                done();
            });
        });

        it("labels shouldn't be visible on right side", (done) => {
            visualBuilder.viewport.height = 500;
            visualBuilder.viewport.width = 500;

            defaultDataViewBuilder.valuesCategoryGroup = lodashRange(20).map((value: number) => {
                return [
                    value + "xxxxxxxxxxx",
                    value + "yyyyyyyyyyyyyy"
                ];
            });

            defaultDataViewBuilder.valuesValue =
                lodashRange(1, defaultDataViewBuilder.valuesCategoryGroup.length);

            dataView = defaultDataViewBuilder.getDataView();

            dataView.metadata.objects = {
                labels: {
                    show: true,
                    fontSize: 40
                }
            };

            visualBuilder.updateRenderTimeout(dataView, () => {
                const rightLabels: JQuery = visualBuilder.dataLabels
                    .filter((i: number, element: Element) => {
                        return parseFloat($(element).attr("x")) > 0;
                    });

                expect(rightLabels).toBeInDOM();

                done();
            });
        });
    });

    describe("Format settings test", () => {
        describe("Axis", () => {
            beforeEach(() => {
                dataView.metadata.objects = {
                    axis: {
                        show: true
                    }
                };
            });

            it("show", () => {
                visualBuilder.updateFlushAllD3Transitions(dataView);

                expect(visualBuilder.sliceTicks).toBeInDOM();

                (<any>dataView.metadata.objects).axis.show = false;
                visualBuilder.updateFlushAllD3Transitions(dataView);

                expect(visualBuilder.sliceTicks).not.toBeInDOM();
            });
        });

        describe("Labels", () => {
            beforeEach(() => {
                dataView.metadata.objects = {
                    labels: {
                        show: true
                    }
                };
            });

            it("show", () => {
                visualBuilder.updateFlushAllD3Transitions(dataView);

                expect(visualBuilder.dataLabels).toBeInDOM();

                (<any>dataView.metadata.objects).labels.show = false;
                visualBuilder.updateFlushAllD3Transitions(dataView);

                expect(visualBuilder.dataLabels).not.toBeInDOM();
            });

            it("color", () => {
                const color: string = "#222222";

                (<any>dataView.metadata.objects).labels.color = getSolidColorStructuralObject(color);
                visualBuilder.updateFlushAllD3Transitions(dataView);

                visualBuilder.dataLabels
                    .toArray()
                    .forEach((element: Element) => {
                        assertColorsMatch($(element).css("fill"), color);
                    });
            });

            it("font size", () => {
                const fontSize: number = 22,
                    expectedFontSize: string = "29.3333px";

                (<any>dataView.metadata.objects).labels.fontSize = fontSize;
                visualBuilder.updateFlushAllD3Transitions(dataView);

                visualBuilder.dataLabels
                    .toArray()
                    .forEach((element: Element) => {
                        expect($(element).css("font-size")).toBe(expectedFontSize);
                    });
            });
        });

        describe("Data colors", () => {
            it("default color", () => {
                const color: string = "#222222";

                dataView.metadata.objects = {
                    dataPoint: {
                        defaultColor: getSolidColorStructuralObject(color)
                    }
                };

                visualBuilder.updateFlushAllD3Transitions(dataView);

                visualBuilder.chords
                    .toArray()
                    .forEach((element: Element) => {
                        assertColorsMatch($(element).css("fill"), color);
                    });
            });

            it("colors", () => {
                dataView.metadata.objects = {
                    dataPoint: {
                        showAllDataPoints: true
                    }
                };

                const category: DataViewCategoryColumn = dataView.categorical.categories[0],
                    colors: string[] = getRandomUniqueHexColors(category.values.length);

                category.objects = [];

                category.values.forEach((value: PrimitiveValue, index: number) => {
                    category.objects[index] = <DataViewObjects>{
                        dataPoint: {
                            fill: getSolidColorStructuralObject(colors[index])
                        }
                    };
                });

                visualBuilder.updateFlushAllD3Transitions(dataView);

                const slices: JQuery<any>[] = visualBuilder.slices.toArray().map($),
                    chords: JQuery<any>[] = visualBuilder.chords.toArray().map($);

                colors.forEach((color: string) => {
                    expect(doElementsUseColor(slices, color)).toBeTruthy();
                    expect(doElementsUseColor(chords, color)).toBeTruthy();
                });
            });

            function doElementsUseColor(elements: JQuery[], color: string): boolean {
                return elements.some((element: JQuery) => {
                    return areColorsEqual(element.css("fill"), color);
                });
            }
        });
    });

    describe("copyArcDescriptorsWithoutNaNValues", () => {
        it("shouldn't throw any unexpected exceptions when argument is undefined", () => {
            expect(() => {
                ChordChart.COPY_ARC_DESCRIPTORS_WITHOUT_NAN_VALUES(undefined);
            }).not.toThrow();
        });

        it("shouldn't throw any unexpected exceptions when argument is null", () => {
            expect(() => {
                ChordChart.COPY_ARC_DESCRIPTORS_WITHOUT_NAN_VALUES(null);
            }).not.toThrow();
        });

        it("result of removeNaNValues shouldn't contain any NaN values", () => {
            const arcDescriptors: ChordGroup[] = ChordChart.COPY_ARC_DESCRIPTORS_WITHOUT_NAN_VALUES(
                createArcDescriptorsWithNaN(5));

            arcDescriptorsShouldntContainNaNValues(arcDescriptors);
        });

        function createArcDescriptorsWithNaN(length: number) {
            const valueNaN: number = NaN,
                arcDescriptors: ChordGroup[] = [];

            for (let i: number = 0; i < length; i++) {
                arcDescriptors.push({
                    value: valueNaN,
                    startAngle: valueNaN,
                    endAngle: valueNaN,
                    index: valueNaN
                });
            }

            return arcDescriptors;
        }

        it("groups shouldn't contain any NaN values", () => {
            let chordChartData: ChordChartDataInterface;

            defaultDataViewBuilder.valuesValue = defaultDataViewBuilder.valuesValue.map(() => {
                return 0;
            });

            chordChartData = ChordChart.CONVERTER(
                defaultDataViewBuilder.getDataView(),
                visualBuilder.visualHost,
                visualBuilder.visualHost.colorPalette,
                null
            );

            arcDescriptorsShouldntContainNaNValues(chordChartData.groups);
        });

        function arcDescriptorsShouldntContainNaNValues(arcDescriptors: ChordGroup[]): void {
            arcDescriptors.forEach((arcDescriptor: ChordGroup) => {
                for (let propertyName of Object.keys(arcDescriptor)) {
                    if (lodashIsNumber(arcDescriptor[propertyName])) {
                        expect(isNaN(arcDescriptor[propertyName])).toBeFalsy();
                    }
                }
            });
        }
    });

    describe("check if values absent", () => {
        it("shouldn't throw any unexpected exceptions when Values field is undefined", () => {
            expect(() => {
                let chordChartData: ChordChartDataInterface = ChordChart.CONVERTER(
                    defaultDataViewBuilder.getDataView(null, true),
                    visualBuilder.visualHost,
                    visualBuilder.visualHost.colorPalette,
                    null
                );
            }).not.toThrow();
        });
    });

    describe("Capabilities tests", () => {
        it("all items having displayName should have displayNameKey property", () => {
            jasmine.getJSONFixtures().fixturesPath = "base";

            let jsonData = getJSONFixture("capabilities.json");

            let objectsChecker: Function = (obj) => {
                for (let property of Object.keys(obj)) {
                    let value: any = obj[property];

                    if (value.displayName) {
                        expect(value.displayNameKey).toBeDefined();
                    }

                    if (typeof value === "object") {
                        objectsChecker(value);
                    }
                }
            };

            objectsChecker(jsonData);
        });
    });

    describe("Selection", () => {
        describe("Power BI Bookmarks", () => {
            it("first identity should be selected", (done) => {
                visualBuilder.updateRenderTimeout(dataView, () => {
                    const firstSelectionId: ISelectionId = <ISelectionId>visualBuilder.instance["data"]["groups"][0]["identity"];

                    visualBuilder.selectionManager.sendSelectionToCallback([firstSelectionId]);

                    const isSelected: boolean = (<SelectableDataPoint>d3select(visualBuilder.slices.get(0)).datum()).selected;

                    expect(isSelected).toBeTruthy();

                    done();
                });
            });

        });
    });

    describe("Accessibility", () => {
        describe("High contrast mode", () => {
            const backgroundColor: string = "#000000";
            const foregroundColor: string = "#ffff00";

            beforeEach(() => {
                visualBuilder.visualHost.colorPalette.isHighContrast = true;

                visualBuilder.visualHost.colorPalette.background = { value: backgroundColor };
                visualBuilder.visualHost.colorPalette.foreground = { value: foregroundColor };
            });

            it("should not use fill style", (done) => {
                visualBuilder.updateRenderTimeout(dataView, () => {
                    const slices: JQuery<any>[] = visualBuilder.slices.toArray().map($);
                    const chords: JQuery<any>[] = visualBuilder.chords.toArray().map($);

                    expect(isColorAppliedToElements(slices, null, "fill"));
                    expect(isColorAppliedToElements(chords, null, "fill"));

                    done();
                });
            });

            it("should use stroke style", (done) => {
                visualBuilder.updateRenderTimeout(dataView, () => {
                    const slices: JQuery<any>[] = visualBuilder.slices.toArray().map($);
                    const chords: JQuery<any>[] = visualBuilder.chords.toArray().map($);

                    expect(isColorAppliedToElements(slices, foregroundColor, "stroke"));
                    expect(isColorAppliedToElements(chords, foregroundColor, "stroke"));

                    done();
                });
            });

            function isColorAppliedToElements(
                elements: JQuery[],
                color?: string,
                colorStyleName: string = "fill"
            ): boolean {
                return elements.some((element: JQuery) => {
                    const currentColor: string = element.css(colorStyleName);

                    if (!currentColor || !color) {
                        return currentColor === color;
                    }

                    return areColorsEqual(currentColor, color);
                });
            }
        });
    });
});
