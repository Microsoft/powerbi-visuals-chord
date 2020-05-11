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
import powerbiVisualsApi from "powerbi-visuals-api";
import DataView = powerbiVisualsApi.DataView;

// powerbi.extensibility.utils.test
import { getRandomNumbers, testDataViewBuilder } from "powerbi-visuals-utils-testutils";
import TestDataViewBuilder = testDataViewBuilder.TestDataViewBuilder;

// powerbi.extensibility.utils.type
import { valueType } from "powerbi-visuals-utils-typeutils";
import ValueType = valueType.ValueType;

export class ChordChartData extends TestDataViewBuilder {
    public static ColumnCategory: string = "Category";
    public static ColumnSeries: string = "Series";
    public static ColumnValues: string = "Y";

    public valuesCategoryGroup: string[][] = [
        ["William", "Aiden"],
        ["William", "Daniel"],
        ["William", "Harper"],

        ["Olivia", "Aiden"],
        ["Olivia", "Harper"],

        ["James", "Daniel"],

        ["Lucas", "Aiden"],
        ["Lucas", "Daniel"],

        ["Henry", "Aiden"],
        ["Henry", "Daniel"],
        ["Henry", "Harper"],
    ];

    public valuesValue: number[] = getRandomNumbers(this.valuesCategoryGroup.length, 50, 100);

    public getDataView(columnNames?: string[], emptyValues: boolean = false): DataView {
        return this.createCategoricalDataViewBuilder([
            {
                source: {
                    displayName: ChordChartData.ColumnCategory,
                    roles: { Category: true },
                    type: ValueType.fromDescriptor({ text: true })
                },
                values: emptyValues ? null : this.valuesCategoryGroup.map((value: string[]) => value[0])
            },
            {
                isGroup: true,
                source: {
                    displayName: ChordChartData.ColumnSeries,
                    roles: { Series: true },
                    type: ValueType.fromDescriptor({ text: true })
                },
                values: emptyValues ? null : this.valuesCategoryGroup.map((value: string[]) => value[1]),
            }
        ], [
                {
                    source: {
                        displayName: ChordChartData.ColumnValues,
                        isMeasure: true,
                        roles: { Y: true },
                        type: ValueType.fromDescriptor({ numeric: true }),
                        objects: { dataPoint: { fill: { solid: { color: "purple" } } } },
                    },
                    values: emptyValues ? null : this.valuesValue
                }], columnNames).build();
    }
}
