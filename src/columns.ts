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

// powerbi.extensibility.utils.dataview
import { converterHelper as ConverterHelper } from "powerbi-visuals-utils-dataviewutils";
import converterHelper = ConverterHelper.converterHelper;

// powerbi
import powerbiVisualsApi from "powerbi-visuals-api";
import DataViewCategoryColumn = powerbiVisualsApi.DataViewCategoryColumn;
import DataViewValueColumn = powerbiVisualsApi.DataViewValueColumn;
import DataViewValueColumns = powerbiVisualsApi.DataViewValueColumns;
import DataViewMetadataColumn = powerbiVisualsApi.DataViewMetadataColumn;
import DataViewTable = powerbiVisualsApi.DataViewTable;
import DataView = powerbiVisualsApi.DataView;
import DataViewCategorical = powerbiVisualsApi.DataViewCategorical;
import DataViewCategoricalColumn = powerbiVisualsApi.DataViewCategoricalColumn;
import PrimitiveValue = powerbiVisualsApi.PrimitiveValue;

import {
    toArray as lodashToArray,
    mapValues as lodashMapValues,
    isEmpty as lodashIsEmpty
} from "lodash";

export type ChordChartCategoricalColumns = DataViewCategoryColumn & DataViewValueColumn[] & DataViewValueColumns;

export class ChordChartColumns<T> {
    public static GET_COLUMN_SOURCES(dataView: DataView): ChordChartColumns<DataViewMetadataColumn> {
        return this.getColumnSourcesT<DataViewMetadataColumn>(dataView);
    }

    public static GET_TABLE_VALUES(dataView: DataView): ChordChartColumns<any> {
        let table: DataViewTable = dataView && dataView.table;
        let columns: ChordChartColumns<any> = this.getColumnSourcesT<any[]>(dataView);
        return columns && table && lodashMapValues(
            columns, (n: DataViewMetadataColumn, i) => n && table.rows.map(row => row[n.index]));
    }

    public static GET_TABLE_ROWS(dataView: DataView): ChordChartColumns<any>[] {
        let table: DataViewTable = dataView && dataView.table;
        let columns: ChordChartColumns<any> = this.getColumnSourcesT<any[]>(dataView);
        return columns && table && table.rows.map(row =>
            lodashMapValues(columns, (n: DataViewMetadataColumn, i) => n && row[n.index]));
    }

    public static GET_CATEGORICAL_VALUES(dataView: DataView): ChordChartColumns<any> {
        let categorical: DataViewCategorical = dataView && dataView.categorical;
        let categories: (DataViewCategoryColumn | DataViewValueColumn)[] = categorical && categorical.categories || [];
        if (!categorical.values || categorical.values.length === 0) {
            return null;
        }
        let values: DataViewValueColumns = categorical && categorical.values || <DataViewValueColumns>[];
        let series: PrimitiveValue[] = categorical && values.source && this.GET_SERIES_VALUES(dataView);
        return categorical && lodashMapValues(new this<any[]>(), (n, i) =>
            (<(DataViewCategoryColumn | DataViewValueColumn)[]>lodashToArray(categories)).concat(lodashToArray(values))
                .filter(x => x.source.roles && x.source.roles[i]).map(x => x.values)[0]
            || values.source && values.source.roles && values.source.roles[i] && series);
    }

    public static GET_SERIES_VALUES(dataView: DataView) {
        return dataView && dataView.categorical && dataView.categorical.values
            && dataView.categorical.values.map(x => converterHelper.getSeriesName(x.source));
    }

    public static GET_CATEGORICAL_COLUMNS(dataView: DataView): ChordChartColumns<ChordChartCategoricalColumns> {
        let categorical: DataViewCategorical = dataView && dataView.categorical;
        let categories: DataViewCategoricalColumn[] = categorical && categorical.categories || [];
        let values: DataViewValueColumns = categorical && categorical.values || <DataViewValueColumns>[];
        return categorical && lodashMapValues(
            new this<ChordChartCategoricalColumns>(),
            (n, i) => {
                let result: any = categories.filter(x => x.source.roles && x.source.roles[i])[0];
                if (!result) {
                    result = values.source && values.source.roles && values.source.roles[i] && values;
                }
                if (!result) {
                    result = values.filter(x => x.source.roles && x.source.roles[i]);
                    if (lodashIsEmpty(result)) {
                        result = undefined;
                    }
                }

                return result;
            });
    }

    public static GET_GROUPED_VALUE_COLUMNS(dataView: DataView): ChordChartColumns<DataViewValueColumn>[] {
        let categorical = dataView && dataView.categorical;
        let values = categorical && categorical.values;
        let grouped = values && values.grouped();
        return grouped && grouped.map(g => lodashMapValues(
            new this<DataViewValueColumn>(),
            (n, i) => g.values.filter(v => v.source.roles[i])[0]));
    }

    private static getColumnSourcesT<T>(dataView: DataView): ChordChartColumns<any> {
        let columns = dataView && dataView.metadata && dataView.metadata.columns;
        return columns && lodashMapValues(
            new this<any>(),
            (n, i) => columns.filter(x => x.roles && x.roles[i])[0]);
    }

    // Data Roles
    public Category: T = null;
    public Series: T = null;
    public Y: T = null;
}
