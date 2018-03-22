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

/// <reference path="_references.ts"/>

module powerbi.extensibility.visual.test {
    // powerbi.extensibility.utils.test
    import VisualBuilderBase = powerbi.extensibility.utils.test.VisualBuilderBase;
    import MockISelectionManager = powerbi.extensibility.utils.test.mocks.MockISelectionManager;

    // ChordChart1444757060245
    import VisualClass = powerbi.extensibility.visual.ChordChart1444757060245.ChordChart;

    export class ChordChartBuilder extends VisualBuilderBase<VisualClass> {
        public selectionManager: SelectionManagerWithBookmarks;

        constructor(width: number, height: number) {
            super(width, height, "ChordChart1444757060245");
        }

        protected build(options: VisualConstructorOptions): VisualClass {
            options.host.createSelectionManager = () => {
                this.selectionManager = new SelectionManagerWithBookmarks();
                return this.selectionManager;
            };

            return new VisualClass(options);
        }

        public get instance(): VisualClass {
            return this.visual;
        }

        public get mainElement(): JQuery {
            return this.element
                .children("svg.chordChart")
                .children("g");
        }

        public get dataLabels(): JQuery {
            return this.mainElement
                .children("g.labels")
                .children("text.data-labels");
        }

        public get sliceTicks(): JQuery {
            return this.mainElement
                .children("g.ticks")
                .children("g.slice-ticks");
        }

        public get chords(): JQuery {
            return this.mainElement
                .children("g.chords")
                .children("path.chord");
        }

        public get slices(): JQuery {
            return this.mainElement
                .children("g.slices")
                .children("path.slice");
        }
    }

    export class SelectionManagerWithBookmarks extends MockISelectionManager {
        private selectionCallback: (ids: ISelectionId[]) => void;
        private selectedSelectionIds: ISelectionId[] = [];

        public registerOnSelectCallback(callback: (ids: ISelectionId[]) => void): void {
            this.selectionCallback = callback;
        }

        public sendSelectionToCallback(selectionIds: ISelectionId[]): void {
            this.selectedSelectionIds = selectionIds;
            this.selectionCallback(selectionIds);
        }

        public getSelectionIds(): visuals.ISelectionId[] {
            return this.selectedSelectionIds as visuals.ISelectionId[];
        }
    }
}
