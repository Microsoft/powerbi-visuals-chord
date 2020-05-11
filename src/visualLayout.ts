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

// powerbi
import powerbiVisualsApi from "powerbi-visuals-api";
import IViewport = powerbiVisualsApi.IViewport;

// powerbi.extensibility.utils.svg
import { IMargin } from "powerbi-visuals-utils-svgutils";

export class VisualLayout {
    private marginValue: IMargin;
    private viewportValue: IViewport;
    private viewportInValue: IViewport;
    private minViewportValue: IViewport;

    public defaultMargin: IMargin;
    public defaultViewport: IViewport;

    constructor(defaultViewport?: IViewport, defaultMargin?: IMargin) {
        this.defaultViewport = defaultViewport || { width: 0, height: 0 };
        this.defaultMargin = defaultMargin || { top: 0, bottom: 0, right: 0, left: 0 };
    }

    public get margin(): IMargin {
        return this.marginValue || (this.margin = this.defaultMargin);
    }

    public set margin(value: IMargin) {
        this.marginValue = VisualLayout.restrictToMinMax(value);
        this.update();
    }

    public get viewport(): IViewport {
        return this.viewportValue || (this.viewportValue = this.defaultViewport);
    }

    public set viewport(value: IViewport) {
        this.viewportValue = VisualLayout.restrictToMinMax(value, this.minViewport);
        this.update();
    }

    public get viewportIn(): IViewport {
        return this.viewportInValue || this.viewport;
    }

    public get minViewport(): IViewport {
        return this.minViewportValue;
    }

    public set minViewport(value: IViewport) {
        this.minViewportValue = value;
    }

    public get viewportInIsZero(): boolean {
        return this.viewportIn.width === 0 || this.viewportIn.height === 0;
    }

    public resetMargin(): void {
        this.margin = this.defaultMargin;
    }

    private update(): void {
        this.viewportInValue = VisualLayout.restrictToMinMax({
            width: this.viewport.width - (this.margin.left + this.margin.right),
            height: this.viewport.height - (this.margin.top + this.margin.bottom)
        }, this.minViewportValue);
    }

    private static restrictToMinMax<T extends Object>(value: T, minValue?: T): T {
        let result: T = <T>{};
        Object.keys(value).forEach(x => result[x] = Math.max(minValue && minValue[x] || 0, value[x]));
        return result;
    }
}
