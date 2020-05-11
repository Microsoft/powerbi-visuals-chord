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

// powerbi.extensibility.utils.svg
import { IRect } from "powerbi-visuals-utils-svgutils";

import { range as lodashRange } from "lodash";

// powerbi.extensibility.utils.test
import { RgbColor, parseColorString, getRandomNumber } from "powerbi-visuals-utils-testutils";

export type IsInRangeFunction = (value: number, min: number, max: number) => boolean;
export type CompareValuesFunction = (value1: number, value2: number) => boolean;

const MinColorValue: number = 0;
const MaxColorValue: number = 16777216;

export function getSolidColorStructuralObject(color: string): any {
    return { solid: { color } };
}

export function areColorsEqual(firstColor: string, secondColor: string): boolean {
    const firstConvertedColor: RgbColor = parseColorString(firstColor),
        secondConvertedColor: RgbColor = parseColorString(secondColor);

    return firstConvertedColor.R === secondConvertedColor.R
        && firstConvertedColor.G === secondConvertedColor.G
        && firstConvertedColor.B === secondConvertedColor.B;
}

export function getRandomUniqueIntegers(
    count: number,
    min: number = 0,
    max: number): number[] {

    const result: number[] = [];

    for (let i: number = 0; i < count; i++) {
        result.push(getRandomNumber(min, max, result, Math.floor));
    }

    return result;
}

export function getRandomUniqueHexColors(count: number): string[] {
    return getRandomUniqueIntegers(
        count,
        MinColorValue,
        MaxColorValue).map(getHexColorFromNumber);
}

export function getHexColorFromNumber(value: number): string {
    const hex: string = value.toString(16).toUpperCase(),
        color: string = hex.length === 6
            ? hex
            : `${lodashRange(0, 6 - hex.length, 0).join("")}${hex}`;

    return `#${color}`;
}

export function isTextElementInOrOutElement(
    mainElement: Element,
    textElement: Element,
    compareValues: CompareValuesFunction): boolean {

    return isRectangleInOrOutRectangle(
        mainElement.getBoundingClientRect(),
        textElement.getBoundingClientRect(),
        compareValues);
}

export function isRectangleInOrOutRectangle(
    mainRect: IRect,
    rect: IRect,
    compareValues: CompareValuesFunction): boolean {

    return compareValues(rect.left, mainRect.left)
        && compareValues(rect.top, mainRect.top)
        && compareValues(mainRect.left + mainRect.width, rect.left + rect.width)
        && compareValues(mainRect.top + mainRect.height, rect.top + rect.height);
}


export function getTextElementRects(textElement: Element): IRect {
    const clientRect: ClientRect = textElement.getBoundingClientRect(),
        fontSizeString: string = window.getComputedStyle(textElement).fontSize,
        fontSize: number = parseFloat(fontSizeString);

    return <IRect>{
        left: clientRect.left,
        top: clientRect.bottom - fontSize,
        height: fontSize,
        width: clientRect.width
    };
}

export function isSomeTextElementOverlapped(
    textElements: Element[],
    isInRange: IsInRangeFunction): boolean {

    return isSomeRectangleOverlapped(textElements.map(getTextElementRects), isInRange);
}

export function isSomeRectangleOverlapped(
    rects: IRect[],
    isInRange: IsInRangeFunction): boolean {

    return rects.some((firstRect: IRect, firstIndex: number) => {
        return rects.some((secondRect: IRect, secondIndex: number) => {
            return firstIndex !== secondIndex
                && isRectangleOverlapped(firstRect, secondRect, isInRange);
        });
    });
}

export function isRectangleOverlapped(
    firstRect: IRect,
    secondRect: IRect,
    isInRange: IsInRangeFunction): boolean {

    const xOverlap: boolean = isInRange(firstRect.left, secondRect.left, secondRect.left + secondRect.width)
        || isInRange(secondRect.left, firstRect.left, firstRect.left + firstRect.width);

    const yOverlap: boolean = isInRange(firstRect.top, secondRect.top, secondRect.top + secondRect.height)
        || isInRange(secondRect.top, firstRect.top, firstRect.top + firstRect.height);

    return xOverlap && yOverlap;
}
