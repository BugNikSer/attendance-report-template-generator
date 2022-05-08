import * as ExcelJS from 'exceljs';
import { blue, green, yellow } from './palette';

export const workdayFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: green },
};
export const weekendFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: yellow },
};
export const freedayFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: blue },
};

export const setBorders = (style: ExcelJS.BorderStyle): Partial<ExcelJS.Borders> => {
    return {
        top: {
            style: style,
        },
        right: {
            style: style,
        },
        bottom: {
            style: style,
        },
        left: {
            style: style,
        },
    };
};
export const defaultBorders: Partial<ExcelJS.Borders> = setBorders('medium');
export const defaultAlignment: Partial<ExcelJS.Alignment> = {
    vertical: 'middle',
    horizontal: 'center',
};

export const setCellStyling = (cell: ExcelJS.Cell, value: string | number) => {
    cell.value = value;
    cell.alignment = defaultAlignment;
    cell.border = defaultBorders;
};