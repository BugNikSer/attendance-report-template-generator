import * as fs from 'fs';
import * as ExcelJS from 'exceljs';
import * as path from 'path';
import { months, days } from './constants';
import { IStudents } from './types';
import { workdayFill, weekendFill, freedayFill, setCellStyling, setBorders } from './styling';

const studentsRaw: string = fs.readFileSync('./students.txt', {
    encoding: 'utf8',
});

let curTeacher: string;
let curGroup: string;
const studentsConfig: IStudents = {};

studentsRaw
    .split('\n')
    .filter((str: string) => str.trim() !== '')
    .forEach((str: string) => {
        if (str.startsWith('--')) {
            curTeacher = str.replace('--', '').trim();
        } else if (str.startsWith('-')) {
            curGroup = str.replace('-', '').trim();
        } else {
            if (!studentsConfig[curTeacher]) {
                studentsConfig[curTeacher] = {};
            }

            if (!studentsConfig[curTeacher][curGroup]) {
                studentsConfig[curTeacher][curGroup] = [];
            }

            studentsConfig[curTeacher][curGroup].push(str.trim());
        }
    });

const now = new Date();
const year = now.getFullYear();
const month = now.getMonth() + 1;
const daysNumber = new Date(year, month, 0).getDate();

const workbook = new ExcelJS.Workbook();

Object.entries(studentsConfig).forEach(([teacher, groups]) => {
    const sheet: ExcelJS.Worksheet = workbook.addWorksheet(teacher);
    let usedRows: number = 3;
    const studentsTotalNumber: number = Object.values(groups).reduce(
        (result, students) => {
            result += students.length;
            return result;
        },
        0
    );

    // Заполнение шапки
    sheet.mergeCells(1, 1, 1, daysNumber + 4);
    const teacherCell = sheet.getCell(1, 1);
    setCellStyling(teacherCell, teacher);

    sheet.mergeCells(2, 1, 3, 1);
    const groupHeaderCell = sheet.getCell(2, 1);
    setCellStyling(groupHeaderCell, 'Группа');

    sheet.mergeCells(2, 2, 3, 2);
    const studentHeaderCell = sheet.getCell(2, 2);
    setCellStyling(studentHeaderCell, 'Студент');

    sheet.mergeCells(2, 3, 2, daysNumber + 2);
    const monthCell = sheet.getCell(2, 3);
    setCellStyling(monthCell, months[month] + ' ' + year);

    for (let i: number = 1; i <= daysNumber; i++) {
        sheet.getColumn(2 + i).width = 3.5;
        const dayHeaderCell = sheet.getCell(3, 2 + i);
        setCellStyling(dayHeaderCell, i);
    }

    sheet.mergeCells(2, daysNumber + 3, 3, daysNumber + 4);
    const lessonsHeaderCell = sheet.getCell(2, daysNumber + 3);
    sheet.getColumn(daysNumber + 3).width = 15;
    sheet.getColumn(daysNumber + 4).width = 15;
    setCellStyling(lessonsHeaderCell, 'Занятия');

    Object.entries(groups).forEach(([group, students]) => {
        sheet.mergeCells(usedRows + 1, 1, usedRows + students.length, 1);
        const [groupName, lessons] = group.split('[');
        const lessonsList = lessons
            .substring(0, lessons.length - 1)
            .split(',')
            .map((lesson) => lesson.trim());

        const groupCell = sheet.getCell(usedRows + 1, 1);
        setCellStyling(groupCell, groupName);

        sheet.mergeCells(
            usedRows + 1,
            daysNumber + 3,
            usedRows + students.length,
            daysNumber + 4
        );
        const lessonsCell = sheet.getCell(usedRows + 1, daysNumber + 3);
        setCellStyling(lessonsCell, lessonsList.join('\n'));

        const workingDaysIndexes = days.reduce((result, day, index) => {
            if (lessonsList.some((lesson) => lesson.includes(day))) {
                result.push(index);
            }
            return result;
        }, []);

        let studentsColumnWidth: number = 0;
        students.forEach((studentName) => {
            const studentCell = sheet.getCell(usedRows + 1, 2);
            setCellStyling(studentCell, studentName);

            if (studentName.length > studentsColumnWidth) {
                studentsColumnWidth = studentName.length;
                sheet.getColumn(2).width = studentsColumnWidth;
            }

            for (let i: number = 1; i <= daysNumber; i++) {
                const dayCell = sheet.getCell(usedRows + 1, i + 2);
                dayCell.border = setBorders('thin');
                const day = new Date(year, now.getMonth(), i).getDay();

                if (workingDaysIndexes.includes(day)) {
                    dayCell.fill = workdayFill;
                } else if ([0, 6].includes(day)) {
                    dayCell.fill = weekendFill;
                } else {
                    dayCell.fill = freedayFill;
                }
            }
            usedRows++;
        });

        for (let i = 1; i <= daysNumber; i++) {
            const bottomCell = sheet.getCell(usedRows, i + 2);
            bottomCell.border = {
                ...setBorders('thin'),
                bottom: {
                    style: 'medium',
                }
            }
        }
    });

    const studentsColumn = sheet.getColumn(2);
    studentsColumn.width = studentsColumn.width;

    sheet.getCell(usedRows + 2, 1).fill = workdayFill;
    sheet.getCell(usedRows + 2, 2).value = 'Рабочий день';
    sheet.getCell(usedRows + 3, 1).fill = weekendFill;
    sheet.getCell(usedRows + 3, 2).value = 'Выходной день';
    sheet.getCell(usedRows + 4, 1).fill = freedayFill;
    sheet.getCell(usedRows + 4, 2).value = 'Свободный день';

});

workbook.xlsx.writeFile(path.resolve(__dirname, 'General english ' + months[month] + ' ' + year + '.xlsx'));
