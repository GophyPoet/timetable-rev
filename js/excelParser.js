/**
 * Модуль парсинга Excel-файлов.
 *
 * Реальный формат расписания — сводная таблица (pivot):
 *   Строка 0: Заголовок "РАСПИСАНИЕ УРОКОВ — X КЛАССЫ"
 *   Строка 1: Учителя / Кл. руководители
 *   Строка 2: "День | № | Время | Класс1 | Класс2 | ..."
 *   Строки 3+: Данные с merged cells для дней недели в колонке A
 *
 * Каждый лист = одна параллель (1 классы, 2 классы, ..., 10-11 классы).
 * Классы расположены по колонкам (D, E, F, G...).
 */

const ExcelParser = (() => {
    'use strict';

    // Паттерны для определения дней недели
    const DAY_PATTERNS = {
        'Пн': /^понедельник$/i,
        'Вт': /^вторник$/i,
        'Ср': /^среда$/i,
        'Чт': /^четверг$/i,
        'Пт': /^пятница$/i,
        'Сб': /^суббота$/i,
        'Вс': /^воскресенье$/i
    };

    // Паттерн для распознавания класса в заголовке: "1 «А»", "5 «Б»", "10«А» СОЦ-ЭК"
    const CLASS_HEADER_PATTERN = /(\d{1,2})\s*[«"(]?\s*([А-Яа-яЁё])\s*[»")]?\s*(.*)?/;

    /**
     * Парсинг Excel-файла.
     * @param {ArrayBuffer} data - Содержимое файла
     * @returns {object} - { rawRecords: [], diagnostics: {} }
     */
    function parseExcelFile(data) {
        const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellStyles: false,
            cellNF: true
        });

        const diagnostics = {
            sheetsProcessed: 0,
            sheetsTotal: workbook.SheetNames.length,
            sheetNames: workbook.SheetNames,
            columnsDetected: {},
            warnings: [],
            skippedRows: 0,
            totalRawRows: 0,
            classesPerSheet: {}
        };

        const allRecords = [];

        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet || !sheet['!ref']) {
                diagnostics.warnings.push(`Лист "${sheetName}" пуст или не имеет данных.`);
                continue;
            }

            const result = parseSheet(sheet, sheetName, diagnostics);
            allRecords.push(...result.records);
            diagnostics.sheetsProcessed++;
        }

        diagnostics.totalRawRows = allRecords.length;

        return {
            rawRecords: allRecords,
            diagnostics
        };
    }

    /**
     * Парсинг одного листа.
     */
    function parseSheet(sheet, sheetName, diagnostics) {
        const merges = sheet['!merges'] || [];
        const range = XLSX.utils.decode_range(sheet['!ref']);

        // Получить все строки как массив массивов
        const rows = [];
        for (let r = range.s.r; r <= range.e.r; r++) {
            const row = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellRef = XLSX.utils.encode_cell({ r, c });
                const cell = sheet[cellRef];
                let value = null;
                if (cell) {
                    if (cell.t === 'd' && cell.v instanceof Date) {
                        value = cell.v;
                    } else if (cell.w !== undefined) {
                        value = cell.w;
                    } else if (cell.v !== undefined) {
                        value = cell.v;
                    }
                }
                row.push(value);
            }
            rows.push(row);
        }

        // Заполнить merged cells — распространить значение на все ячейки merge
        for (const merge of merges) {
            const val = rows[merge.s.r] ? rows[merge.s.r][merge.s.c] : null;
            if (val !== null && val !== undefined) {
                for (let r = merge.s.r; r <= merge.e.r; r++) {
                    for (let c = merge.s.c; c <= merge.e.c; c++) {
                        if (r === merge.s.r && c === merge.s.c) continue;
                        if (rows[r]) {
                            rows[r][c] = val;
                        }
                    }
                }
            }
        }

        // Найти строку заголовков с классами
        const headerInfo = findClassHeaderRow(rows);
        if (!headerInfo) {
            diagnostics.warnings.push(
                `Лист "${sheetName}": не найдена строка заголовков с классами (ожидалось: "День | № | Время | Класс1 | Класс2 | ...").`
            );
            return { records: [] };
        }

        const { headerRow, dayCol, numCol, timeCol, classColumns } = headerInfo;
        const classNames = classColumns.map(c => c.name);

        diagnostics.columnsDetected[sheetName] = [
            `День (кол ${dayCol})`,
            `№ урока (кол ${numCol})`,
            `Время (кол ${timeCol})`,
            ...classNames.map((n, i) => `${n} (кол ${classColumns[i].col})`)
        ];
        diagnostics.classesPerSheet[sheetName] = classNames;

        // Извлечение записей
        const records = [];

        for (let r = headerRow + 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row) continue;

            // День недели (колонка A, merged)
            const dayRaw = cellStr(row[dayCol]);
            const dayOfWeek = detectDayOfWeek(dayRaw);

            // Номер урока
            const numRaw = cellStr(row[numCol]);

            // Время
            const timeRaw = cellStr(row[timeCol]);

            // Если нет ни дня, ни номера урока — пропустить строку
            if (!dayOfWeek && !numRaw) {
                diagnostics.skippedRows++;
                continue;
            }

            // Для каждого класса извлечь предмет
            for (const classInfo of classColumns) {
                const subject = cellStr(row[classInfo.col]);
                if (!subject) continue;

                records.push({
                    className: classInfo.name,
                    subject: subject,
                    dayOfWeek: dayRaw,
                    date: null,
                    lessonNumber: numRaw,
                    shift: null,
                    startTime: null,
                    endTime: null,
                    time: timeRaw,
                    sourceSheet: sheetName,
                    rawRow: row.map(c => c !== null && c !== undefined ? String(c) : '').join(' | ')
                });
            }
        }

        return { records };
    }

    /**
     * Найти строку заголовков с классами.
     *
     * Ищем строку где есть "День" (или подобное) + "№" + "Время" +
     * названия классов вида "1 «А»", "5 «Б»", "10«А» СОЦ-ЭК" и т.д.
     */
    function findClassHeaderRow(rows) {
        const maxScan = Math.min(rows.length, 10);

        for (let i = 0; i < maxScan; i++) {
            const row = rows[i];
            if (!row) continue;

            // Ищем колонки "День", "№", "Время"
            let dayCol = -1;
            let numCol = -1;
            let timeCol = -1;
            const classColumns = [];

            for (let j = 0; j < row.length; j++) {
                const val = cellStr(row[j]);
                if (!val) continue;
                const lower = val.toLowerCase();

                if (lower === 'день' || lower === 'день недели') {
                    dayCol = j;
                } else if (lower === '№' || lower === '№ урока' || lower === 'урок') {
                    numCol = j;
                } else if (lower === 'время') {
                    timeCol = j;
                } else {
                    // Проверяем, это может быть класс
                    const classMatch = val.match(CLASS_HEADER_PATTERN);
                    if (classMatch) {
                        const grade = classMatch[1];
                        const letter = classMatch[2];
                        const suffix = classMatch[3] ? classMatch[3].trim() : '';
                        let name;
                        if (suffix) {
                            name = `${grade}${letter.toLowerCase()} ${suffix}`;
                        } else {
                            name = `${grade}${letter.toLowerCase()}`;
                        }
                        classColumns.push({ col: j, name, rawHeader: val });
                    }
                }
            }

            // Должны найти хотя бы "День" или "№" + минимум один класс
            if (classColumns.length > 0 && (dayCol >= 0 || numCol >= 0)) {
                // Если не нашли какую-то из служебных колонок, попробуем угадать
                if (dayCol < 0) dayCol = 0;
                if (numCol < 0) numCol = 1;
                if (timeCol < 0) timeCol = 2;

                return { headerRow: i, dayCol, numCol, timeCol, classColumns };
            }
        }

        return null;
    }

    /**
     * Определить день недели из строки.
     */
    function detectDayOfWeek(val) {
        if (!val) return null;
        const str = val.trim();
        for (const [abbrev, pattern] of Object.entries(DAY_PATTERNS)) {
            if (pattern.test(str)) return abbrev;
        }
        return null;
    }

    /**
     * Получить строковое значение ячейки (или null).
     */
    function cellStr(val) {
        if (val === null || val === undefined) return null;
        const str = String(val).trim();
        return str || null;
    }

    return {
        parseExcelFile,
        detectDayOfWeek,
        DAY_PATTERNS
    };
})();
