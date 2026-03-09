/**
 * Модуль парсинга Excel-файлов.
 *
 * Отвечает за:
 * - Чтение файла через SheetJS
 * - Обработку merged cells
 * - Поиск строки заголовков (header detection)
 * - Определение колонок по эвристикам
 * - Извлечение сырых записей расписания
 */

const ExcelParser = (() => {
    'use strict';

    // Паттерны для распознавания заголовков колонок
    const COLUMN_PATTERNS = {
        className: {
            keywords: ['класс', 'класcы', 'group', 'класcы', 'кл'],
            test: (val) => /класс|кл\.?$|group/i.test(val)
        },
        subject: {
            keywords: ['предмет', 'дисциплина', 'урок', 'занятие', 'subject'],
            test: (val) => /предмет|дисциплин|subject|назван.*урок|назван.*занят/i.test(val)
        },
        dayOfWeek: {
            keywords: ['день', 'день недели', 'day'],
            test: (val) => /день\s*(недели)?|day/i.test(val)
        },
        date: {
            keywords: ['дата', 'date'],
            test: (val) => /^дата$|^date$/i.test(val)
        },
        lessonNumber: {
            keywords: ['урок', '№ урока', 'номер урока', 'lesson', '№', 'пара'],
            test: (val) => /№\s*урок|номер\s*урок|^урок$|^№$|^пара$|lesson\s*#?n/i.test(val)
        },
        shift: {
            keywords: ['смена', 'shift'],
            test: (val) => /смена|shift/i.test(val)
        },
        startTime: {
            keywords: ['начало', 'время начала', 'start', 'с ', 'нач'],
            test: (val) => /начал|время\s*начал|^start|^с\s|^нач\.?$/i.test(val)
        },
        endTime: {
            keywords: ['конец', 'окончание', 'время окончания', 'end', 'до', 'оконч'],
            test: (val) => /конец|окончан|время\s*оконч|^end|^до\s|^оконч\.?$/i.test(val)
        },
        time: {
            keywords: ['время', 'time'],
            test: (val) => /^время$|^time$/i.test(val)
        }
    };

    // Паттерны для определения дней недели в данных
    const DAY_PATTERNS = {
        'Пн': /^пн$|^понедельник$/i,
        'Вт': /^вт$|^вторник$/i,
        'Ср': /^ср$|^среда$/i,
        'Чт': /^чт$|^четверг$/i,
        'Пт': /^пт$|^пятница$/i,
        'Сб': /^сб$|^суббота$/i,
        'Вс': /^вс$|^воскресенье$/i
    };

    // Паттерн для класса (например, "6В", "10А", "11б")
    const CLASS_PATTERN = /^(\d{1,2})\s*([а-яА-ЯёЁa-zA-Z])?$/;

    /**
     * Парсинг Excel-файла.
     * @param {ArrayBuffer} data - Содержимое файла
     * @returns {object} - { sheets: [], rawRecords: [], diagnostics: {} }
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
            totalRawRows: 0
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
        // Развернуть merged cells
        const merges = sheet['!merges'] || [];

        // Получить все строки как массив массивов
        const range = XLSX.utils.decode_range(sheet['!ref']);
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

        // Заполнить merged cells
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

        // Поиск строки заголовков
        const headerInfo = detectHeaders(rows, diagnostics, sheetName);

        if (!headerInfo) {
            // Попытка альтернативного парсинга — структура "день/класс" (сводная таблица)
            const pivotResult = tryParsePivotSchedule(rows, sheetName, diagnostics);
            if (pivotResult.length > 0) {
                return { records: pivotResult };
            }

            diagnostics.warnings.push(
                `Лист "${sheetName}": не удалось определить заголовки колонок. ` +
                `Попробуйте файл с более стандартной структурой.`
            );
            return { records: [] };
        }

        diagnostics.columnsDetected[sheetName] = headerInfo.detectedColumns;

        // Извлечение записей из строк данных
        const records = extractRecords(rows, headerInfo, sheetName, diagnostics);
        return { records };
    }

    /**
     * Поиск строки заголовков в первых 20 строках.
     */
    function detectHeaders(rows, diagnostics, sheetName) {
        const maxScan = Math.min(rows.length, 20);

        let bestMatch = null;
        let bestScore = 0;

        for (let i = 0; i < maxScan; i++) {
            const row = rows[i];
            if (!row || row.every(c => c === null || c === undefined || String(c).trim() === '')) continue;

            const mapping = {};
            const detected = [];
            let score = 0;

            for (let j = 0; j < row.length; j++) {
                const val = row[j];
                if (val === null || val === undefined) continue;
                const str = String(val).trim().toLowerCase();
                if (str === '') continue;

                for (const [field, pattern] of Object.entries(COLUMN_PATTERNS)) {
                    if (mapping[field] !== undefined) continue;
                    if (pattern.test(str)) {
                        mapping[field] = j;
                        detected.push(field);
                        score++;
                        break;
                    }
                }
            }

            // Минимум нужен класс или предмет + хотя бы одно еще поле
            if (score > bestScore && score >= 2) {
                bestScore = score;
                bestMatch = {
                    headerRow: i,
                    mapping,
                    detectedColumns: detected
                };
            }
        }

        return bestMatch;
    }

    /**
     * Извлечение записей по маппингу колонок.
     */
    function extractRecords(rows, headerInfo, sheetName, diagnostics) {
        const records = [];
        const { headerRow, mapping } = headerInfo;
        let currentDay = null; // Для наследования дня недели сверху

        for (let i = headerRow + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.every(c => c === null || c === undefined || String(c).trim() === '')) {
                diagnostics.skippedRows++;
                continue;
            }

            const record = {
                className: getCellValue(row, mapping.className),
                subject: getCellValue(row, mapping.subject),
                dayOfWeek: getCellValue(row, mapping.dayOfWeek),
                date: getCellValue(row, mapping.date),
                lessonNumber: getCellValue(row, mapping.lessonNumber),
                shift: getCellValue(row, mapping.shift),
                startTime: getCellValue(row, mapping.startTime),
                endTime: getCellValue(row, mapping.endTime),
                time: getCellValue(row, mapping.time),
                sourceSheet: sheetName,
                rawRow: row.map(c => c !== null && c !== undefined ? String(c) : '').join(' | ')
            };

            // Наследование дня недели от предыдущей непустой строки
            if (record.dayOfWeek) {
                currentDay = record.dayOfWeek;
            } else if (currentDay) {
                record.dayOfWeek = currentDay;
            }

            // Пропустить полностью пустые записи
            if (!record.className && !record.subject) {
                diagnostics.skippedRows++;
                continue;
            }

            records.push(record);
        }

        return records;
    }

    /**
     * Попытка распознать сводное расписание формата:
     * Строки = дни недели, столбцы = номера уроков,
     * с блоками для каждого класса.
     */
    function tryParsePivotSchedule(rows, sheetName, diagnostics) {
        const records = [];
        let currentClass = null;
        let lessonColumns = {};

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;

            // Проверяем, не является ли строка заголовком класса
            const firstCell = row[0] !== null && row[0] !== undefined ? String(row[0]).trim() : '';

            // Проверяем, содержит ли строка номера уроков (1, 2, 3...)
            const numericCells = row.filter(c => {
                const n = parseInt(c, 10);
                return !isNaN(n) && n >= 1 && n <= 8;
            });

            if (numericCells.length >= 3) {
                // Это может быть строка с номерами уроков
                lessonColumns = {};
                for (let j = 0; j < row.length; j++) {
                    const n = parseInt(row[j], 10);
                    if (!isNaN(n) && n >= 1 && n <= 8) {
                        lessonColumns[j] = n;
                    }
                }
                continue;
            }

            // Проверяем, не это ли класс
            if (CLASS_PATTERN.test(firstCell) || /^класс\s*/i.test(firstCell)) {
                currentClass = firstCell.replace(/^класс\s*/i, '').trim();
                continue;
            }

            // Проверяем, является ли первая ячейка днём недели
            const detectedDay = detectDayOfWeek(firstCell);
            if (detectedDay && currentClass && Object.keys(lessonColumns).length > 0) {
                for (const [colIdx, lessonNum] of Object.entries(lessonColumns)) {
                    const subject = getCellValue(row, parseInt(colIdx, 10));
                    if (subject && String(subject).trim() !== '') {
                        records.push({
                            className: currentClass,
                            subject: String(subject).trim(),
                            dayOfWeek: detectedDay,
                            date: null,
                            lessonNumber: String(lessonNum),
                            shift: null,
                            startTime: null,
                            endTime: null,
                            time: null,
                            sourceSheet: sheetName,
                            rawRow: row.map(c => c !== null && c !== undefined ? String(c) : '').join(' | ')
                        });
                    }
                }
            }
        }

        return records;
    }

    /**
     * Определить день недели из строки.
     */
    function detectDayOfWeek(val) {
        if (!val) return null;
        const str = String(val).trim();
        for (const [abbrev, pattern] of Object.entries(DAY_PATTERNS)) {
            if (pattern.test(str)) return abbrev;
        }
        return null;
    }

    /**
     * Безопасное чтение ячейки.
     */
    function getCellValue(row, colIndex) {
        if (colIndex === undefined || colIndex === null) return null;
        if (colIndex >= row.length) return null;
        const val = row[colIndex];
        if (val === null || val === undefined) return null;
        return val;
    }

    return {
        parseExcelFile,
        detectDayOfWeek,
        CLASS_PATTERN,
        DAY_PATTERNS
    };
})();
