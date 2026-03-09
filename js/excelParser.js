/**
 * Модуль парсинга Excel-файлов.
 *
 * Реальный формат расписания — сводная таблица (pivot):
 *   Строка 0: Заголовок "РАСПИСАНИЕ УРОКОВ — X КЛАССЫ"
 *   Строка 1: Учителя / Кл. руководители
 *   Строка 2: "День | № | Время | Класс1 | Класс2 | ..."
 *   Строки 3+: Данные с merged cells для дней недели в колонке A
 *
 * Парсер устойчив к:
 * - Разному регистру заголовков
 * - Разным обозначениям классов (1«А», 1 "А", 1А, 1 а)
 * - Сокращённым/полным дням недели (Пн, ПН, Понедельник)
 * - Разному порядку/именованию служебных колонок
 * - Дополнительным колонкам и строкам
 * - Разному количеству классов на листе
 */

const ExcelParser = (() => {
    'use strict';

    // Паттерны для определения дней недели (полные и сокращённые)
    const DAY_PATTERNS = {
        'Пн': /^пн\.?$|^понедельник$/i,
        'Вт': /^вт\.?$|^вторник$/i,
        'Ср': /^ср\.?$|^среда$/i,
        'Чт': /^чт\.?$|^четверг$/i,
        'Пт': /^пт\.?$|^пятница$/i,
        'Сб': /^сб\.?$|^суббота$/i,
        'Вс': /^вс\.?$|^воскресенье$/i
    };

    // Паттерн для распознавания класса в заголовке:
    //   "1 «А»", "5«Б»", "10 А", "10«А» СОЦ-ЭК", "11 «А» ТЕХНОЛ",
    //   "1 \"А\"", "1'А'", "1А", "1 а", "7 Г"
    const CLASS_HEADER_PATTERN = /^(\d{1,2})\s*[«""'(]?\s*([А-Яа-яЁё])\s*[»""')\.]?\s*(.*)?$/;

    // Паттерн для заголовков служебных колонок
    const SERVICE_HEADER_PATTERNS = {
        day: /^день(\s+недели)?$|^дни?$/i,
        num: /^№\s*(урока)?$|^номер(\s+урока)?$|^урок$|^п[\/.]?п\.?$/i,
        time: /^время(\s+урока)?$|^время\s+начала$/i
    };

    // Стоп-слова для строки заголовков — если ячейка содержит это, это НЕ класс
    const HEADER_STOP_WORDS = /^учитель|^кл[\.\s]*руковод|^преподават|^расписание|^предмет|^кабинет|^каб\.?$|^примечан|^смена$/i;

    /**
     * Парсинг Excel-файла.
     */
    function parseExcelFile(data) {
        const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellStyles: true,
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

        // Ищем лист "предметы" — в нём могут быть предметы с цветной заливкой
        const highlightedFromSubjectsSheet = parseSubjectsSheet(workbook);
        if (highlightedFromSubjectsSheet.size > 0) {
            diagnostics.warnings.push(
                `Лист «предметы»: найдено ${highlightedFromSubjectsSheet.size} предмет(ов) с заливкой: ` +
                Array.from(highlightedFromSubjectsSheet).join(', ')
            );
        }

        for (const sheetName of workbook.SheetNames) {
            // Пропускаем служебный лист "предметы"
            if (sheetName.toLowerCase().replace(/\s/g, '') === 'предметы') continue;

            const sheet = workbook.Sheets[sheetName];
            if (!sheet || !sheet['!ref']) {
                diagnostics.warnings.push(`Лист "${sheetName}" пуст или не имеет данных.`);
                continue;
            }

            const result = parseSheet(sheet, sheetName, diagnostics, highlightedFromSubjectsSheet);
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
     * Парсинг листа «предметы» — ищет ячейки с цветной заливкой.
     * Возвращает Set с названиями предметов, у которых есть заливка.
     */
    function parseSubjectsSheet(workbook) {
        const result = new Set();
        const sheetIdx = workbook.SheetNames.findIndex(
            n => n.toLowerCase().replace(/\s/g, '') === 'предметы'
        );
        if (sheetIdx < 0) return result;

        const sheet = workbook.Sheets[workbook.SheetNames[sheetIdx]];
        if (!sheet || !sheet['!ref']) return result;

        const range = XLSX.utils.decode_range(sheet['!ref']);
        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const ref = XLSX.utils.encode_cell({ r, c });
                const cell = sheet[ref];
                if (!cell) continue;
                const val = cell.w || (cell.v !== undefined ? String(cell.v) : '');
                if (!val || !val.trim()) continue;
                if (hasCellBackground(cell)) {
                    result.add(val.trim());
                }
            }
        }
        return result;
    }

    /**
     * Парсинг одного листа.
     * @param {Set} highlightedSubjects — предметы с заливкой из листа «предметы»
     */
    function parseSheet(sheet, sheetName, diagnostics, highlightedSubjects) {
        const merges = sheet['!merges'] || [];
        const range = XLSX.utils.decode_range(sheet['!ref']);

        // Получить все строки как массив массивов + карта подсветки
        const rows = [];
        const highlighted = {}; // 'r,c' → true если ячейка жёлтая
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
                    if (hasCellBackground(cell)) {
                        highlighted[`${r},${c}`] = true;
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
                `Лист "${sheetName}": не найдена строка с классами в заголовке. ` +
                `Ожидается формат: "День | № | Время | 1«А» | 1«Б» | ..."`
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
        let lastDay = null; // для строк где день пустой (не merged, а просто пустая ячейка)

        for (let r = headerRow + 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row) continue;

            // День недели
            const dayRaw = cellStr(row[dayCol]);
            const dayOfWeek = detectDayOfWeek(dayRaw);

            if (dayOfWeek) {
                lastDay = dayRaw;
            }

            // Используем текущий день или наследуем от предыдущей строки
            const effectiveDay = dayRaw || lastDay;

            // Номер урока
            const numRaw = cellStr(row[numCol]);

            // Время
            const timeRaw = cellStr(row[timeCol]);

            // Если нет номера урока — пропустить (это пустая или служебная строка)
            if (!numRaw) {
                diagnostics.skippedRows++;
                continue;
            }

            // Проверка: номер урока должен содержать хотя бы цифру
            const numDigits = numRaw.replace(/[^\d]/g, '');
            if (!numDigits) {
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
                    dayOfWeek: effectiveDay,
                    date: null,
                    lessonNumber: numRaw,
                    shift: null,
                    startTime: null,
                    endTime: null,
                    time: timeRaw,
                    isHighlighted: !!highlighted[`${r},${classInfo.col}`] ||
                        (highlightedSubjects && highlightedSubjects.has(subject)),
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
     * Стратегия:
     * 1. Сканируем первые 15 строк
     * 2. Для каждой строки считаем сколько ячеек похожи на класс
     * 3. Выбираем строку с максимальным количеством классов (минимум 1)
     * 4. В этой же строке ищем служебные колонки (День, №, Время)
     */
    function findClassHeaderRow(rows) {
        const maxScan = Math.min(rows.length, 15);

        let bestResult = null;
        let bestClassCount = 0;

        for (let i = 0; i < maxScan; i++) {
            const row = rows[i];
            if (!row) continue;

            let dayCol = -1;
            let numCol = -1;
            let timeCol = -1;
            const classColumns = [];

            for (let j = 0; j < row.length; j++) {
                const val = cellStr(row[j]);
                if (!val) continue;
                const trimmed = val.trim();

                // Пропускаем стоп-слова (учителя, расписание и т.д.)
                if (HEADER_STOP_WORDS.test(trimmed)) continue;

                // Проверяем служебные колонки
                if (SERVICE_HEADER_PATTERNS.day.test(trimmed)) {
                    dayCol = j;
                    continue;
                }
                if (SERVICE_HEADER_PATTERNS.num.test(trimmed)) {
                    numCol = j;
                    continue;
                }
                if (SERVICE_HEADER_PATTERNS.time.test(trimmed)) {
                    timeCol = j;
                    continue;
                }

                // Проверяем, это может быть класс
                const classMatch = trimmed.match(CLASS_HEADER_PATTERN);
                if (classMatch) {
                    const grade = classMatch[1];
                    const letter = classMatch[2];
                    const suffix = classMatch[3] ? classMatch[3].trim() : '';

                    // Защита от ложных срабатываний: номер > 12 — вряд ли класс
                    const gradeNum = parseInt(grade, 10);
                    if (gradeNum < 1 || gradeNum > 12) continue;

                    let name;
                    if (suffix) {
                        name = `${grade}${letter.toLowerCase()} ${suffix}`;
                    } else {
                        name = `${grade}${letter.toLowerCase()}`;
                    }
                    classColumns.push({ col: j, name, rawHeader: trimmed });
                }
            }

            // Строка-кандидат: должна содержать хотя бы 1 класс
            // и хотя бы одну служебную колонку ИЛИ больше 1 класса
            const hasService = dayCol >= 0 || numCol >= 0 || timeCol >= 0;
            if (classColumns.length > bestClassCount && (hasService || classColumns.length >= 2)) {
                bestClassCount = classColumns.length;
                bestResult = { headerRow: i, dayCol, numCol, timeCol, classColumns };
            }
        }

        if (!bestResult) return null;

        // Если служебные колонки не найдены — угадываем по позиции
        // Типичный порядок: День(0) | №(1) | Время(2) | Классы(3+)
        const firstClassCol = bestResult.classColumns[0].col;

        if (bestResult.dayCol < 0) {
            // День — обычно первая колонка (до первого класса)
            bestResult.dayCol = Math.max(0, firstClassCol - 3);
        }
        if (bestResult.numCol < 0) {
            bestResult.numCol = Math.max(0, firstClassCol - 2);
        }
        if (bestResult.timeCol < 0) {
            bestResult.timeCol = Math.max(0, firstClassCol - 1);
        }

        return bestResult;
    }

    /**
     * Определить день недели из строки.
     * Поддерживает: ПОНЕДЕЛЬНИК, Понедельник, понедельник, Пн, ПН, пн, Пн.
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
     * Определить, является ли ячейка жёлтой (выделенной цветом).
     * Проверяет фоновый цвет заливки: R >= 200, G >= 200, B <= 100.
     */
    /**
     * Проверяет, есть ли у ячейки фоновая заливка любого цвета (не белая / не прозрачная).
     * Поддерживает два формата стилей XLSX:
     *   - вложенный: cell.s.fill.fgColor  (или cell.s.patternFill.fgColor)
     *   - плоский:   cell.s.fgColor  +  cell.s.patternType
     */
    function hasCellBackground(cell) {
        if (!cell || !cell.s) return false;
        const s = cell.s;

        // Определяем patternType и fgColor из обоих форматов
        let patternType, fgColor;

        const fill = s.fill || s.patternFill;
        if (fill && fill.fgColor) {
            // Вложенный формат
            patternType = fill.patternType;
            fgColor = fill.fgColor;
        } else if (s.fgColor) {
            // Плоский формат (SheetJS иногда кладёт прямо на s)
            patternType = s.patternType;
            fgColor = s.fgColor;
        }

        // patternType "none" означает отсутствие заливки
        if (patternType === 'none') return false;
        if (!fgColor) return false;

        // Если тема задана без rgb — считаем что заливка есть
        if (fgColor.theme !== undefined && !fgColor.rgb) return true;

        let rgb = fgColor.rgb;
        if (!rgb) return false;

        // Убрать альфа-префикс: "FFFFFF00" → "FFFF00"
        if (rgb.length === 8) rgb = rgb.substring(2);
        if (rgb.length !== 6) return false;

        const red = parseInt(rgb.substring(0, 2), 16);
        const green = parseInt(rgb.substring(2, 4), 16);
        const blue = parseInt(rgb.substring(4, 6), 16);

        // Белый или почти белый — не считаем заливкой
        if (red >= 250 && green >= 250 && blue >= 250) return false;

        // Чёрный (000000) тоже пропускаем — это часто дефолт шрифта, не заливка
        if (red === 0 && green === 0 && blue === 0) return false;

        return true;
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
