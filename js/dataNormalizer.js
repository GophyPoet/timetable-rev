/**
 * Модуль нормализации данных.
 *
 * Отвечает за:
 * - Нормализацию названий классов (6В → 6в)
 * - Нормализацию предметов (убрать лишние пробелы, скобки)
 * - Нормализацию дней недели
 * - Преобразование дат в дни недели
 * - Нормализацию времени (Excel time → HH:MM)
 * - Нормализацию номеров уроков
 * - Разрешение времени через школьную сетку
 */

const DataNormalizer = (() => {
    'use strict';

    const DAY_MAP = {
        'понедельник': 'Пн', 'пн': 'Пн',
        'вторник': 'Вт', 'вт': 'Вт',
        'среда': 'Ср', 'ср': 'Ср',
        'четверг': 'Чт', 'чт': 'Чт',
        'пятница': 'Пт', 'пт': 'Пт',
        'суббота': 'Сб', 'сб': 'Сб',
        'воскресенье': 'Вс', 'вс': 'Вс'
    };

    // JS Date: 0=Вс, 1=Пн, 2=Вт, 3=Ср, 4=Чт, 5=Пт, 6=Сб
    const JS_DAY_MAP = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

    /**
     * Нормализация одной записи расписания.
     * @param {object} record - Сырая запись
     * @param {string} timeMode - Режим времени: 'mixed', 'excel', 'grid'
     * @returns {object} - Нормализованная запись
     */
    function normalizeRecord(record, timeMode) {
        const normalized = {
            className: normalizeClassName(record.className),
            subject: normalizeSubject(record.subject),
            dayOfWeek: normalizeDayOfWeek(record.dayOfWeek, record.date),
            lessonNumber: normalizeLessonNumber(record.lessonNumber),
            shift: normalizeShift(record.shift),
            startTime: null,
            endTime: null,
            timeSource: 'не определено',
            sourceSheet: record.sourceSheet || '',
            rawValue: record.rawRow || ''
        };

        // Разрешение времени
        const excelTime = resolveExcelTime(record.startTime, record.endTime, record.time);
        const gridTime = LessonTimeGrid.getLessonTime(
            normalized.dayOfWeek,
            normalized.lessonNumber,
            normalized.shift
        );

        timeMode = timeMode || 'mixed';

        if (timeMode === 'grid') {
            // Всегда использовать школьную сетку
            if (gridTime.found) {
                normalized.startTime = gridTime.startTime;
                normalized.endTime = gridTime.endTime;
                normalized.shift = gridTime.shift;
                normalized.timeSource = 'школьная сетка';
            }
        } else if (timeMode === 'excel') {
            // Всегда использовать время из Excel
            if (excelTime.found) {
                normalized.startTime = excelTime.startTime;
                normalized.endTime = excelTime.endTime;
                normalized.timeSource = 'Excel';
            }
        } else {
            // Смешанный: Excel если есть, иначе сетка
            if (excelTime.found) {
                normalized.startTime = excelTime.startTime;
                normalized.endTime = excelTime.endTime;
                normalized.timeSource = 'Excel';
            } else if (gridTime.found) {
                normalized.startTime = gridTime.startTime;
                normalized.endTime = gridTime.endTime;
                normalized.shift = gridTime.shift;
                normalized.timeSource = 'школьная сетка';
            }
        }

        return normalized;
    }

    /**
     * Нормализация массива записей.
     */
    function normalizeAll(records, timeMode) {
        return records.map(r => normalizeRecord(r, timeMode));
    }

    /**
     * Нормализация класса: "6 В" → "6в", "10А" → "10а", "6В" → "6в"
     */
    function normalizeClassName(val) {
        if (val === null || val === undefined) return null;
        let str = String(val).trim();
        str = str.replace(/\s+/g, ''); // Убрать все пробелы
        str = str.toLowerCase();
        // Убрать "класс" если прилепилось
        str = str.replace(/^класс\s*/i, '');
        return str || null;
    }

    /**
     * Нормализация предмета:
     * - убрать двойные пробелы
     * - обрезать
     * - убрать служебные хвосты
     */
    function normalizeSubject(val) {
        if (val === null || val === undefined) return null;
        let str = String(val).trim();
        str = str.replace(/\s{2,}/g, ' ');
        // Убрать замыкающие точки, если это не сокращение
        str = str.replace(/\.+$/, '');
        str = str.trim();
        return str || null;
    }

    /**
     * Нормализация дня недели.
     * Если есть дата, вычислить день недели автоматически.
     */
    function normalizeDayOfWeek(dayVal, dateVal) {
        // Сначала попробовать из явного дня
        if (dayVal !== null && dayVal !== undefined) {
            const str = String(dayVal).trim().toLowerCase();
            if (DAY_MAP[str]) return DAY_MAP[str];

            // Попробовать частичное совпадение
            for (const [key, abbrev] of Object.entries(DAY_MAP)) {
                if (str.includes(key) || key.includes(str)) {
                    return abbrev;
                }
            }
        }

        // Попробовать из даты
        if (dateVal !== null && dateVal !== undefined) {
            let date;
            if (dateVal instanceof Date) {
                date = dateVal;
            } else {
                // Попытка распарсить строку даты
                const str = String(dateVal).trim();
                date = new Date(str);

                // Попробовать формат DD.MM.YYYY
                if (isNaN(date.getTime())) {
                    const match = str.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/);
                    if (match) {
                        const day = parseInt(match[1], 10);
                        const month = parseInt(match[2], 10) - 1;
                        let year = parseInt(match[3], 10);
                        if (year < 100) year += 2000;
                        date = new Date(year, month, day);
                    }
                }
            }

            if (date && !isNaN(date.getTime())) {
                return JS_DAY_MAP[date.getDay()];
            }
        }

        return null;
    }

    /**
     * Нормализация номера урока.
     */
    function normalizeLessonNumber(val) {
        if (val === null || val === undefined) return null;
        const str = String(val).trim();
        // Убрать нечисловые символы (кроме цифр)
        const numStr = str.replace(/[^\d]/g, '');
        const num = parseInt(numStr, 10);
        if (isNaN(num) || num < 1 || num > 15) return null;
        return num;
    }

    /**
     * Нормализация смены.
     */
    function normalizeShift(val) {
        if (val === null || val === undefined) return 1; // По умолчанию 1-я смена
        const num = parseInt(String(val).trim(), 10);
        if (isNaN(num) || num < 1 || num > 3) return 1;
        return num;
    }

    /**
     * Разрешение времени из Excel-данных.
     */
    function resolveExcelTime(startTime, endTime, combinedTime) {
        const result = { startTime: null, endTime: null, found: false };

        // Попробовать из отдельных полей
        const start = parseTimeValue(startTime);
        const end = parseTimeValue(endTime);

        if (start && end) {
            result.startTime = start;
            result.endTime = end;
            result.found = true;
            return result;
        }

        // Попробовать из объединённого поля "8:00-8:40" или "8:00 - 8:40"
        if (combinedTime !== null && combinedTime !== undefined) {
            const str = String(combinedTime).trim();
            const match = str.match(/(\d{1,2}[:.]\d{2})\s*[-–—]\s*(\d{1,2}[:.]\d{2})/);
            if (match) {
                result.startTime = normalizeTimeString(match[1]);
                result.endTime = normalizeTimeString(match[2]);
                result.found = true;
                return result;
            }
        }

        return result;
    }

    /**
     * Парсинг значения времени из различных форматов.
     */
    function parseTimeValue(val) {
        if (val === null || val === undefined) return null;

        // Excel хранит время как дробь от суток (0.0 - 1.0)
        if (typeof val === 'number' && val >= 0 && val < 1) {
            const totalMinutes = Math.round(val * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            return `${hours}:${String(minutes).padStart(2, '0')}`;
        }

        // Date object
        if (val instanceof Date) {
            const hours = val.getHours();
            const minutes = val.getMinutes();
            return `${hours}:${String(minutes).padStart(2, '0')}`;
        }

        // Строка
        const str = String(val).trim();
        const match = str.match(/^(\d{1,2})[:.h](\d{2})$/);
        if (match) {
            return normalizeTimeString(str);
        }

        return null;
    }

    /**
     * Нормализация строки времени "08:35" → "8:35"
     */
    function normalizeTimeString(str) {
        str = str.replace(/[.h]/g, ':');
        const match = str.match(/^0?(\d{1,2}):(\d{2})$/);
        if (match) {
            return `${parseInt(match[1], 10)}:${match[2]}`;
        }
        return str;
    }

    return {
        normalizeAll,
        normalizeRecord,
        normalizeClassName,
        normalizeSubject,
        normalizeDayOfWeek,
        normalizeLessonNumber,
        normalizeShift,
        parseTimeValue,
        DAY_MAP,
        JS_DAY_MAP
    };
})();
