/**
 * Модуль нормализации данных.
 *
 * Отвечает за:
 * - Нормализацию названий классов
 * - Нормализацию предметов
 * - Нормализацию дней недели
 * - Нормализацию времени (парсинг "8.35 - 9.15", "8:00-8:40" и т.д.)
 * - Нормализацию номеров уроков ("1.", "2." → 1, 2)
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

    const JS_DAY_MAP = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

    /**
     * Таблица каноничных названий предметов.
     * Ключ — регулярное выражение для сопоставления (после базовой нормализации).
     * Значение — каноничное название.
     *
     * Порядок важен: более специфичные паттерны идут первыми.
     */
    const SUBJECT_CANONICAL = [
        // Литературное чтение (все сокращения)
        { pattern: /^литерат(?:ур)?(?:ное)?[\s.]*чтен/i, canonical: 'Литературное чтение' },
        // Иностранный язык
        { pattern: /^иностр(?:ан)?(?:ный)?[\s.]*яз/i, canonical: 'Иностранный язык' },
        // Физическая культура / физкультура
        { pattern: /^физ(?:ическая)?[\s.]*культ|^физкультура$/i, canonical: 'Физическая культура' },
        // Окружающий мир
        { pattern: /^окружающ(?:ий)?[\s.]*мир/i, canonical: 'Окружающий мир' },
        // Вероятность и статистика
        { pattern: /^вероятность[\s]+и[\s]+стат/i, canonical: 'Вероятность и статистика' },
        // Изобразительное искусство
        { pattern: /^изо$/i, canonical: 'ИЗО' },
    ];

    /**
     * Нормализация одной записи расписания.
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

        // Разрешение времени из Excel (поле "time" содержит "8.35 - 9.15")
        const excelTime = parseTimeRange(record.time, record.startTime, record.endTime);
        const gridTime = LessonTimeGrid.getLessonTime(
            normalized.dayOfWeek,
            normalized.lessonNumber,
            normalized.shift
        );

        timeMode = timeMode || 'mixed';

        if (timeMode === 'grid') {
            if (gridTime.found) {
                normalized.startTime = gridTime.startTime;
                normalized.endTime = gridTime.endTime;
                normalized.shift = gridTime.shift;
                normalized.timeSource = 'школьная сетка';
            }
        } else if (timeMode === 'excel') {
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
     * Нормализация класса.
     * Вход уже нормализован парсером ("6в", "10а соц-эк"), просто trim.
     */
    function normalizeClassName(val) {
        if (val === null || val === undefined) return null;
        let str = String(val).trim();
        if (!str) return null;
        return str;
    }

    /**
     * Нормализация предмета:
     * 1. Убрать двойные пробелы, замыкающие точки
     * 2. Свести сокращения к каноничным полным названиям:
     *    "иностран.язык" → "Иностранный язык"
     *    "литерат.чтение" → "Литературное чтение"
     *    "физ.культура" / "физкультура" → "Физическая культура"
     *    "окружающ.мир" → "Окружающий мир"
     *    "вероятность и стат" → "Вероятность и статистика"
     */
    function normalizeSubject(val) {
        if (val === null || val === undefined) return null;
        let str = String(val).trim();
        // Убрать двойные+ пробелы
        str = str.replace(/\s{2,}/g, ' ');
        // Убрать замыкающие точки (но не в сокращениях вроде "стат.")
        str = str.replace(/\.+$/, '');
        str = str.trim();
        if (!str) return null;

        // Попробовать привести к каноничному названию
        for (const entry of SUBJECT_CANONICAL) {
            if (entry.pattern.test(str)) {
                return entry.canonical;
            }
        }

        return str;
    }

    /**
     * Нормализация дня недели.
     */
    function normalizeDayOfWeek(dayVal, dateVal) {
        if (dayVal !== null && dayVal !== undefined) {
            const str = String(dayVal).trim().toLowerCase();
            if (DAY_MAP[str]) return DAY_MAP[str];

            for (const [key, abbrev] of Object.entries(DAY_MAP)) {
                if (str.includes(key)) return abbrev;
            }
        }

        if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
            return JS_DAY_MAP[dateVal.getDay()];
        }

        return null;
    }

    /**
     * Нормализация номера урока.
     * Обрабатывает "1.", "2.", "3" и т.д.
     */
    function normalizeLessonNumber(val) {
        if (val === null || val === undefined) return null;
        const str = String(val).trim();
        const numStr = str.replace(/[^\d]/g, '');
        const num = parseInt(numStr, 10);
        if (isNaN(num) || num < 1 || num > 15) return null;
        return num;
    }

    /**
     * Нормализация смены.
     */
    function normalizeShift(val) {
        if (val === null || val === undefined) return 1;
        const num = parseInt(String(val).trim(), 10);
        if (isNaN(num) || num < 1 || num > 3) return 1;
        return num;
    }

    /**
     * Парсинг диапазона времени из строки.
     * Форматы: "8.35 - 9.15", "8.00-8.40", "10.25 -11.05", "8:35-9:15"
     */
    function parseTimeRange(timeStr, startTime, endTime) {
        const result = { startTime: null, endTime: null, found: false };

        // Попробовать из отдельных полей
        if (startTime && endTime) {
            const s = parseTimeSingle(startTime);
            const e = parseTimeSingle(endTime);
            if (s && e) {
                result.startTime = s;
                result.endTime = e;
                result.found = true;
                return result;
            }
        }

        // Попробовать из объединённого поля
        if (timeStr) {
            const str = String(timeStr).trim();
            // "8.35 - 9.15", "8.00-8.40", "10.25 -11.05", "14.55-15.35"
            const match = str.match(/(\d{1,2}[.:]\d{2})\s*[-–—]\s*(\d{1,2}[.:]\d{2})/);
            if (match) {
                result.startTime = normalizeTimeStr(match[1]);
                result.endTime = normalizeTimeStr(match[2]);
                result.found = true;
                return result;
            }
        }

        return result;
    }

    /**
     * Парсинг одиночного значения времени.
     */
    function parseTimeSingle(val) {
        if (val === null || val === undefined) return null;

        // Excel числовой формат времени (дробь от суток)
        if (typeof val === 'number' && val >= 0 && val < 1) {
            const totalMinutes = Math.round(val * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            return `${hours}:${String(minutes).padStart(2, '0')}`;
        }

        if (val instanceof Date) {
            return `${val.getHours()}:${String(val.getMinutes()).padStart(2, '0')}`;
        }

        const str = String(val).trim();
        const match = str.match(/^(\d{1,2})[.:](\d{2})$/);
        if (match) {
            return normalizeTimeStr(str);
        }

        return null;
    }

    /**
     * "08.35" → "8:35", "8:00" → "8:00"
     */
    function normalizeTimeStr(str) {
        str = str.replace(/\./g, ':');
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
        DAY_MAP
    };
})();
