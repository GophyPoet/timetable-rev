/**
 * Модуль фильтрации данных.
 *
 * Отвечает за:
 * - Исключение элективов, ПОУ, ОГЭ, платных занятий
 * - Исключение записей без класса или предмета
 * - Подсчёт исключённых записей с причинами
 */

const DataFilter = (() => {
    'use strict';

    // Стоп-слова для фильтрации (case-insensitive)
    const STOP_WORDS = [
        'электив',
        'элективн',
        'поу',
        'огэ',
        'платн',
        'платные',
        'факультатив',
        'консультац',
        'внеурочн',
        'доп.образ',
        'дополнительное образ'
    ];

    // Компилированное регулярное выражение для стоп-слов
    const STOP_REGEX = new RegExp(
        STOP_WORDS.map(w => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'),
        'i'
    );

    /**
     * Фильтрация массива нормализованных записей.
     * @param {Array} records - Нормализованные записи
     * @returns {{ filtered: Array, excluded: Array, stats: object }}
     */
    function filterRecords(records) {
        const filtered = [];
        const excluded = [];
        const stats = {
            total: records.length,
            passed: 0,
            excludedByStopWord: 0,
            excludedNoClass: 0,
            excludedNoSubject: 0,
            excludedNoDay: 0
        };

        for (const record of records) {
            const reason = getExclusionReason(record);
            if (reason) {
                excluded.push({ ...record, exclusionReason: reason });
                if (reason.includes('стоп-слов')) stats.excludedByStopWord++;
                else if (reason.includes('класс')) stats.excludedNoClass++;
                else if (reason.includes('предмет')) stats.excludedNoSubject++;
                else if (reason.includes('день')) stats.excludedNoDay++;
            } else {
                filtered.push(record);
                stats.passed++;
            }
        }

        return { filtered, excluded, stats };
    }

    /**
     * Проверить, нужно ли исключить запись.
     * @returns {string|null} - Причина исключения или null
     */
    function getExclusionReason(record) {
        // Проверка на отсутствие класса
        if (!record.className) {
            return 'Отсутствует класс';
        }

        // Проверка на отсутствие предмета
        if (!record.subject) {
            return 'Отсутствует предмет';
        }

        // Проверка стоп-слов в предмете
        const subjectClean = record.subject.replace(/\s+/g, ' ').trim();
        if (STOP_REGEX.test(subjectClean)) {
            return `Содержит стоп-слово в предмете: "${record.subject}"`;
        }

        // Проверка стоп-слов в сырых данных строки
        if (record.rawValue) {
            const rawClean = record.rawValue.replace(/\s+/g, ' ').trim();
            // Проверяем только если стоп-слово не в названии класса/предмета
            // а в служебных колонках
            if (STOP_REGEX.test(rawClean) && !STOP_REGEX.test(subjectClean)) {
                // Дополнительная проверка — стоп-слово должно быть отдельным
                const rawLower = rawClean.toLowerCase();
                for (const word of STOP_WORDS) {
                    if (rawLower.includes(word)) {
                        // Проверим, не является ли это частью предмета
                        if (!subjectClean.toLowerCase().includes(word)) {
                            return `Содержит стоп-слово в строке: "${word}"`;
                        }
                    }
                }
            }
        }

        return null;
    }

    /**
     * Проверка текста на стоп-слова.
     */
    function containsStopWord(text) {
        if (!text) return false;
        return STOP_REGEX.test(text.replace(/\s+/g, ' ').trim());
    }

    return {
        filterRecords,
        getExclusionReason,
        containsStopWord,
        STOP_WORDS
    };
})();
