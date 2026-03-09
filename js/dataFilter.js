/**
 * Модуль фильтрации данных.
 *
 * Исключает из расписания:
 * - Элективы (ЭЛЕКТИВ ИНФОРМ, ЭЛЕКТИВ ЧЕРЧЕНИЕ)
 * - ПОУ (ПОУ РУССКИЙ ЯЗЫК, ПОУ МАТЕМАТИКА, ПОУ ИНФОРМАТИКА)
 * - ОГЭ (ОГЭ (РУССКИЙ), ОГЭ (МАТЕМАТИКА))
 * - ЕГЭ (ЕГЭ (РУССКИЙ), ЕГЭ (МАТЕМАТИКА))
 * - Платные (ПЛАТНЫЕ №1, ПЛАТНЫЕ №2)
 * - Динамическая пауза
 * - Записи без класса или предмета
 */

const DataFilter = (() => {
    'use strict';

    // Стоп-слова для фильтрации (case-insensitive, ищем вхождение)
    const STOP_WORDS = [
        'электив',
        'поу ',       // "ПОУ " с пробелом чтобы не цеплять другие слова
        'поу\u00a0',  // неразрывный пробел
        'огэ',
        'егэ',
        'платн',
        'динамическая пауза',
        'факультатив',
        'внеурочн',
        'консультац'
    ];

    // Паттерны-префиксы: предмет НАЧИНАЕТСЯ с одного из них
    const STOP_PREFIX_PATTERNS = [
        /^поу\s/i,        // "ПОУ РУССКИЙ ЯЗЫК"
        /^огэ[\s(]/i,     // "ОГЭ (МАТЕМАТИКА)"
        /^егэ[\s(]/i,     // "ЕГЭ (РУССКИЙ)"
        /^платн/i,        // "ПЛАТНЫЕ №1"
        /^электив/i,      // "ЭЛЕКТИВ ИНФОРМ"
        /^динамическ/i,   // "Динамическая пауза"
        /^факультатив/i,
        /^внеурочн/i,
        /^консультац/i,
    ];

    /**
     * Фильтрация массива нормализованных записей.
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
            const exclusion = getExclusionReason(record);
            if (exclusion) {
                excluded.push({
                    ...record,
                    exclusionReason: exclusion.reason,
                    canInclude: exclusion.canInclude,
                    excludedByHighlight: !!exclusion.isHighlighted
                });
                if (exclusion.reason.includes('стоп')) stats.excludedByStopWord++;
                else if (exclusion.reason.includes('класс')) stats.excludedNoClass++;
                else if (exclusion.reason.includes('предмет')) stats.excludedNoSubject++;
                else if (exclusion.reason.includes('цветом')) stats.excludedByHighlight = (stats.excludedByHighlight || 0) + 1;
            } else {
                filtered.push(record);
                stats.passed++;
            }
        }

        return { filtered, excluded, stats };
    }

    /**
     * Проверить, нужно ли исключить запись.
     * Возвращает null (не исключать) или объект { reason, canInclude }
     * canInclude = true если запись можно вернуть через галочку.
     */
    function getExclusionReason(record) {
        if (!record.className) {
            return { reason: 'Отсутствует класс', canInclude: false };
        }

        if (!record.subject) {
            return { reason: 'Отсутствует предмет', canInclude: false };
        }

        // Жёлтые ячейки — факультативы, можно включить через галочку
        if (record.isHighlighted) {
            return { reason: `Выделено цветом (факультатив): "${record.subject}"`, canInclude: true, isHighlighted: true };
        }

        const subject = record.subject.trim();

        // Проверка по префиксным паттернам — можно вернуть через галочку
        for (const pattern of STOP_PREFIX_PATTERNS) {
            if (pattern.test(subject)) {
                return { reason: `Исключено по стоп-слову: "${subject}"`, canInclude: true };
            }
        }

        // Дополнительная проверка: "ОБЗР  ОГЭ (МАТЕМАТИКА)" — содержит ОГЭ внутри
        const subjectLower = subject.toLowerCase().replace(/\s+/g, ' ');
        for (const word of STOP_WORDS) {
            if (subjectLower.includes(word.trim())) {
                return { reason: `Исключено по стоп-слову "${word.trim()}" в: "${subject}"`, canInclude: true };
            }
        }

        return null;
    }

    /**
     * Проверка текста на стоп-слова.
     */
    function containsStopWord(text) {
        if (!text) return false;
        const lower = text.toLowerCase().replace(/\s+/g, ' ');
        for (const word of STOP_WORDS) {
            if (lower.includes(word.trim())) return true;
        }
        for (const pattern of STOP_PREFIX_PATTERNS) {
            if (pattern.test(text.trim())) return true;
        }
        return false;
    }

    return {
        filterRecords,
        getExclusionReason,
        containsStopWord,
        STOP_WORDS
    };
})();
