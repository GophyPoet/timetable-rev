/**
 * Модуль отрисовки UI.
 *
 * Отвечает за:
 * - Рендер таблицы школьной сетки времени
 * - Рендер результирующей таблицы расписания
 * - Рендер диагностики
 * - Рендер debug-информации
 * - Экспорт в CSV и JSON
 */

const UIRenderer = (() => {
    'use strict';

    const DAY_ORDER = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];

    /**
     * Отрисовать сетки времени уроков.
     */
    function renderTimeGrids() {
        renderGridTable('mondayGrid', LessonTimeGrid.getMondaySchedule());
        renderGridTable('tueSatGrid', LessonTimeGrid.getTueToSatSchedule());
    }

    function renderGridTable(elementId, schedule) {
        const table = document.getElementById(elementId);
        if (!table) return;

        let html = `
            <thead>
                <tr>
                    <th>Смена</th>
                    <th>№ урока</th>
                    <th>Время начала</th>
                    <th>Время окончания</th>
                </tr>
            </thead>
            <tbody>
        `;

        for (const item of schedule) {
            html += `
                <tr>
                    <td>${item.shift}</td>
                    <td>${item.lessonNumber}</td>
                    <td>${item.startTime}</td>
                    <td>${item.endTime}</td>
                </tr>
            `;
        }

        html += '</tbody>';
        table.innerHTML = html;
    }

    /**
     * Обновить статус анализа.
     */
    function renderStatus(stats) {
        const section = document.getElementById('statusSection');
        section.hidden = false;

        document.getElementById('statusRecords').textContent = stats.totalRecords || 0;
        document.getElementById('statusClasses').textContent = stats.totalClasses || 0;
        document.getElementById('statusSubjects').textContent = stats.totalSubjects || 0;
        document.getElementById('statusExcluded').textContent = stats.totalExcluded || 0;
        document.getElementById('statusSheets').textContent = stats.sheetsProcessed || 0;
    }

    /**
     * Заполнить выпадающий список классов.
     */
    function populateClassSelect(classes) {
        const select = document.getElementById('classSelect');
        select.innerHTML = '<option value="">— Выберите класс —</option>';

        const sorted = sortClasses(classes);
        for (const cls of sorted) {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = cls;
            select.appendChild(option);
        }

        select.disabled = false;
    }

    /**
     * Заполнить выпадающий список предметов.
     * @param {string[]} subjects — основные предметы
     * @param {string[]} [extraSubjects] — доп. предметы (из исключённых, включены галочкой)
     */
    function populateSubjectSelect(subjects, extraSubjects) {
        const select = document.getElementById('subjectSelect');
        select.innerHTML = '<option value="">— Выберите предмет —</option>';

        const sorted = subjects.slice().sort((a, b) => a.localeCompare(b, 'ru'));
        for (const subj of sorted) {
            const option = document.createElement('option');
            option.value = subj;
            option.textContent = subj;
            select.appendChild(option);
        }

        // Добавить доп. предметы в отдельной группе
        if (extraSubjects && extraSubjects.length > 0) {
            const group = document.createElement('optgroup');
            group.label = 'Дополнительные';
            const sortedExtra = extraSubjects.slice().sort((a, b) => a.localeCompare(b, 'ru'));
            for (const subj of sortedExtra) {
                const option = document.createElement('option');
                option.value = subj;
                option.textContent = subj;
                option.className = 'extra-subject-option';
                group.appendChild(option);
            }
            select.appendChild(group);
        }

        select.disabled = false;
    }

    /**
     * Сбросить выпадающий список предметов.
     */
    function resetSubjectSelect() {
        const select = document.getElementById('subjectSelect');
        select.innerHTML = '<option value="">— Сначала выберите класс —</option>';
        select.disabled = true;
    }

    /**
     * Отрисовать галочки доп. предметов для выбранного класса.
     * @param {string[]} subjects — все доступные для включения предметы
     * @param {Set} enabledSet — включённые предметы
     * @param {Set} highlightedSet — предметы, исключённые жёлтой заливкой
     * @param {Function} onChange — callback(subject, enabled)
     */
    function renderExtraSubjectCheckboxes(subjects, enabledSet, highlightedSet, onChange) {
        const section = document.getElementById('extraSubjectsSection');
        const list = document.getElementById('extraSubjectsList');

        if (!subjects || subjects.length === 0) {
            section.hidden = true;
            return;
        }

        section.hidden = false;
        list.innerHTML = '';

        const sorted = subjects.slice().sort((a, b) => a.localeCompare(b, 'ru'));
        for (const subj of sorted) {
            const isHL = highlightedSet && highlightedSet.has(subj);
            const label = document.createElement('label');
            label.className = isHL
                ? 'extra-subject-label extra-subject-highlighted'
                : 'extra-subject-label';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = enabledSet.has(subj);
            checkbox.addEventListener('change', () => {
                onChange(subj, checkbox.checked);
            });

            const span = document.createElement('span');
            span.textContent = subj;

            label.appendChild(checkbox);
            label.appendChild(span);

            if (isHL) {
                const badge = document.createElement('span');
                badge.className = 'highlight-badge';
                badge.textContent = 'заливка';
                label.appendChild(badge);
            }

            list.appendChild(label);
        }
    }

    /**
     * Скрыть секцию галочек доп. предметов.
     */
    function hideExtraSubjectCheckboxes() {
        document.getElementById('extraSubjectsSection').hidden = true;
    }

    /**
     * Отрисовать результаты расписания.
     */
    function renderResults(records, className, subject, isExtra) {
        const section = document.getElementById('resultsSection');
        const title = document.getElementById('resultsTitle');
        const count = document.getElementById('resultsCount');
        const tbody = document.getElementById('resultsBody');

        if (!records || records.length === 0) {
            section.hidden = true;
            return;
        }

        section.hidden = false;
        title.textContent = `${className} — ${subject}`;
        const extraNote = isExtra ? ' (дополнительный предмет)' : '';
        count.textContent = `Найдено занятий: ${records.length}${extraNote}`;

        // Сортировка по дню недели, затем по номеру урока
        const sorted = records.slice().sort((a, b) => {
            const dayA = DAY_ORDER.indexOf(a.dayOfWeek);
            const dayB = DAY_ORDER.indexOf(b.dayOfWeek);
            if (dayA !== dayB) return dayA - dayB;
            return (a.lessonNumber || 0) - (b.lessonNumber || 0);
        });

        let html = '';
        for (const rec of sorted) {
            const timeStr = rec.startTime && rec.endTime
                ? `${rec.startTime}–${rec.endTime}`
                : '—';

            const sourceClass = getTimeSourceClass(rec.timeSource);

            html += `
                <tr>
                    <td><strong>${rec.dayOfWeek || '—'}</strong></td>
                    <td>${rec.lessonNumber || '—'}</td>
                    <td>${timeStr}</td>
                    <td>${rec.shift || '—'}</td>
                    <td><span class="time-source ${sourceClass}">${rec.timeSource}</span></td>
                </tr>
            `;
        }

        tbody.innerHTML = html;
    }

    /**
     * Скрыть результаты.
     */
    function hideResults() {
        document.getElementById('resultsSection').hidden = true;
    }

    /**
     * Отрисовать диагностику.
     */
    function renderDiagnostics(diagnostics, filterStats, excluded) {
        const section = document.getElementById('diagnosticsSection');
        const content = document.getElementById('diagnosticsContent');
        section.hidden = false;

        let html = '<div class="diagnostics-block">';

        // Обработанные листы
        html += '<h4>Обработка листов</h4>';
        html += `<p>Обработано листов: ${diagnostics.sheetsProcessed} из ${diagnostics.sheetsTotal}</p>`;
        html += `<p>Листы: ${diagnostics.sheetNames.join(', ')}</p>`;

        // Определённые колонки
        html += '<h4>Определённые колонки по листам</h4>';
        for (const [sheet, cols] of Object.entries(diagnostics.columnsDetected)) {
            html += `<p><strong>${sheet}:</strong> ${cols.join(', ') || 'не определены'}</p>`;
        }

        // Статистика фильтрации
        html += '<h4>Статистика фильтрации</h4>';
        html += `<ul>`;
        html += `<li>Всего записей до фильтрации: ${filterStats.total}</li>`;
        html += `<li>Прошло фильтрацию: ${filterStats.passed}</li>`;
        html += `<li>Исключено по стоп-словам: ${filterStats.excludedByStopWord}</li>`;
        if (filterStats.excludedByHighlight) {
            html += `<li>Выделено цветом (заливка фона): ${filterStats.excludedByHighlight}</li>`;
        }
        html += `<li>Без класса: ${filterStats.excludedNoClass}</li>`;
        html += `<li>Без предмета: ${filterStats.excludedNoSubject}</li>`;
        html += `</ul>`;

        // Предупреждения
        if (diagnostics.warnings.length > 0) {
            html += '<h4>Предупреждения</h4>';
            html += '<ul class="diagnostics-warnings">';
            for (const w of diagnostics.warnings) {
                html += `<li class="warning-item">${escapeHtml(w)}</li>`;
            }
            html += '</ul>';
        }

        // Пропущенные строки
        html += `<p>Пропущено пустых/служебных строк: ${diagnostics.skippedRows}</p>`;

        // Примеры исключённых записей
        if (excluded.length > 0) {
            html += '<h4>Примеры исключённых записей (первые 10)</h4>';
            html += '<div class="excluded-examples">';
            const examples = excluded.slice(0, 10);
            for (const ex of examples) {
                html += `<div class="excluded-item">`;
                html += `<span class="excluded-reason">${escapeHtml(ex.exclusionReason)}</span>`;
                html += `<span class="excluded-detail">`;
                if (ex.className) html += `Класс: ${escapeHtml(ex.className)}; `;
                if (ex.subject) html += `Предмет: ${escapeHtml(ex.subject)}; `;
                html += `</span>`;
                html += `</div>`;
            }
            html += '</div>';
        }

        html += '</div>';
        content.innerHTML = html;
    }

    /**
     * Отрисовать debug-информацию.
     */
    function renderDebug(records) {
        const section = document.getElementById('debugSection');
        const content = document.getElementById('debugContent');
        section.hidden = false;

        const examples = records.slice(0, 10);

        let html = `<p>Всего нормализованных записей: ${records.length}. Показано первых ${examples.length}:</p>`;
        html += '<div class="debug-records">';

        for (const rec of examples) {
            html += `<pre class="debug-record">${escapeHtml(JSON.stringify(rec, null, 2))}</pre>`;
        }

        html += '</div>';
        content.innerHTML = html;
    }

    /**
     * Показать секцию фильтров.
     */
    function showFilters() {
        document.getElementById('filtersSection').hidden = false;
    }

    /**
     * Сортировка классов: сначала по числу, потом по букве.
     */
    function sortClasses(classes) {
        return classes.slice().sort((a, b) => {
            const matchA = a.match(/^(\d+)(.*)$/);
            const matchB = b.match(/^(\d+)(.*)$/);
            if (matchA && matchB) {
                const numA = parseInt(matchA[1], 10);
                const numB = parseInt(matchB[1], 10);
                if (numA !== numB) return numA - numB;
                return matchA[2].localeCompare(matchB[2], 'ru');
            }
            return a.localeCompare(b, 'ru');
        });
    }

    /**
     * CSS-класс для источника времени.
     */
    function getTimeSourceClass(source) {
        if (!source) return '';
        if (source.includes('сетка')) return 'source-grid';
        if (source.includes('Excel')) return 'source-excel';
        return 'source-unknown';
    }

    /**
     * Экспорт записей в CSV.
     */
    function exportCSV(records, className, subject) {
        const BOM = '\uFEFF';
        const header = 'День занятия;Урок;Время начала;Время окончания;Смена;Источник времени\n';
        const lines = records.map(r => {
            return [
                r.dayOfWeek || '',
                r.lessonNumber || '',
                r.startTime || '',
                r.endTime || '',
                r.shift || '',
                r.timeSource || ''
            ].join(';');
        });

        const csv = BOM + header + lines.join('\n');
        downloadFile(csv, `расписание_${className}_${subject}.csv`, 'text/csv;charset=utf-8');
    }

    /**
     * Экспорт записей в JSON.
     */
    function exportJSON(records, className, subject) {
        const json = JSON.stringify(records, null, 2);
        downloadFile(json, `расписание_${className}_${subject}.json`, 'application/json;charset=utf-8');
    }

    /**
     * Скачать файл.
     */
    function downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    /**
     * Экранировать HTML.
     */
    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    return {
        renderTimeGrids,
        renderStatus,
        populateClassSelect,
        populateSubjectSelect,
        resetSubjectSelect,
        renderExtraSubjectCheckboxes,
        hideExtraSubjectCheckboxes,
        renderResults,
        hideResults,
        renderDiagnostics,
        renderDebug,
        showFilters,
        exportCSV,
        exportJSON
    };
})();
