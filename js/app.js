/**
 * Главный модуль приложения.
 *
 * Отвечает за:
 * - Инициализацию
 * - Управление состоянием
 * - Координацию между модулями
 * - Обработку событий UI
 */

const App = (() => {
    'use strict';

    // Состояние приложения
    const state = {
        rawRecords: [],
        normalizedRecords: [],
        filteredRecords: [],
        excludedRecords: [],
        diagnostics: null,
        filterStats: null,

        // Индексы
        classes: [],
        subjectsByClass: {},   // { className: [subjects] }
        recordsByClassSubject: {}, // { 'className|subject': [records] }

        // UI state
        selectedClass: null,
        selectedSubject: null,
        timeMode: 'mixed'
    };

    /**
     * Инициализация приложения.
     */
    function init() {
        // Отрисовать сетки времени
        UIRenderer.renderTimeGrids();

        // Привязать обработчики событий
        bindEvents();
    }

    /**
     * Привязка событий UI.
     */
    function bindEvents() {
        // Загрузка файла
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');
        const selectFileBtn = document.getElementById('selectFileBtn');
        const clearFileBtn = document.getElementById('clearFileBtn');

        selectFileBtn.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        });

        // Drag & Drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('drag-over');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('drag-over');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) {
                handleFile(e.dataTransfer.files[0]);
            }
        });

        // Очистка файла
        clearFileBtn.addEventListener('click', resetAll);

        // Селекты
        document.getElementById('classSelect').addEventListener('change', onClassChange);
        document.getElementById('subjectSelect').addEventListener('change', onSubjectChange);

        // Кнопки
        document.getElementById('resetBtn').addEventListener('click', resetSelection);
        document.getElementById('exportCsvBtn').addEventListener('click', exportCSV);
        document.getElementById('exportJsonBtn').addEventListener('click', exportJSON);

        // Режим времени
        document.querySelectorAll('input[name="timeMode"]').forEach(radio => {
            radio.addEventListener('change', onTimeModeChange);
        });
    }

    /**
     * Обработка загруженного файла.
     */
    async function handleFile(file) {
        const validExtensions = ['.xlsx', '.xls'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();

        if (!validExtensions.includes(ext)) {
            alert('Пожалуйста, загрузите файл в формате .xlsx или .xls');
            return;
        }

        // Показать имя файла
        document.getElementById('fileInfo').hidden = false;
        document.getElementById('fileName').textContent = file.name;

        try {
            const data = await readFileAsArrayBuffer(file);
            processExcelData(data);
        } catch (error) {
            console.error('Ошибка чтения файла:', error);
            alert('Ошибка чтения файла: ' + error.message);
        }
    }

    /**
     * Чтение файла как ArrayBuffer.
     */
    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(new Uint8Array(e.target.result));
            reader.onerror = (e) => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Обработка данных Excel.
     */
    function processExcelData(data) {
        // Этап 1: Парсинг
        const parseResult = ExcelParser.parseExcelFile(data);
        state.rawRecords = parseResult.rawRecords;
        state.diagnostics = parseResult.diagnostics;

        // Этап 2: Нормализация
        state.normalizedRecords = DataNormalizer.normalizeAll(
            state.rawRecords,
            state.timeMode
        );

        // Этап 3: Фильтрация
        const filterResult = DataFilter.filterRecords(state.normalizedRecords);
        state.filteredRecords = filterResult.filtered;
        state.excludedRecords = filterResult.excluded;
        state.filterStats = filterResult.stats;

        // Этап 4: Построение индексов
        buildIndices();

        // Этап 5: Рендер UI
        renderAfterParse();
    }

    /**
     * Построить индексы для быстрого доступа.
     */
    function buildIndices() {
        const classSet = new Set();
        const subjectsByClass = {};
        const recordsByClassSubject = {};

        for (const record of state.filteredRecords) {
            if (!record.className) continue;

            classSet.add(record.className);

            if (!subjectsByClass[record.className]) {
                subjectsByClass[record.className] = new Set();
            }
            if (record.subject) {
                subjectsByClass[record.className].add(record.subject);
            }

            const key = `${record.className}|${record.subject}`;
            if (!recordsByClassSubject[key]) {
                recordsByClassSubject[key] = [];
            }
            recordsByClassSubject[key].push(record);
        }

        state.classes = Array.from(classSet);
        state.subjectsByClass = {};
        for (const [cls, subjects] of Object.entries(subjectsByClass)) {
            state.subjectsByClass[cls] = Array.from(subjects);
        }
        state.recordsByClassSubject = recordsByClassSubject;
    }

    /**
     * Рендер UI после парсинга.
     */
    function renderAfterParse() {
        // Подсчёт уникальных предметов
        const allSubjects = new Set();
        for (const subjects of Object.values(state.subjectsByClass)) {
            subjects.forEach(s => allSubjects.add(s));
        }

        UIRenderer.renderStatus({
            totalRecords: state.filteredRecords.length,
            totalClasses: state.classes.length,
            totalSubjects: allSubjects.size,
            totalExcluded: state.excludedRecords.length,
            sheetsProcessed: state.diagnostics.sheetsProcessed
        });

        UIRenderer.showFilters();
        UIRenderer.populateClassSelect(state.classes);
        UIRenderer.resetSubjectSelect();
        UIRenderer.hideResults();

        // Диагностика
        UIRenderer.renderDiagnostics(
            state.diagnostics,
            state.filterStats,
            state.excludedRecords
        );

        // Debug
        UIRenderer.renderDebug(state.filteredRecords);

        // Активировать кнопку сброса
        document.getElementById('resetBtn').disabled = false;
    }

    /**
     * Обработка выбора класса.
     */
    function onClassChange(e) {
        state.selectedClass = e.target.value || null;
        state.selectedSubject = null;

        if (state.selectedClass && state.subjectsByClass[state.selectedClass]) {
            UIRenderer.populateSubjectSelect(state.subjectsByClass[state.selectedClass]);
        } else {
            UIRenderer.resetSubjectSelect();
        }

        UIRenderer.hideResults();
        updateExportButtons();
    }

    /**
     * Обработка выбора предмета.
     */
    function onSubjectChange(e) {
        state.selectedSubject = e.target.value || null;

        if (state.selectedClass && state.selectedSubject) {
            const key = `${state.selectedClass}|${state.selectedSubject}`;
            const records = state.recordsByClassSubject[key] || [];
            UIRenderer.renderResults(records, state.selectedClass, state.selectedSubject);
        } else {
            UIRenderer.hideResults();
        }

        updateExportButtons();
    }

    /**
     * Смена режима времени.
     */
    function onTimeModeChange(e) {
        state.timeMode = e.target.value;

        // Перенормализовать данные
        state.normalizedRecords = DataNormalizer.normalizeAll(
            state.rawRecords,
            state.timeMode
        );

        const filterResult = DataFilter.filterRecords(state.normalizedRecords);
        state.filteredRecords = filterResult.filtered;
        state.excludedRecords = filterResult.excluded;
        state.filterStats = filterResult.stats;

        buildIndices();
        renderAfterParse();

        // Восстановить выбор если возможно
        if (state.selectedClass) {
            document.getElementById('classSelect').value = state.selectedClass;
            onClassChange({ target: { value: state.selectedClass } });

            if (state.selectedSubject) {
                document.getElementById('subjectSelect').value = state.selectedSubject;
                onSubjectChange({ target: { value: state.selectedSubject } });
            }
        }
    }

    /**
     * Сброс выбора фильтров.
     */
    function resetSelection() {
        state.selectedClass = null;
        state.selectedSubject = null;

        document.getElementById('classSelect').value = '';
        UIRenderer.resetSubjectSelect();
        UIRenderer.hideResults();
        updateExportButtons();
    }

    /**
     * Полный сброс.
     */
    function resetAll() {
        state.rawRecords = [];
        state.normalizedRecords = [];
        state.filteredRecords = [];
        state.excludedRecords = [];
        state.diagnostics = null;
        state.filterStats = null;
        state.classes = [];
        state.subjectsByClass = {};
        state.recordsByClassSubject = {};
        state.selectedClass = null;
        state.selectedSubject = null;

        document.getElementById('fileInput').value = '';
        document.getElementById('fileInfo').hidden = true;
        document.getElementById('statusSection').hidden = true;
        document.getElementById('filtersSection').hidden = true;
        document.getElementById('resultsSection').hidden = true;
        document.getElementById('diagnosticsSection').hidden = true;
        document.getElementById('debugSection').hidden = true;

        document.getElementById('classSelect').disabled = true;
        document.getElementById('classSelect').value = '';
        UIRenderer.resetSubjectSelect();

        document.getElementById('resetBtn').disabled = true;
        updateExportButtons();
    }

    /**
     * Обновить состояние кнопок экспорта.
     */
    function updateExportButtons() {
        const hasResults = state.selectedClass && state.selectedSubject;
        document.getElementById('exportCsvBtn').disabled = !hasResults;
        document.getElementById('exportJsonBtn').disabled = !hasResults;
    }

    /**
     * Экспорт текущих результатов в CSV.
     */
    function exportCSV() {
        if (!state.selectedClass || !state.selectedSubject) return;
        const key = `${state.selectedClass}|${state.selectedSubject}`;
        const records = state.recordsByClassSubject[key] || [];
        UIRenderer.exportCSV(records, state.selectedClass, state.selectedSubject);
    }

    /**
     * Экспорт текущих результатов в JSON.
     */
    function exportJSON() {
        if (!state.selectedClass || !state.selectedSubject) return;
        const key = `${state.selectedClass}|${state.selectedSubject}`;
        const records = state.recordsByClassSubject[key] || [];
        UIRenderer.exportJSON(records, state.selectedClass, state.selectedSubject);
    }

    // Запуск при загрузке страницы
    document.addEventListener('DOMContentLoaded', init);

    return { state };
})();
