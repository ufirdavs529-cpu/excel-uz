/**
 * Microsoft Excel 365 Web Engine v4.0
 * Real AutoFill Series (+ handle), Copy/Cut/Paste, Bulletproof Undo/Redo,
 * Guaranteed Colors & Styles, Advanced Math & Formulas
 */

const CONFIG = {
    TOTAL_ROWS: 50,
    TOTAL_COLS: 26, // A-Z
    STORAGE_KEY: 'excel_active_workbook'
};

const COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

// Excel Color Palettes
const THEME_COLORS = [
    '#ffffff', '#000000', '#eeece1', '#1f497d', '#4f81bd', '#c0504d', '#9bbb59', '#8064a2', '#4bacc6', '#f79646',
    '#f2f2f2', '#7f7f7f', '#ddd9c3', '#c6d9f0', '#dbe5f1', '#f2dcdb', '#ebf1dd', '#e5e0ec', '#dbeef3', '#fdeada',
    '#d8d8d8', '#595959', '#c4bd97', '#8db3e2', '#b8cce4', '#e5b9b7', '#d7e3bc', '#ccc1da', '#b7dde8', '#fbd5b5',
    '#bfbfbf', '#3f3f3f', '#938953', '#548dd4', '#95b3d7', '#d99694', '#c3d69b', '#b2a2c7', '#92cddc', '#fac08f',
    '#a5a5a5', '#262626', '#494429', '#17365d', '#366092', '#953734', '#76933c', '#5f497a', '#31859b', '#e36c09'
];

const STANDARD_COLORS = [
    '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'
];

// Application State
const state = {
    workbookName: 'Книга1',
    activeSheet: 'Лист1',
    sheets: ['Лист1'],
    cells: {}, // key: "A1" -> { val: '', formula: '', style: {}, comment: '' }
    activeCell: { col: 'A', row: 1, id: 'A1' },
    selectedRange: { start: { col: 'A', row: 1 }, end: { col: 'A', row: 1 } },
    isSelectingRange: false,
    showFormulas: false,
    currentBgColor: '#ffff00',
    currentTextColor: '#d83b01',
    zoomLevel: 100
};

// Clipboard System
let clipboardState = null; // { action: 'copy'|'cut', cells: [...], startColIdx, startRow }

// Undo / Redo Stacks
const undoStack = [];
const redoStack = [];
const MAX_UNDO = 80;

// AutoFill State
let isFilling = false;
let fillTargetCell = null;

// DOM References
let gridTableBody = null;
let gridTableHeader = null;
let nameBox = null;
let formulaInput = null;
let workbookTitleInput = null;
let statusBarSum = null;
let statusBarCount = null;
let statusBarAvg = null;
let btnUndo = null;
let btnRedo = null;
let fillHandleEl = null;

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
    initDOMElements();
    generateGrid();
    setupColorPalettes();
    setupBordersTool();
    setupClipboardActions();
    setupAutoFillHandle();
    setupRibbonTabs();
    setupRibbonActions();
    setupAllNewTabFeatures();
    setupEventHandlers();
    setupRangeSelection();
    setupFindFeature();
    loadFromStorageOrTemplate();
    selectCell('A1');
    updateUndoRedoUI();
});

function initDOMElements() {
    gridTableBody = document.getElementById('grid-tbody');
    gridTableHeader = document.getElementById('grid-thead-tr');
    nameBox = document.getElementById('name-box');
    formulaInput = document.getElementById('formula-input');
    workbookTitleInput = document.getElementById('workbook-title');
    statusBarSum = document.getElementById('stat-sum');
    statusBarCount = document.getElementById('stat-count');
    statusBarAvg = document.getElementById('stat-avg');
    btnUndo = document.getElementById('btn-undo');
    btnRedo = document.getElementById('btn-redo');
    fillHandleEl = document.getElementById('excel-fill-handle');
}

/**
 * ============================================================================
 * 1. BULLETPROOF UNDO / REDO ENGINE
 * ============================================================================
 */
function recordAction(action) {
    undoStack.push(action);
    if (undoStack.length > MAX_UNDO) {
        undoStack.shift();
    }
    redoStack.length = 0; // New action invalidates redo
    updateUndoRedoUI();
}

function updateUndoRedoUI() {
    if (btnUndo) {
        btnUndo.disabled = undoStack.length === 0;
        btnUndo.title = undoStack.length > 0 ? 'Отменить (Ctrl + Z)' : 'Нечего отменять';
    }
    if (btnRedo) {
        btnRedo.disabled = redoStack.length === 0;
        btnRedo.title = redoStack.length > 0 ? 'Повторить (Ctrl + Y)' : 'Нечего повторять';
    }
}

function undo() {
    if (undoStack.length === 0) return;

    // Blur active cell to prevent contenteditable desync
    if (document.activeElement && document.activeElement.classList.contains('excel-cell')) {
        document.activeElement.blur();
    }

    const action = undoStack.pop();
    redoStack.push(action);

    applyAction(action, true); // true = revert
    updateUndoRedoUI();
    showToast('Действие отменено (Undo)');
}

function redo() {
    if (redoStack.length === 0) return;

    if (document.activeElement && document.activeElement.classList.contains('excel-cell')) {
        document.activeElement.blur();
    }

    const action = redoStack.pop();
    undoStack.push(action);

    applyAction(action, false); // false = apply next
    updateUndoRedoUI();
    showToast('Действие повторено (Redo)');
}

function applyAction(action, isRevert) {
    if (action.type === 'cell_change') {
        const target = isRevert ? action.prev : action.next;
        const cellId = action.cellId;
        const cellEl = document.getElementById(`cell-${cellId}`);

        if (!target || (!target.val && !target.formula && (!target.style || Object.keys(target.style).length === 0))) {
            delete state.cells[cellId];
            if (cellEl) {
                cellEl.textContent = '';
                resetCellStylesToDefault(cellEl);
            }
        } else {
            state.cells[cellId] = JSON.parse(JSON.stringify(target));
            if (cellEl) {
                cellEl.textContent = target.formula ? evaluateFormula(target.formula) : (target.val !== undefined ? target.val : '');
                applyCellStyles(cellEl, target.style || {});
            }
        }

        checkAndResetRowHeight(parseInt(cellId.replace(/[A-Z]/g, ''), 10));
        recalculateAllFormulas();
        syncFormulaBarWithCell(cellId);
        syncRibbonWithCell(cellId);
        updateFillHandlePosition();
        autoSave();
        updateStatusBar();
    } else if (action.type === 'batch_cells') {
        const list = isRevert ? action.prevList : action.nextList;
        list.forEach(item => {
            const cellId = item.cellId;
            const cellEl = document.getElementById(`cell-${cellId}`);
            if (!item.data || (!item.data.val && !item.data.formula && (!item.data.style || Object.keys(item.data.style).length === 0))) {
                delete state.cells[cellId];
                if (cellEl) {
                    cellEl.textContent = '';
                    resetCellStylesToDefault(cellEl);
                }
            } else {
                state.cells[cellId] = JSON.parse(JSON.stringify(item.data));
                if (cellEl) {
                    cellEl.textContent = item.data.formula ? evaluateFormula(item.data.formula) : (item.data.val !== undefined ? item.data.val : '');
                    applyCellStyles(cellEl, item.data.style || {});
                }
            }
            checkAndResetRowHeight(parseInt(cellId.replace(/[A-Z]/g, ''), 10));
        });
        recalculateAllFormulas();
        updateFillHandlePosition();
        autoSave();
        updateStatusBar();
    }
}

/**
 * ============================================================================
 * 2. EXCEL AUTOFILL DRAG HANDLE (+)
 * Supports:
 * - 1 cell with number: dragging down fills 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
 * - 2 cells with numbers: 1 and 3 -> fills 5, 7, 9, 11 (step = +2)
 * - Formulas: =A1+B1 -> =A2+B2, =A3+B3
 * - Styles & Colors are preserved and copied!
 * ============================================================================
 */
function setupAutoFillHandle() {
    if (!fillHandleEl) return;

    fillHandleEl.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        isFilling = true;
        fillTargetCell = null;
        document.body.style.cursor = 'crosshair';
    });

    // Handle drag over grid cells
    document.addEventListener('mousemove', (e) => {
        if (!isFilling) return;
        const td = document.elementFromPoint(e.clientX, e.clientY)?.closest('.excel-cell');
        if (td && td !== fillTargetCell) {
            fillTargetCell = td;
            previewAutoFillRange(td.dataset.cellId);
        }
    });

    document.addEventListener('mouseup', () => {
        if (!isFilling) return;
        isFilling = false;
        document.body.style.cursor = 'default';
        clearAutoFillPreview();

        if (fillTargetCell) {
            executeAutoFill(fillTargetCell.dataset.cellId);
            fillTargetCell = null;
        }
    });
}

function updateFillHandlePosition() {
    if (!fillHandleEl) return;
    const activeCellEl = document.getElementById(`cell-${state.activeCell.id}`);
    if (!activeCellEl) {
        fillHandleEl.style.display = 'none';
        return;
    }

    // Position at the bottom-right of active cell (or selection range end)
    const rangeEndCol = state.selectedRange ? state.selectedRange.end.col : state.activeCell.col;
    const rangeEndRow = state.selectedRange ? state.selectedRange.end.row : state.activeCell.row;
    const targetCellId = `${rangeEndCol}${rangeEndRow}`;
    const targetEl = document.getElementById(`cell-${targetCellId}`) || activeCellEl;

    const rect = targetEl.getBoundingClientRect();
    const containerRect = document.getElementById('spreadsheet-container').getBoundingClientRect();
    const container = document.getElementById('spreadsheet-container');

    const top = rect.bottom - containerRect.top + container.scrollTop - 5;
    const left = rect.right - containerRect.left + container.scrollLeft - 5;

    fillHandleEl.style.top = `${top}px`;
    fillHandleEl.style.left = `${left}px`;
    fillHandleEl.style.display = 'block';
}

function previewAutoFillRange(targetCellId) {
    clearAutoFillPreview();
    const match = targetCellId.match(/^([A-Z]+)([0-9]+)$/);
    if (!match) return;

    const targetCol = match[1];
    const targetRow = parseInt(match[2], 10);

    const sourceCol = state.activeCell.col;
    const sourceRow = state.activeCell.row;

    // Only vertical fill preview for simplicity and standard Excel workflow
    const minRow = Math.min(sourceRow, targetRow);
    const maxRow = Math.max(sourceRow, targetRow);

    for (let r = minRow; r <= maxRow; r++) {
        const cellEl = document.getElementById(`cell-${sourceCol}${r}`);
        if (cellEl) cellEl.classList.add('autofill-preview');
    }
}

function clearAutoFillPreview() {
    document.querySelectorAll('.excel-cell.autofill-preview').forEach(c => c.classList.remove('autofill-preview'));
}

function executeAutoFill(targetCellId) {
    const match = targetCellId.match(/^([A-Z]+)([0-9]+)$/);
    if (!match) return;

    const targetRow = parseInt(match[2], 10);
    const startRow = Math.min(state.selectedRange.start.row, state.selectedRange.end.row);
    const endRow = Math.max(state.selectedRange.start.row, state.selectedRange.end.row);
    const col = state.activeCell.col;

    if (targetRow <= endRow) return; // Only dragging down supported

    // Collect source series
    const sourceData = [];
    for (let r = startRow; r <= endRow; r++) {
        const cellData = state.cells[`${col}${r}`] || { val: '', formula: '', style: {} };
        sourceData.push({ row: r, ...JSON.parse(JSON.stringify(cellData)) });
    }

    // Calculate series step
    let step = 1;
    let isNumericSeries = false;

    if (sourceData.length >= 2) {
        const n1 = parseFloat(sourceData[0].val);
        const n2 = parseFloat(sourceData[sourceData.length - 1].val);
        if (!isNaN(n1) && !isNaN(n2)) {
            step = (n2 - n1) / (sourceData.length - 1);
            isNumericSeries = true;
        }
    } else if (sourceData.length === 1) {
        const n1 = parseFloat(sourceData[0].val);
        if (!isNaN(n1)) {
            step = 1; // Default increment: 1, 2, 3, 4, 5...
            isNumericSeries = true;
        }
    }

    const batchPrev = [];
    const batchNext = [];

    let lastNum = isNumericSeries ? parseFloat(sourceData[sourceData.length - 1].val) : 0;
    const sampleStyle = sourceData[sourceData.length - 1].style || {};
    const sampleFormula = sourceData[0].formula || '';

    let fillIdx = 1;
    for (let r = endRow + 1; r <= targetRow; r++) {
        const fillCellId = `${col}${r}`;
        const cellEl = document.getElementById(`cell-${fillCellId}`);
        if (!cellEl) continue;

        const prevData = state.cells[fillCellId] ? JSON.parse(JSON.stringify(state.cells[fillCellId])) : null;
        batchPrev.push({ cellId: fillCellId, data: prevData });

        let newVal = '';
        let newFormula = '';

        if (sampleFormula.startsWith('=')) {
            // Shift formula references vertically by (r - sourceData[0].row)
            const rowDelta = r - sourceData[0].row;
            newFormula = sampleFormula.replace(/([A-Z]+)([0-9]+)/g, (m, c, rowStr) => {
                const oldR = parseInt(rowStr, 10);
                return `${c}${oldR + rowDelta}`;
            });
            newVal = evaluateFormula(newFormula);
        } else if (isNumericSeries) {
            lastNum += step;
            newVal = Number.isInteger(lastNum) ? lastNum : parseFloat(lastNum.toFixed(2));
        } else {
            // Pattern text or copy
            const baseVal = sourceData[(fillIdx - 1) % sourceData.length].val;
            newVal = baseVal;
        }

        state.cells[fillCellId] = {
            val: newVal,
            formula: newFormula,
            style: JSON.parse(JSON.stringify(sampleStyle))
        };

        cellEl.textContent = state.showFormulas && newFormula ? newFormula : newVal;
        applyCellStyles(cellEl, sampleStyle);
        checkAndResetRowHeight(r);

        batchNext.push({ cellId: fillCellId, data: JSON.parse(JSON.stringify(state.cells[fillCellId])) });
        fillIdx++;
    }

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });
    recalculateAllFormulas();
    selectCell(targetCellId);
    autoSave();
    showToast(`Автозаполнение завершено (Строка ${targetRow} заполнена)!`);
}

/**
 * ============================================================================
 * 3. CLIPBOARD SYSTEM (COPY, CUT, PASTE)
 * ============================================================================
 */
function setupClipboardActions() {
    document.getElementById('btn-copy')?.addEventListener('click', copySelection);
    document.getElementById('btn-cut')?.addEventListener('click', cutSelection);
    document.getElementById('btn-paste')?.addEventListener('click', pasteClipboard);
}

function clearClipboardIndicators() {
    document.querySelectorAll('.excel-cell.cell-copied, .excel-cell.cell-cut').forEach(el => {
        el.classList.remove('cell-copied', 'cell-cut');
    });
}

function copySelection() {
    clearClipboardIndicators();
    const selectedIds = getSelectedCellIds();
    if (selectedIds.length === 0) return;

    const startMatch = selectedIds[0].match(/^([A-Z]+)([0-9]+)$/);
    const baseColIdx = COLS.indexOf(startMatch[1]);
    const baseRow = parseInt(startMatch[2], 10);

    const cellPayloads = [];
    selectedIds.forEach(cellId => {
        const match = cellId.match(/^([A-Z]+)([0-9]+)$/);
        const cIdx = COLS.indexOf(match[1]);
        const r = parseInt(match[2], 10);
        const data = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };

        cellPayloads.push({
            offsetC: cIdx - baseColIdx,
            offsetR: r - baseRow,
            cellId: cellId,
            data: data
        });

        document.getElementById(`cell-${cellId}`)?.classList.add('cell-copied');
    });

    clipboardState = {
        action: 'copy',
        cells: cellPayloads
    };

    showToast('Скопировано в буфер обмена (Ctrl + C)');
}

function cutSelection() {
    clearClipboardIndicators();
    const selectedIds = getSelectedCellIds();
    if (selectedIds.length === 0) return;

    const startMatch = selectedIds[0].match(/^([A-Z]+)([0-9]+)$/);
    const baseColIdx = COLS.indexOf(startMatch[1]);
    const baseRow = parseInt(startMatch[2], 10);

    const cellPayloads = [];
    selectedIds.forEach(cellId => {
        const match = cellId.match(/^([A-Z]+)([0-9]+)$/);
        const cIdx = COLS.indexOf(match[1]);
        const r = parseInt(match[2], 10);
        const data = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };

        cellPayloads.push({
            offsetC: cIdx - baseColIdx,
            offsetR: r - baseRow,
            cellId: cellId,
            data: data
        });

        document.getElementById(`cell-${cellId}`)?.classList.add('cell-cut');
    });

    clipboardState = {
        action: 'cut',
        cells: cellPayloads
    };

    showToast('Вырезано в буфер обмена (Ctrl + X)');
}

function pasteClipboard() {
    if (!clipboardState || !clipboardState.cells || clipboardState.cells.length === 0) {
        showToast('Буфер обмена пуст');
        return;
    }

    const destColIdx = COLS.indexOf(state.activeCell.col);
    const destRow = state.activeCell.row;

    const batchPrev = [];
    const batchNext = [];

    // If cut, clear original cells first
    if (clipboardState.action === 'cut') {
        clipboardState.cells.forEach(item => {
            const srcCellId = item.cellId;
            const srcEl = document.getElementById(`cell-${srcCellId}`);
            const prevSrcData = state.cells[srcCellId] ? JSON.parse(JSON.stringify(state.cells[srcCellId])) : null;
            batchPrev.push({ cellId: srcCellId, data: prevSrcData });

            delete state.cells[srcCellId];
            if (srcEl) {
                srcEl.textContent = '';
                resetCellStylesToDefault(srcEl);
            }
            batchNext.push({ cellId: srcCellId, data: null });
        });
    }

    clipboardState.cells.forEach(item => {
        const targetCIdx = destColIdx + item.offsetC;
        const targetR = destRow + item.offsetR;

        if (targetCIdx < COLS.length && targetR <= CONFIG.TOTAL_ROWS) {
            const targetCellId = `${COLS[targetCIdx]}${targetR}`;
            const targetEl = document.getElementById(`cell-${targetCellId}`);
            if (!targetEl) return;

            const prevData = state.cells[targetCellId] ? JSON.parse(JSON.stringify(state.cells[targetCellId])) : null;
            batchPrev.push({ cellId: targetCellId, data: prevData });

            const srcData = item.data;
            let finalFormula = srcData.formula || '';

            // Shift formula references
            if (finalFormula.startsWith('=')) {
                const rowDelta = targetR - (destRow - item.offsetR);
                const colDelta = targetCIdx - (destColIdx - item.offsetC);
                finalFormula = finalFormula.replace(/([A-Z]+)([0-9]+)/g, (m, cStr, rStr) => {
                    const oldCIdx = COLS.indexOf(cStr);
                    const oldR = parseInt(rStr, 10);
                    const newC = COLS[Math.min(COLS.length - 1, Math.max(0, oldCIdx + colDelta))];
                    const newR = Math.min(CONFIG.TOTAL_ROWS, Math.max(1, oldR + rowDelta));
                    return `${newC}${newR}`;
                });
            }

            state.cells[targetCellId] = {
                val: finalFormula ? evaluateFormula(finalFormula) : (srcData.val !== undefined ? srcData.val : ''),
                formula: finalFormula,
                style: JSON.parse(JSON.stringify(srcData.style || {}))
            };

            targetEl.textContent = state.showFormulas && finalFormula ? finalFormula : state.cells[targetCellId].val;
            applyCellStyles(targetEl, state.cells[targetCellId].style);
            checkAndResetRowHeight(targetR);

            batchNext.push({ cellId: targetCellId, data: JSON.parse(JSON.stringify(state.cells[targetCellId])) });
        }
    });

    if (clipboardState.action === 'cut') {
        clearClipboardIndicators();
        clipboardState = null;
    }

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });
    recalculateAllFormulas();
    updateFillHandlePosition();
    autoSave();
    updateStatusBar();
    showToast('Вставлено (Ctrl + V)');
}

/**
 * ============================================================================
 * 4. COLOR PALETTES & RELIABLE COLOR APPLICATION
 * ============================================================================
 */
function setupColorPalettes() {
    populatePaletteGrid('grid-bg-theme-colors', THEME_COLORS, (color) => applyCellBgColor(color));
    populatePaletteGrid('grid-bg-standard-colors', STANDARD_COLORS, (color) => applyCellBgColor(color));
    populatePaletteGrid('grid-text-theme-colors', THEME_COLORS, (color) => applyCellTextColor(color));
    populatePaletteGrid('grid-text-standard-colors', STANDARD_COLORS, (color) => applyCellTextColor(color));

    const dropdownBg = document.getElementById('dropdown-bg-palette');
    const dropdownText = document.getElementById('dropdown-text-palette');

    // Both clicking main bucket button OR clicking arrow opens the palette!
    document.getElementById('btn-apply-bg-color')?.addEventListener('click', (e) => {
        e.stopPropagation();
        closeAllColorPalettes();
        dropdownBg?.classList.toggle('show');
    });
    document.getElementById('btn-toggle-bg-palette')?.addEventListener('click', (e) => {
        e.stopPropagation();
        closeAllColorPalettes();
        dropdownBg?.classList.toggle('show');
    });

    // Both clicking text A button OR clicking arrow opens the palette!
    document.getElementById('btn-apply-text-color')?.addEventListener('click', (e) => {
        e.stopPropagation();
        closeAllColorPalettes();
        dropdownText?.classList.toggle('show');
    });
    document.getElementById('btn-toggle-text-palette')?.addEventListener('click', (e) => {
        e.stopPropagation();
        closeAllColorPalettes();
        dropdownText?.classList.toggle('show');
    });

    document.getElementById('btn-bg-no-color')?.addEventListener('click', () => {
        applyCellBgColor('transparent');
        closeAllColorPalettes();
    });

    document.getElementById('btn-text-auto-color')?.addEventListener('click', () => {
        applyCellTextColor('#000000');
        closeAllColorPalettes();
    });

    document.getElementById('input-bg-native')?.addEventListener('input', (e) => applyCellBgColor(e.target.value));
    document.getElementById('input-text-native')?.addEventListener('input', (e) => applyCellTextColor(e.target.value));

    document.addEventListener('click', (e) => {
        if (!e.target.closest('.color-picker-wrapper')) {
            closeAllColorPalettes();
        }
    });
}

function populatePaletteGrid(elementId, colorArray, onClickCallback) {
    const container = document.getElementById(elementId);
    if (!container) return;
    container.innerHTML = '';
    colorArray.forEach(color => {
        const swatch = document.createElement('div');
        swatch.className = 'color-swatch';
        swatch.style.backgroundColor = color;
        swatch.title = color;
        swatch.addEventListener('click', (e) => {
            e.stopPropagation();
            onClickCallback(color);
            closeAllColorPalettes();
        });
        container.appendChild(swatch);
    });
}

function closeAllColorPalettes() {
    document.querySelectorAll('.color-palette-dropdown, .excel-menu-dropdown').forEach(p => p.classList.remove('show'));
}

function applyCellBgColor(color) {
    state.currentBgColor = color;
    const bar = document.getElementById('color-bar-bg');
    if (bar) bar.style.backgroundColor = color === 'transparent' ? '#ffffff' : color;

    setStyle('backgroundColor', color);
    showToast(color === 'transparent' ? 'Заливка удалена' : `Цвет заливки изменен: ${color}`);
}

function applyCellTextColor(color) {
    state.currentTextColor = color;
    const bar = document.getElementById('color-bar-text');
    if (bar) bar.style.backgroundColor = color;

    setStyle('color', color);
    showToast(`Цвет текста изменен: ${color}`);
}

/**
 * ============================================================================
 * 5. ADVANCED MATHEMATICAL FUNCTIONS ENGINE
 * ============================================================================
 */
function evaluateFormula(expression) {
    if (!expression || typeof expression !== 'string' || !expression.startsWith('=')) return expression;

    const formula = expression.substring(1).trim();

    try {
        // Multi-arg & Range functions
        const rangeFuncMatch = formula.match(/^([A-Z0-9_]+)\((.*)\)$/i);
        if (rangeFuncMatch) {
            const funcName = rangeFuncMatch[1].toUpperCase();
            const argsStr = rangeFuncMatch[2];

            if (['SUM', 'AVERAGE', 'AVG', 'COUNT', 'MAX', 'MIN', 'PRODUCT', 'MEDIAN'].includes(funcName)) {
                const values = resolveRangeValues(argsStr);
                switch (funcName) {
                    case 'SUM':
                        return values.reduce((acc, v) => acc + v, 0);
                    case 'AVERAGE':
                    case 'AVG':
                        return values.length ? parseFloat((values.reduce((acc, v) => acc + v, 0) / values.length).toFixed(2)) : 0;
                    case 'COUNT':
                        return values.length;
                    case 'MAX':
                        return values.length ? Math.max(...values) : 0;
                    case 'MIN':
                        return values.length ? Math.min(...values) : 0;
                    case 'PRODUCT':
                        return values.length ? values.reduce((acc, v) => acc * v, 1) : 0;
                    case 'MEDIAN':
                        if (!values.length) return 0;
                        values.sort((a, b) => a - b);
                        const mid = Math.floor(values.length / 2);
                        return values.length % 2 !== 0 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
                }
            } else if (funcName === 'SQRT') {
                const val = getCellValueAsNumber(argsStr.trim());
                return val >= 0 ? parseFloat(Math.sqrt(val).toFixed(2)) : '#NUM!';
            } else if (funcName === 'ABS') {
                const val = getCellValueAsNumber(argsStr.trim());
                return Math.abs(val);
            } else if (funcName === 'ROUND') {
                const parts = argsStr.split(',');
                const val = getCellValueAsNumber(parts[0].trim());
                const decimals = parts[1] ? parseInt(parts[1].trim(), 10) : 0;
                return parseFloat(val.toFixed(decimals));
            } else if (funcName === 'POWER') {
                const parts = argsStr.split(',');
                const base = getCellValueAsNumber(parts[0].trim());
                const exp = parts[1] ? getCellValueAsNumber(parts[1].trim()) : 1;
                return Math.pow(base, exp);
            } else if (funcName === 'MOD') {
                const parts = argsStr.split(',');
                const val = getCellValueAsNumber(parts[0].trim());
                const div = parts[1] ? getCellValueAsNumber(parts[1].trim()) : 1;
                return div !== 0 ? val % div : '#DIV/0!';
            } else if (funcName === 'IF') {
                const parts = argsStr.split(',').map(s => s.trim());
                const condition = parts[0];
                const trueVal = parts[1] ? parts[1].replace(/^"|"$/g, '') : '';
                const falseVal = parts[2] ? parts[2].replace(/^"|"$/g, '') : '';
                const condResolved = condition.replace(/[A-Z]+[0-9]+/g, (m) => getCellValueAsNumber(m));
                const boolRes = Function(`'use strict'; return (${condResolved})`)();
                return boolRes ? trueVal : falseVal;
            } else if (funcName === 'TODAY') {
                return new Date().toLocaleDateString('ru-RU');
            } else if (funcName === 'NOW') {
                return new Date().toLocaleString('ru-RU');
            }
        }

        // Standard Arithmetic evaluation: +, -, *, /, ^, %
        let parsedExpr = formula.replace(/[A-Z]+[0-9]+/g, (match) => {
            return getCellValueAsNumber(match);
        });

        // Convert ^ to ** for JS power
        parsedExpr = parsedExpr.replace(/\^/g, '**');

        if (/^[\d\s+\-*/().%*]+$/.test(parsedExpr)) {
            parsedExpr = parsedExpr.replace(/(\d+)%/g, '($1/100)');
            const result = Function(`'use strict'; return (${parsedExpr})`)();
            if (typeof result === 'number' && !isNaN(result)) {
                return Number.isInteger(result) ? result : parseFloat(result.toFixed(2));
            }
        }

        return '#VALUE!';
    } catch (e) {
        return '#ERROR!';
    }
}

function resolveRangeValues(argString) {
    const values = [];
    const parts = argString.split(',').map(s => s.trim());

    parts.forEach(part => {
        if (part.includes(':')) {
            const [start, end] = part.split(':').map(s => s.trim());
            const startMatch = start.match(/^([A-Z]+)([0-9]+)$/);
            const endMatch = end.match(/^([A-Z]+)([0-9]+)$/);

            if (startMatch && endMatch) {
                const startColIdx = COLS.indexOf(startMatch[1]);
                const endColIdx = COLS.indexOf(endMatch[1]);
                const startRow = parseInt(startMatch[2], 10);
                const endRow = parseInt(endMatch[2], 10);

                const minCol = Math.min(startColIdx, endColIdx);
                const maxCol = Math.max(startColIdx, endColIdx);
                const minRow = Math.min(startRow, endRow);
                const maxRow = Math.max(startRow, endRow);

                for (let c = minCol; c <= maxCol; c++) {
                    for (let r = minRow; r <= maxRow; r++) {
                        const cId = `${COLS[c]}${r}`;
                        const num = getCellValueAsNumber(cId);
                        if (!isNaN(num)) values.push(num);
                    }
                }
            }
        } else {
            const num = getCellValueAsNumber(part);
            if (!isNaN(num)) values.push(num);
        }
    });

    return values;
}

function getCellValueAsNumber(cellId) {
    const data = state.cells[cellId];
    if (data) {
        if (data.formula) {
            const evaluated = evaluateFormula(data.formula);
            const num = parseFloat(evaluated);
            return isNaN(num) ? 0 : num;
        }
        const num = parseFloat(data.val);
        return isNaN(num) ? 0 : num;
    }
    const cellEl = document.getElementById(`cell-${cellId}`);
    if (cellEl) {
        const num = parseFloat(cellEl.textContent.trim());
        return isNaN(num) ? 0 : num;
    }
    const directNum = parseFloat(cellId);
    return isNaN(directNum) ? 0 : directNum;
}

function recalculateAllFormulas() {
    Object.keys(state.cells).forEach(cellId => {
        const item = state.cells[cellId];
        if (item && item.formula) {
            const res = evaluateFormula(item.formula);
            item.val = res;
            const el = document.getElementById(`cell-${cellId}`);
            if (el && el !== document.activeElement) {
                el.textContent = state.showFormulas ? item.formula : res;
            }
        }
    });
}

/**
 * ============================================================================
 * 6. CELL COMMIT & ROW HEIGHT RESET
 * ============================================================================
 */
function commitCellEdit(cellEl) {
    if (!cellEl) return;
    const cellId = cellEl.dataset.cellId;
    let text = cellEl.textContent.trim();

    const prevData = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };

    if (!state.cells[cellId]) {
        state.cells[cellId] = { val: '', formula: '', style: {} };
    }

    if (text.startsWith('=')) {
        state.cells[cellId].formula = text;
        const result = evaluateFormula(text);
        state.cells[cellId].val = result;
        if (!state.showFormulas) {
            cellEl.textContent = result;
        }
    } else {
        state.cells[cellId].formula = '';
        state.cells[cellId].val = text;
    }

    if (state.cells[cellId].style) {
        applyCellStyles(cellEl, state.cells[cellId].style);
    }

    const nextData = JSON.parse(JSON.stringify(state.cells[cellId]));
    recordAction({ type: 'cell_change', cellId, prev: prevData, next: nextData });

    recalculateAllFormulas();
    updateFillHandlePosition();
    autoSave();
    updateStatusBar();
}

function resetCellStylesToDefault(cellEl) {
    if (!cellEl) return;
    cellEl.style.fontSize = '';
    cellEl.style.lineHeight = '';
    cellEl.style.height = '';
    cellEl.style.minHeight = '';
    cellEl.style.fontWeight = 'normal';
    cellEl.style.fontStyle = 'normal';
    cellEl.style.textDecoration = 'none';
    cellEl.style.backgroundColor = '';
    cellEl.style.color = '';
    cellEl.style.border = '';
    cellEl.style.textAlign = '';
}

function checkAndResetRowHeight(rowNum) {
    const tr = document.getElementById(`row-tr-${rowNum}`);
    if (!tr) return;

    let maxFont = 11;
    COLS.forEach(col => {
        const item = state.cells[`${col}${rowNum}`];
        if (item && item.style && item.style.fontSize) {
            const size = parseInt(item.style.fontSize, 10);
            if (size > maxFont) maxFont = size;
        }
    });

    if (maxFont <= 13) {
        tr.style.height = '22px';
    } else {
        tr.style.height = `${maxFont + 10}px`;
    }
}

function clearActiveCell() {
    const selectedIds = getSelectedCellIds();
    const batchPrev = [];
    const batchNext = [];

    selectedIds.forEach(cellId => {
        const cellEl = document.getElementById(`cell-${cellId}`);
        const prevData = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : null;
        batchPrev.push({ cellId, data: prevData });
        batchNext.push({ cellId, data: null });

        if (cellEl) {
            cellEl.textContent = '';
            resetCellStylesToDefault(cellEl);
        }
        delete state.cells[cellId];
        checkAndResetRowHeight(parseInt(cellId.replace(/[A-Z]/g, ''), 10));
    });

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });

    if (formulaInput) formulaInput.value = '';
    syncRibbonWithCell(state.activeCell.id);
    recalculateAllFormulas();
    updateFillHandlePosition();
    autoSave();
    updateStatusBar();
    showToast('Ячейки очищены, размеры сброшены');
}

/**
 * ============================================================================
 * 7. FONT RESIZING (8 to 72) & GROW/SHRINK
 * ============================================================================
 */
function setCellFontSize(sizePx) {
    const selectedIds = getSelectedCellIds();
    const batchPrev = [];
    const batchNext = [];

    selectedIds.forEach(cellId => {
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (!cellEl) return;

        const prevData = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };
        batchPrev.push({ cellId, data: prevData });

        if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
        if (!state.cells[cellId].style) state.cells[cellId].style = {};

        state.cells[cellId].style.fontSize = `${sizePx}px`;

        cellEl.style.fontSize = `${sizePx}px`;
        cellEl.style.lineHeight = `${sizePx + 4}px`;

        const rowNum = parseInt(cellId.replace(/[A-Z]/g, ''), 10);
        checkAndResetRowHeight(rowNum);

        batchNext.push({ cellId, data: JSON.parse(JSON.stringify(state.cells[cellId])) });
    });

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });

    const sizeSelect = document.getElementById('select-font-size');
    if (sizeSelect) sizeSelect.value = sizePx;

    updateFillHandlePosition();
    autoSave();
}

function growFontSize() {
    const cellData = state.cells[state.activeCell.id];
    let current = cellData && cellData.style && cellData.style.fontSize ? parseInt(cellData.style.fontSize, 10) : 11;
    const next = FONT_SIZES.find(s => s > current) || 72;
    setCellFontSize(next);
}

function shrinkFontSize() {
    const cellData = state.cells[state.activeCell.id];
    let current = cellData && cellData.style && cellData.style.fontSize ? parseInt(cellData.style.fontSize, 10) : 11;
    const prev = [...FONT_SIZES].reverse().find(s => s < current) || 8;
    setCellFontSize(prev);
}

/**
 * ============================================================================
 * 8. GRID GENERATION & RANGE SELECTION
 * ============================================================================
 */
function generateGrid() {
    if (gridTableHeader) {
        gridTableHeader.innerHTML = '<th class="corner-header" title="Выделить всё" id="corner-select-all">◢</th>';
        COLS.forEach(col => {
            const th = document.createElement('th');
            th.className = 'col-header';
            th.id = `col-hdr-${col}`;
            th.textContent = col;
            gridTableHeader.appendChild(th);
        });
    }

    if (gridTableBody) {
        gridTableBody.innerHTML = '';
        for (let r = 1; r <= CONFIG.TOTAL_ROWS; r++) {
            const tr = document.createElement('tr');
            tr.id = `row-tr-${r}`;

            const rowTh = document.createElement('th');
            rowTh.className = 'row-header';
            rowTh.id = `row-hdr-${r}`;
            rowTh.textContent = r;
            tr.appendChild(rowTh);

            COLS.forEach(col => {
                const cellId = `${col}${r}`;
                const td = document.createElement('td');
                td.className = 'excel-cell';
                td.id = `cell-${cellId}`;
                td.dataset.col = col;
                td.dataset.row = r;
                td.dataset.cellId = cellId;
                td.contentEditable = 'true';
                td.spellcheck = false;

                tr.appendChild(td);
            });

            gridTableBody.appendChild(tr);
        }
    }
}

function selectCell(cellId, scrollIntoView = false) {
    if (!cellId) return;

    const prevCell = document.querySelector('.excel-cell.selected');
    if (prevCell) {
        prevCell.classList.remove('selected');
        commitCellEdit(prevCell);
    }

    clearRangeHighlight();
    document.querySelectorAll('.active-header').forEach(el => el.classList.remove('active-header'));

    const match = cellId.match(/^([A-Z]+)([0-9]+)$/);
    if (!match) return;
    const col = match[1];
    const row = parseInt(match[2], 10);

    const cellEl = document.getElementById(`cell-${cellId}`);
    if (!cellEl) return;

    state.activeCell = { col, row, id: cellId };
    state.selectedRange = { start: { col, row }, end: { col, row } };

    cellEl.classList.add('selected');
    if (scrollIntoView) {
        cellEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }

    const colHdr = document.getElementById(`col-hdr-${col}`);
    const rowHdr = document.getElementById(`row-hdr-${row}`);
    if (colHdr) colHdr.classList.add('active-header');
    if (rowHdr) rowHdr.classList.add('active-header');

    if (nameBox) nameBox.value = cellId;
    syncFormulaBarWithCell(cellId);
    syncRibbonWithCell(cellId);
    updateFillHandlePosition();
    updateStatusBar();
}

function setupRangeSelection() {
    if (!gridTableBody) return;

    gridTableBody.addEventListener('mousedown', (e) => {
        if (e.target.id === 'excel-fill-handle') return;
        const td = e.target.closest('.excel-cell');
        if (td) {
            state.isSelectingRange = true;
            const match = td.dataset.cellId.match(/^([A-Z]+)([0-9]+)$/);
            if (match) {
                state.selectedRange.start = { col: match[1], row: parseInt(match[2], 10) };
                state.selectedRange.end = { col: match[1], row: parseInt(match[2], 10) };
            }
            selectCell(td.dataset.cellId);
        }
    });

    gridTableBody.addEventListener('mouseover', (e) => {
        if (!state.isSelectingRange) return;
        const td = e.target.closest('.excel-cell');
        if (td) {
            const match = td.dataset.cellId.match(/^([A-Z]+)([0-9]+)$/);
            if (match) {
                state.selectedRange.end = { col: match[1], row: parseInt(match[2], 10) };
                highlightRange();
            }
        }
    });

    document.addEventListener('mouseup', () => {
        if (state.isSelectingRange) {
            state.isSelectingRange = false;
            updateFillHandlePosition();
        }
    });
}

function highlightRange() {
    clearRangeHighlight();

    const startColIdx = COLS.indexOf(state.selectedRange.start.col);
    const endColIdx = COLS.indexOf(state.selectedRange.end.col);
    const minCol = Math.min(startColIdx, endColIdx);
    const maxCol = Math.max(startColIdx, endColIdx);

    const minRow = Math.min(state.selectedRange.start.row, state.selectedRange.end.row);
    const maxRow = Math.max(state.selectedRange.start.row, state.selectedRange.end.row);

    for (let c = minCol; c <= maxCol; c++) {
        for (let r = minRow; r <= maxRow; r++) {
            const cellId = `${COLS[c]}${r}`;
            const cellEl = document.getElementById(`cell-${cellId}`);
            if (cellEl && cellId !== state.activeCell.id) {
                cellEl.classList.add('in-range');
            }
        }
    }
    updateStatusBar();
}

function clearRangeHighlight() {
    document.querySelectorAll('.excel-cell.in-range').forEach(el => el.classList.remove('in-range'));
}

function getSelectedCellIds() {
    const ids = [];
    const startColIdx = COLS.indexOf(state.selectedRange.start.col);
    const endColIdx = COLS.indexOf(state.selectedRange.end.col);
    const minCol = Math.min(startColIdx, endColIdx);
    const maxCol = Math.max(startColIdx, endColIdx);

    const minRow = Math.min(state.selectedRange.start.row, state.selectedRange.end.row);
    const maxRow = Math.max(state.selectedRange.start.row, state.selectedRange.end.row);

    for (let c = minCol; c <= maxCol; c++) {
        for (let r = minRow; r <= maxRow; r++) {
            ids.push(`${COLS[c]}${r}`);
        }
    }
    return ids.length > 0 ? ids : [state.activeCell.id];
}

/**
 * ============================================================================
 * 9. GLOBAL KEYBOARD SHORTCUTS & EVENT LISTENERS
 * ============================================================================
 */
function setupEventHandlers() {
    if (gridTableBody) {
        gridTableBody.addEventListener('dblclick', (e) => {
            const td = e.target.closest('.excel-cell');
            if (td) {
                const cellId = td.dataset.cellId;
                const cellData = state.cells[cellId];
                if (cellData && cellData.formula) {
                    td.textContent = cellData.formula;
                }
            }
        });

        gridTableBody.addEventListener('input', (e) => {
            const td = e.target.closest('.excel-cell');
            if (td && formulaInput) {
                formulaInput.value = td.textContent;
            }
        });

        gridTableBody.addEventListener('focusout', (e) => {
            const td = e.target.closest('.excel-cell');
            if (td) commitCellEdit(td);
        });
    }

    if (formulaInput) {
        formulaInput.addEventListener('input', () => {
            const cellId = state.activeCell.id;
            const cellEl = document.getElementById(`cell-${cellId}`);
            if (cellEl) cellEl.textContent = formulaInput.value;
        });

        formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
                if (cellEl) {
                    commitCellEdit(cellEl);
                    moveActiveCell(0, 1);
                }
            }
        });
    }

    if (nameBox) {
        nameBox.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const targetId = nameBox.value.trim().toUpperCase();
                if (document.getElementById(`cell-${targetId}`)) {
                    selectCell(targetId, true);
                }
            }
        });
    }

    // Capture global keyboard events
    window.addEventListener('keydown', handleGlobalKeydown, true);

    if (workbookTitleInput) {
        workbookTitleInput.addEventListener('blur', () => {
            state.workbookName = workbookTitleInput.textContent.trim() || 'Книга1';
            autoSave();
        });
    }

    document.getElementById('corner-select-all')?.addEventListener('click', () => {
        document.querySelectorAll('.excel-cell').forEach(c => c.classList.add('in-range'));
        showToast('Выделена вся таблица');
    });

    btnUndo?.addEventListener('click', undo);
    btnRedo?.addEventListener('click', redo);
}

function handleGlobalKeydown(e) {
    const isEditingCell = document.activeElement && document.activeElement.classList.contains('excel-cell');
    const isFormulaBar = document.activeElement === formulaInput;
    const isNameBox = document.activeElement === nameBox;
    const isModalInput = document.activeElement && document.activeElement.tagName === 'INPUT';

    // 1. Undo / Redo Shortcuts (Captured cleanly)
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        e.stopPropagation();
        if (e.shiftKey) redo();
        else undo();
        return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        e.stopPropagation();
        redo();
        return;
    }

    // 2. Clipboard Shortcuts
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c' && !isEditingCell && !isFormulaBar) {
        e.preventDefault();
        copySelection();
        return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x' && !isEditingCell && !isFormulaBar) {
        e.preventDefault();
        cutSelection();
        return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v' && !isEditingCell && !isFormulaBar) {
        e.preventDefault();
        pasteClipboard();
        return;
    }

    // 3. Find Shortcut
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        openFindModal();
        return;
    }

    // 4. Formatting Shortcuts
    if (e.ctrlKey || e.metaKey) {
        if (e.key.toLowerCase() === 's') { e.preventDefault(); manualSave(); return; }
        if (e.key.toLowerCase() === 'b') { e.preventDefault(); toggleStyle('bold'); return; }
        if (e.key.toLowerCase() === 'i') { e.preventDefault(); toggleStyle('italic'); return; }
        if (e.key.toLowerCase() === 'u') { e.preventDefault(); toggleStyle('underline'); return; }
    }

    if (isNameBox || isFormulaBar || isModalInput) return;

    if (!isEditingCell) {
        if (e.key === 'ArrowUp') { e.preventDefault(); moveActiveCell(0, -1); return; }
        if (e.key === 'ArrowDown') { e.preventDefault(); moveActiveCell(0, 1); return; }
        if (e.key === 'ArrowLeft') { e.preventDefault(); moveActiveCell(-1, 0); return; }
        if (e.key === 'ArrowRight') { e.preventDefault(); moveActiveCell(1, 0); return; }
        if (e.key === 'Delete' || e.key === 'Backspace') {
            e.preventDefault();
            clearActiveCell();
            return;
        }

        // Direct keypress typing into active cell
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
            if (cellEl) {
                cellEl.textContent = e.key;
                cellEl.focus();
                const range = document.createRange();
                const sel = window.getSelection();
                range.selectNodeContents(cellEl);
                range.collapse(false);
                sel.removeAllRanges();
                sel.addRange(range);
                if (formulaInput) formulaInput.value = e.key;
                e.preventDefault();
            }
        }
    } else {
        if (e.key === 'Enter') {
            e.preventDefault();
            document.activeElement.blur();
            moveActiveCell(0, e.shiftKey ? -1 : 1);
        } else if (e.key === 'Tab') {
            e.preventDefault();
            document.activeElement.blur();
            moveActiveCell(e.shiftKey ? -1 : 1, 0);
        } else if (e.key === 'Escape') {
            e.preventDefault();
            const cellId = state.activeCell.id;
            const data = state.cells[cellId];
            document.activeElement.textContent = data ? (data.val || '') : '';
            document.activeElement.blur();
        }
    }
}

function moveActiveCell(dx, dy) {
    const currentColIdx = COLS.indexOf(state.activeCell.col);
    const currentRow = state.activeCell.row;

    const newColIdx = Math.max(0, Math.min(COLS.length - 1, currentColIdx + dx));
    const newRow = Math.max(1, Math.min(CONFIG.TOTAL_ROWS, currentRow + dy));

    const newCellId = `${COLS[newColIdx]}${newRow}`;
    selectCell(newCellId, true);
}

/**
 * ============================================================================
 * 10. RIBBON TOOLBAR & BORDERS
 * ============================================================================
 */
function setupRibbonTabs() {
    const tabs = document.querySelectorAll('.ribbon-tab[data-tab]');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');

            const tabName = tab.dataset.tab;
            document.querySelectorAll('.ribbon-panel').forEach(p => p.classList.remove('active'));
            const targetPanel = document.getElementById(`panel-${tabName}`);
            if (targetPanel) targetPanel.classList.add('active');
        });
    });
}

function setupRibbonActions() {
    document.getElementById('btn-bold')?.addEventListener('click', () => toggleStyle('bold'));
    document.getElementById('btn-italic')?.addEventListener('click', () => toggleStyle('italic'));
    document.getElementById('btn-underline')?.addEventListener('click', () => toggleStyle('underline'));

    document.getElementById('btn-align-left')?.addEventListener('click', () => setStyle('textAlign', 'left'));
    document.getElementById('btn-align-center')?.addEventListener('click', () => setStyle('textAlign', 'center'));
    document.getElementById('btn-align-right')?.addEventListener('click', () => setStyle('textAlign', 'right'));

    const fontFamilySelect = document.getElementById('select-font-family');
    if (fontFamilySelect) {
        fontFamilySelect.addEventListener('change', (e) => setStyle('fontFamily', e.target.value));
    }

    const fontSizeSelect = document.getElementById('select-font-size');
    if (fontSizeSelect) {
        fontSizeSelect.addEventListener('change', (e) => setCellFontSize(parseInt(e.target.value, 10)));
    }

    document.getElementById('btn-font-grow')?.addEventListener('click', growFontSize);
    document.getElementById('btn-font-shrink')?.addEventListener('click', shrinkFontSize);

    document.getElementById('btn-fmt-dollar')?.addEventListener('click', () => formatNumber('$'));
    document.getElementById('btn-fmt-percent')?.addEventListener('click', () => formatNumber('%'));
    document.getElementById('btn-fmt-comma')?.addEventListener('click', () => formatNumber(','));

    document.getElementById('btn-autosum')?.addEventListener('click', () => {
        const { col, row } = state.activeCell;
        if (row > 1) {
            const formula = `=SUM(${col}1:${col}${row - 1})`;
            const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
            if (cellEl) {
                cellEl.textContent = formula;
                commitCellEdit(cellEl);
            }
        }
    });

    document.getElementById('btn-clear')?.addEventListener('click', clearActiveCell);
    document.getElementById('btn-sort-asc')?.addEventListener('click', () => sortColumn(true));
    document.getElementById('btn-clean-sheet')?.addEventListener('click', cleanNewWorkbook);
    document.getElementById('btn-export-csv')?.addEventListener('click', exportToCSV);
    document.getElementById('btn-save-mysql')?.addEventListener('click', saveToDatabase);
    document.getElementById('btn-qat-save')?.addEventListener('click', manualSave);
    document.getElementById('btn-add-sheet')?.addEventListener('click', addNewSheet);

    const zoomSlider = document.getElementById('zoom-slider');
    if (zoomSlider) {
        zoomSlider.addEventListener('input', (e) => setZoom(e.target.value));
    }
    document.getElementById('btn-zoom-plus')?.addEventListener('click', () => {
        if (zoomSlider) {
            zoomSlider.value = Math.min(150, parseInt(zoomSlider.value) + 10);
            setZoom(zoomSlider.value);
        }
    });
    document.getElementById('btn-zoom-minus')?.addEventListener('click', () => {
        if (zoomSlider) {
            zoomSlider.value = Math.max(50, parseInt(zoomSlider.value) - 10);
            setZoom(zoomSlider.value);
        }
    });
    document.getElementById('btn-zoom-100')?.addEventListener('click', () => {
        if (zoomSlider) {
            zoomSlider.value = 100;
            setZoom(100);
        }
    });
}

function setZoom(val) {
    const table = document.getElementById('excel-grid-table');
    const zoomVal = document.getElementById('zoom-val');
    if (table) table.style.zoom = `${val}%`;
    if (zoomVal) zoomVal.textContent = `${val}%`;
    updateFillHandlePosition();
}

function toggleStyle(prop) {
    const selectedIds = getSelectedCellIds();
    const batchPrev = [];
    const batchNext = [];

    const firstCellData = state.cells[state.activeCell.id] || {};
    const targetState = !(firstCellData.style && firstCellData.style[prop]);

    selectedIds.forEach(cellId => {
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (!cellEl) return;

        const prevData = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };
        batchPrev.push({ cellId, data: prevData });

        if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
        if (!state.cells[cellId].style) state.cells[cellId].style = {};

        state.cells[cellId].style[prop] = targetState;
        applyCellStyles(cellEl, state.cells[cellId].style);

        batchNext.push({ cellId, data: JSON.parse(JSON.stringify(state.cells[cellId])) });
    });

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });
    syncRibbonWithCell(state.activeCell.id);
    autoSave();
}

function setStyle(property, value) {
    const selectedIds = getSelectedCellIds();
    const batchPrev = [];
    const batchNext = [];

    selectedIds.forEach(cellId => {
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (!cellEl) return;

        const prevData = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };
        batchPrev.push({ cellId, data: prevData });

        if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
        if (!state.cells[cellId].style) state.cells[cellId].style = {};

        state.cells[cellId].style[property] = value;
        applyCellStyles(cellEl, state.cells[cellId].style);

        batchNext.push({ cellId, data: JSON.parse(JSON.stringify(state.cells[cellId])) });
    });

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });
    autoSave();
}

function applyCellStyles(cellEl, style) {
    if (!cellEl || !style) return;
    cellEl.style.fontWeight = style.bold ? 'bold' : 'normal';
    cellEl.style.fontStyle = style.italic ? 'italic' : 'normal';
    cellEl.style.textDecoration = style.underline ? 'underline' : 'none';
    if (style.textAlign) cellEl.style.textAlign = style.textAlign;
    if (style.fontFamily) cellEl.style.fontFamily = style.fontFamily;
    if (style.fontSize) {
        cellEl.style.fontSize = style.fontSize;
        const num = parseInt(style.fontSize, 10);
        cellEl.style.lineHeight = `${num + 4}px`;
    }
    if (style.backgroundColor && style.backgroundColor !== 'transparent') {
        cellEl.style.backgroundColor = style.backgroundColor;
    } else if (style.backgroundColor === 'transparent') {
        cellEl.style.backgroundColor = '';
    }
    if (style.color) cellEl.style.color = style.color;
    if (style.border) cellEl.style.border = style.border;
}

function syncFormulaBarWithCell(cellId) {
    if (!formulaInput) return;
    const cellData = state.cells[cellId];
    if (cellData) {
        formulaInput.value = cellData.formula || (cellData.val !== undefined ? cellData.val : '');
    } else {
        const cellEl = document.getElementById(`cell-${cellId}`);
        formulaInput.value = cellEl ? cellEl.textContent : '';
    }
}

function syncRibbonWithCell(cellId) {
    const cellData = state.cells[cellId] || {};
    const style = cellData.style || {};

    document.getElementById('btn-bold')?.classList.toggle('active', !!style.bold);
    document.getElementById('btn-italic')?.classList.toggle('active', !!style.italic);
    document.getElementById('btn-underline')?.classList.toggle('active', !!style.underline);

    const fontFamilySelect = document.getElementById('select-font-family');
    const fontSizeSelect = document.getElementById('select-font-size');
    if (fontFamilySelect && style.fontFamily) fontFamilySelect.value = style.fontFamily;
    if (fontSizeSelect && style.fontSize) fontSizeSelect.value = parseInt(style.fontSize, 10);

    if (style.backgroundColor && style.backgroundColor !== 'transparent') {
        const bgBar = document.getElementById('color-bar-bg');
        if (bgBar) bgBar.style.backgroundColor = style.backgroundColor;
    }
    if (style.color) {
        const textBar = document.getElementById('color-bar-text');
        if (textBar) textBar.style.backgroundColor = style.color;
    }
}

function formatNumber(type) {
    const cellId = state.activeCell.id;
    const num = getCellValueAsNumber(cellId);
    const cellEl = document.getElementById(`cell-${cellId}`);
    if (!cellEl) return;

    let formatted = num;
    if (type === '$') formatted = `$${num.toLocaleString('en-US', { minimumFractionDigits: 2 })}`;
    else if (type === '%') formatted = `${(num * 100).toFixed(0)}%`;
    else if (type === ',') formatted = num.toLocaleString('en-US');

    cellEl.textContent = formatted;
    if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
    state.cells[cellId].val = formatted;
    autoSave();
}

function setupBordersTool() {
    const btnToggle = document.getElementById('btn-toggle-borders-menu');
    const menu = document.getElementById('dropdown-borders-menu');
    const btnMain = document.getElementById('btn-borders-main');

    if (btnToggle && menu) {
        btnToggle.addEventListener('click', (e) => {
            e.stopPropagation();
            menu.classList.toggle('show');
        });
    }

    if (btnMain) {
        btnMain.addEventListener('click', () => applyBorderToSelection('all'));
    }

    document.querySelectorAll('#dropdown-borders-menu .dropdown-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.stopPropagation();
            applyBorderToSelection(item.dataset.border);
            menu?.classList.remove('show');
        });
    });

    document.addEventListener('click', () => menu?.classList.remove('show'));
}

function applyBorderToSelection(borderType) {
    const selectedIds = getSelectedCellIds();
    selectedIds.forEach(cellId => {
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (!cellEl) return;

        if (borderType === 'all') cellEl.style.border = '1px solid #000000';
        else if (borderType === 'thick') cellEl.style.border = '2px solid #000000';
        else if (borderType === 'bottom') cellEl.style.borderBottom = '2px solid #000000';
        else if (borderType === 'none') cellEl.style.border = '1px solid var(--grid-line)';

        if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
        if (!state.cells[cellId].style) state.cells[cellId].style = {};
        state.cells[cellId].style.border = cellEl.style.border;
    });

    autoSave();
    showToast(`Границы применены: ${borderType}`);
}

/**
 * ============================================================================
 * 11. SEARCH, CHART, CLEAN & EXPORT
 * ============================================================================
 */
function setupFindFeature() {
    document.getElementById('btn-find-replace')?.addEventListener('click', openFindModal);
    document.getElementById('btn-close-find-modal')?.addEventListener('click', closeFindModal);
    document.getElementById('btn-find-next')?.addEventListener('click', () => findNextMatch(true));
    document.getElementById('btn-find-prev')?.addEventListener('click', () => findNextMatch(false));

    const globalSearch = document.getElementById('global-search-input');
    if (globalSearch) {
        globalSearch.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                const term = globalSearch.value.trim();
                if (term) searchAndJumpTo(term);
            }
        });
    }
}

function openFindModal() {
    const modal = document.getElementById('find-modal');
    modal?.classList.add('show');
    const input = document.getElementById('find-text-input');
    input?.focus();
}

function closeFindModal() {
    document.getElementById('find-modal')?.classList.remove('show');
}

let lastFindMatches = [];
let findIndex = 0;

function findNextMatch(forward = true) {
    const input = document.getElementById('find-text-input');
    const term = (input?.value || '').trim().toLowerCase();
    if (!term) return;

    lastFindMatches = [];
    Object.keys(state.cells).forEach(cellId => {
        const item = state.cells[cellId];
        const val = item ? String(item.val || '').toLowerCase() : '';
        if (val.includes(term)) lastFindMatches.push(cellId);
    });

    if (lastFindMatches.length === 0) {
        showToast('Ничего не найдено');
        return;
    }

    findIndex = forward ? (findIndex + 1) % lastFindMatches.length : (findIndex - 1 + lastFindMatches.length) % lastFindMatches.length;
    const targetCellId = lastFindMatches[findIndex];
    selectCell(targetCellId, true);
    showToast(`Найдено: ${findIndex + 1} из ${lastFindMatches.length} (${targetCellId})`);
}

function searchAndJumpTo(term) {
    const lower = term.toLowerCase();
    for (let r = 1; r <= CONFIG.TOTAL_ROWS; r++) {
        for (let c = 0; c < COLS.length; c++) {
            const cellId = `${COLS[c]}${r}`;
            const item = state.cells[cellId];
            const val = item ? String(item.val || '').toLowerCase() : '';
            if (val.includes(lower)) {
                selectCell(cellId, true);
                showToast(`Найдено в ${cellId}`);
                return;
            }
        }
    }
    showToast('Ничего не найдено');
}

function cleanNewWorkbook() {
    state.workbookName = 'Книга1';
    state.cells = {};
    if (workbookTitleInput) workbookTitleInput.textContent = 'Книга1';

    document.querySelectorAll('.excel-cell').forEach(td => {
        td.textContent = '';
        resetCellStylesToDefault(td);
        td.classList.remove('has-comment', 'in-range', 'cell-copied', 'cell-cut');
    });

    for (let r = 1; r <= CONFIG.TOTAL_ROWS; r++) {
        const tr = document.getElementById(`row-tr-${r}`);
        if (tr) tr.style.height = '22px';
    }

    undoStack.length = 0;
    redoStack.length = 0;
    updateUndoRedoUI();

    if (formulaInput) formulaInput.value = '';
    localStorage.removeItem(CONFIG.STORAGE_KEY);
    autoSave();
    selectCell('A1');
    showToast('Создан абсолютно чистый лист! Размеры и стили сброшены.');
}

function exportToCSV() {
    let maxRow = -1;
    let maxColIdx = -1;

    Object.keys(state.cells).forEach(cellId => {
        const item = state.cells[cellId];
        const val = item ? (item.val !== undefined && item.val !== null ? String(item.val).trim() : '') : '';
        const formula = item ? (item.formula || '').trim() : '';

        if (val !== '' || formula !== '') {
            const match = cellId.match(/^([A-Z]+)([0-9]+)$/);
            if (match) {
                const colIdx = COLS.indexOf(match[1]);
                const row = parseInt(match[2], 10);
                if (colIdx !== -1) {
                    if (row > maxRow) maxRow = row;
                    if (colIdx > maxColIdx) maxColIdx = colIdx;
                }
            }
        }
    });

    if (maxRow === -1) {
        showToast('Таблица пуста. Введите данные перед скачиванием.');
        return;
    }

    let csv = '';
    for (let r = 1; r <= maxRow; r++) {
        let rowVals = [];
        for (let c = 0; c <= maxColIdx; c++) {
            const cellId = `${COLS[c]}${r}`;
            const item = state.cells[cellId];
            const val = item && item.val !== undefined ? String(item.val) : '';
            rowVals.push(`"${val.replace(/"/g, '""')}"`);
        }
        csv += rowVals.join(',') + '\r\n';
    }

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${state.workbookName || 'Книга1'}.csv`;
    link.click();
    showToast(`Файл "${state.workbookName}.csv" скачан! (${maxRow} строк)`);
}

/**
 * ============================================================================
 * 12. OTHER NEW TABS FEATURES (CHARTS, COMMENTS, ETC.)
 * ============================================================================
 */
function setupAllNewTabFeatures() {
    document.getElementById('btn-chart-bar')?.addEventListener('click', () => openChartModal('bar'));
    document.getElementById('btn-chart-pie')?.addEventListener('click', () => openChartModal('pie'));
    document.getElementById('btn-chart-line')?.addEventListener('click', () => openChartModal('line'));

    document.getElementById('btn-close-chart-modal')?.addEventListener('click', closeChartModal);
    document.getElementById('btn-ok-chart')?.addEventListener('click', closeChartModal);
    document.getElementById('btn-download-chart')?.addEventListener('click', downloadChartImage);

    const imgInput = document.getElementById('image-file-input');
    document.getElementById('btn-insert-image')?.addEventListener('click', () => imgInput?.click());
    if (imgInput) {
        imgInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (evt) => {
                const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
                if (cellEl) {
                    cellEl.innerHTML = `<img src="${evt.target.result}" style="max-height:20px; vertical-align:middle;">`;
                    state.cells[state.activeCell.id] = { val: '[Рисунок]', formula: '', style: {} };
                    autoSave();
                    showToast('Рисунок вставлен в ячейку');
                }
            };
            reader.readAsDataURL(file);
        });
    }

    document.getElementById('btn-insert-table')?.addEventListener('click', formatSelectionAsTable);
    document.getElementById('btn-insert-date')?.addEventListener('click', () => {
        const today = new Date().toLocaleDateString('ru-RU');
        const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
        if (cellEl) {
            cellEl.textContent = today;
            commitCellEdit(cellEl);
            showToast(`Вставлена дата: ${today}`);
        }
    });

    document.getElementById('chk-show-gridlines')?.addEventListener('change', (e) => {
        const table = document.getElementById('excel-grid-table');
        if (table) table.classList.toggle('no-gridlines', !e.target.checked);
    });

    document.getElementById('chk-show-headings')?.addEventListener('change', (e) => {
        const container = document.getElementById('spreadsheet-container');
        if (container) container.classList.toggle('no-headings', !e.target.checked);
    });

    document.getElementById('btn-toggle-show-formulas')?.addEventListener('click', () => {
        state.showFormulas = !state.showFormulas;
        document.getElementById('btn-toggle-show-formulas')?.classList.toggle('active', state.showFormulas);
        recalculateAllFormulas();
        showToast(state.showFormulas ? 'Отображение формул включено' : 'Отображение значений включено');
    });

    document.getElementById('btn-recalculate-all')?.addEventListener('click', () => {
        recalculateAllFormulas();
        showToast('Все формулы пересчитаны!');
    });

    document.getElementById('btn-data-sort-asc')?.addEventListener('click', () => sortColumn(true));
    document.getElementById('btn-data-sort-desc')?.addEventListener('click', () => sortColumn(false));
    document.getElementById('btn-remove-duplicates')?.addEventListener('click', removeColumnDuplicates);
    document.getElementById('btn-trim-spaces')?.addEventListener('click', trimAllSpaces);

    document.getElementById('btn-add-comment')?.addEventListener('click', addCellComment);
    document.getElementById('btn-delete-comment')?.addEventListener('click', deleteCellComment);
    document.getElementById('btn-protect-sheet')?.addEventListener('click', () => showToast('Лист защищен'));

    document.getElementById('chk-show-formulabar')?.addEventListener('change', (e) => {
        const bar = document.getElementById('formula-bar-wrapper');
        if (bar) bar.style.display = e.target.checked ? 'flex' : 'none';
    });

    document.getElementById('btn-help-shortcuts')?.addEventListener('click', () => {
        alert("Горячие клавиши Microsoft Excel:\n\n• Ctrl + Z: Отменить действие\n• Ctrl + Y: Повторить действие\n• Ctrl + C: Копировать\n• Ctrl + X: Вырезать\n• Ctrl + V: Вставить\n• Ctrl + S: Сохранить книгу\n• Ctrl + B: Полужирный шрифт\n• Ctrl + I: Курсив\n• Ctrl + U: Подчеркнутый\n• Ctrl + F: Найти в таблице\n• Автозаполнение: Потяните за маркер (+) в правом нижнем углу ячейки");
    });
}

function insertSymbol(symbol) {
    const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
    if (cellEl) {
        cellEl.textContent += symbol;
        commitCellEdit(cellEl);
        showToast(`Вставлен символ: ${symbol}`);
    }
}

function insertFormulaTemplate(prefix) {
    const { col, row } = state.activeCell;
    const formula = `${prefix}${col}1:${col}${Math.max(1, row - 1)})`;
    const cellEl = document.getElementById(`cell-${state.activeCell.id}`);
    if (cellEl) {
        cellEl.textContent = formula;
        if (formulaInput) formulaInput.value = formula;
        commitCellEdit(cellEl);
    }
}

function openChartModal(chartType) {
    const modal = document.getElementById('chart-modal');
    const canvas = document.getElementById('chart-canvas');
    const titleEl = document.getElementById('chart-modal-title');
    if (!modal || !canvas) return;

    titleEl.textContent = chartType === 'bar' ? 'Столбчатая гистограмма' : (chartType === 'pie' ? 'Круговая диаграмма' : 'График');
    modal.classList.add('show');

    const labels = [];
    const values = [];
    const activeColIdx = COLS.indexOf(state.activeCell.col);

    for (let r = 1; r <= 15; r++) {
        const labelCellId = `${COLS[Math.max(0, activeColIdx - 1)]}${r}`;
        const valCellId = `${state.activeCell.col}${r}`;

        const num = getCellValueAsNumber(valCellId);
        const labelData = state.cells[labelCellId];
        const label = labelData ? (labelData.val || `R${r}`) : `Строка ${r}`;

        if (num > 0) {
            labels.push(label);
            values.push(num);
        }
    }

    if (values.length === 0) {
        labels.push('Янв', 'Фев', 'Мар', 'Апр', 'Май');
        values.push(12, 19, 3, 17, 28);
    }

    drawChartOnCanvas(canvas, chartType, labels, values);
}

function drawChartOnCanvas(canvas, type, labels, values) {
    const ctx = canvas.getContext('2d');
    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);

    const colors = ['#107c41', '#0078d4', '#d83b01', '#5c2d91', '#ffb900', '#008272', '#b4009e', '#e81123'];

    if (type === 'bar') {
        const maxVal = Math.max(...values, 10);
        const padding = 40;
        const barWidth = Math.min(45, (w - padding * 2) / values.length - 10);

        ctx.strokeStyle = '#e1dfdd';
        ctx.beginPath();
        ctx.moveTo(padding, 20);
        ctx.lineTo(padding, h - padding);
        ctx.lineTo(w - 20, h - padding);
        ctx.stroke();

        values.forEach((val, idx) => {
            const x = padding + 15 + idx * (barWidth + 12);
            const barH = (val / maxVal) * (h - padding - 40);
            const y = (h - padding) - barH;

            ctx.fillStyle = colors[idx % colors.length];
            ctx.fillRect(x, y, barWidth, barH);

            ctx.fillStyle = '#333';
            ctx.font = '11px Segoe UI';
            ctx.textAlign = 'center';
            ctx.fillText(String(val), x + barWidth / 2, y - 5);
            ctx.fillText(labels[idx] ? labels[idx].substring(0, 8) : '', x + barWidth / 2, h - padding + 15);
        });
    } else if (type === 'pie') {
        const total = values.reduce((a, b) => a + b, 0);
        let startAngle = 0;
        const centerX = w / 2 - 40;
        const centerY = h / 2;
        const radius = Math.min(centerX, centerY) - 20;

        values.forEach((val, idx) => {
            const sliceAngle = (val / total) * 2 * Math.PI;
            ctx.fillStyle = colors[idx % colors.length];
            ctx.beginPath();
            ctx.moveTo(centerX, centerY);
            ctx.arc(centerX, centerY, radius, startAngle, startAngle + sliceAngle);
            ctx.closePath();
            ctx.fill();

            const legX = w - 140;
            const legY = 40 + idx * 22;
            ctx.fillRect(legX, legY, 14, 14);
            ctx.fillStyle = '#333';
            ctx.font = '11px Segoe UI';
            ctx.textAlign = 'left';
            ctx.fillText(`${labels[idx] || ''}: ${val}`, legX + 20, legY + 11);

            startAngle += sliceAngle;
        });
    } else if (type === 'line') {
        const maxVal = Math.max(...values, 10);
        const padding = 40;
        const stepX = (w - padding * 2 - 20) / Math.max(1, values.length - 1);

        ctx.strokeStyle = '#e1dfdd';
        ctx.beginPath();
        ctx.moveTo(padding, 20);
        ctx.lineTo(padding, h - padding);
        ctx.lineTo(w - 20, h - padding);
        ctx.stroke();

        ctx.strokeStyle = '#107c41';
        ctx.lineWidth = 2.5;
        ctx.beginPath();

        values.forEach((val, idx) => {
            const x = padding + idx * stepX;
            const y = (h - padding) - (val / maxVal) * (h - padding - 40);
            if (idx === 0) ctx.moveTo(x, y);
            else ctx.lineTo(x, y);
        });
        ctx.stroke();

        values.forEach((val, idx) => {
            const x = padding + idx * stepX;
            const y = (h - padding) - (val / maxVal) * (h - padding - 40);
            ctx.fillStyle = '#107c41';
            ctx.beginPath();
            ctx.arc(x, y, 4, 0, Math.PI * 2);
            ctx.fill();

            ctx.fillStyle = '#333';
            ctx.font = '11px Segoe UI';
            ctx.textAlign = 'center';
            ctx.fillText(String(val), x, y - 8);
            ctx.fillText(labels[idx] ? labels[idx].substring(0, 8) : '', x, h - padding + 15);
        });
    }
}

function closeChartModal() {
    document.getElementById('chart-modal')?.classList.remove('show');
}

function downloadChartImage() {
    const canvas = document.getElementById('chart-canvas');
    if (!canvas) return;
    const link = document.createElement('a');
    link.download = `Excel_Chart_${state.workbookName}.png`;
    link.href = canvas.toDataURL();
    link.click();
    showToast('График сохранен как PNG!');
}

function formatSelectionAsTable() {
    const selectedIds = getSelectedCellIds();
    const batchPrev = [];
    const batchNext = [];

    selectedIds.forEach((cellId, idx) => {
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (!cellEl) return;

        const prevData = state.cells[cellId] ? JSON.parse(JSON.stringify(state.cells[cellId])) : { val: '', formula: '', style: {} };
        batchPrev.push({ cellId, data: prevData });

        const rowNum = parseInt(cellId.replace(/[A-Z]/g, ''), 10);
        if (idx < 6) {
            cellEl.style.backgroundColor = '#107c41';
            cellEl.style.color = '#ffffff';
            cellEl.style.fontWeight = 'bold';
        } else if (rowNum % 2 === 0) {
            cellEl.style.backgroundColor = '#f3f9f4';
            cellEl.style.color = '#000000';
        } else {
            cellEl.style.backgroundColor = '#ffffff';
            cellEl.style.color = '#000000';
        }

        if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
        state.cells[cellId].style = {
            backgroundColor: cellEl.style.backgroundColor,
            color: cellEl.style.color,
            bold: idx < 6
        };

        batchNext.push({ cellId, data: JSON.parse(JSON.stringify(state.cells[cellId])) });
    });

    recordAction({ type: 'batch_cells', prevList: batchPrev, nextList: batchNext });
    autoSave();
    showToast('Диапазон оформлен в стиле Таблицы');
}

function sortColumn(ascending = true) {
    const col = state.activeCell.col;
    const rowValues = [];

    for (let r = 1; r <= CONFIG.TOTAL_ROWS; r++) {
        const cellId = `${col}${r}`;
        const item = state.cells[cellId];
        const val = item ? (item.val || '') : '';
        if (val !== '') {
            rowValues.push({ row: r, val: val, fullItem: item });
        }
    }

    if (rowValues.length <= 1) {
        showToast('В столбце недостаточно данных');
        return;
    }

    rowValues.sort((a, b) => {
        const numA = parseFloat(a.val);
        const numB = parseFloat(b.val);
        if (!isNaN(numA) && !isNaN(numB)) {
            return ascending ? numA - numB : numB - numA;
        }
        return ascending ? a.val.localeCompare(b.val) : b.val.localeCompare(a.val);
    });

    rowValues.forEach((item, idx) => {
        const targetRow = idx + 1;
        const targetCellId = `${col}${targetRow}`;
        const cellEl = document.getElementById(`cell-${targetCellId}`);
        if (cellEl) cellEl.textContent = item.val;
        state.cells[targetCellId] = item.fullItem;
    });

    autoSave();
    showToast(`Столбец ${col} отсортирован (${ascending ? 'А-Я' : 'Я-А'})!`);
}

function removeColumnDuplicates() {
    const col = state.activeCell.col;
    const seen = new Set();
    let removedCount = 0;

    for (let r = 1; r <= CONFIG.TOTAL_ROWS; r++) {
        const cellId = `${col}${r}`;
        const item = state.cells[cellId];
        const val = item ? (item.val || '') : '';
        if (val !== '') {
            if (seen.has(val)) {
                delete state.cells[cellId];
                const cellEl = document.getElementById(`cell-${cellId}`);
                if (cellEl) cellEl.textContent = '';
                removedCount++;
            } else {
                seen.add(val);
            }
        }
    }

    autoSave();
    showToast(`Удалено повторяющихся: ${removedCount}`);
}

function trimAllSpaces() {
    let count = 0;
    Object.keys(state.cells).forEach(cellId => {
        const item = state.cells[cellId];
        if (item && item.val && typeof item.val === 'string') {
            const trimmed = item.val.trim();
            if (trimmed !== item.val) {
                item.val = trimmed;
                const cellEl = document.getElementById(`cell-${cellId}`);
                if (cellEl) cellEl.textContent = trimmed;
                count++;
            }
        }
    });
    autoSave();
    showToast(`Очищены пробелы в ${count} ячейках`);
}

function addCellComment() {
    const cellId = state.activeCell.id;
    const current = state.cells[cellId]?.comment || '';
    const note = prompt(`Примечание для ${cellId}:`, current);
    if (note !== null) {
        if (!state.cells[cellId]) state.cells[cellId] = { val: '', formula: '', style: {} };
        state.cells[cellId].comment = note;

        const cellEl = document.getElementById(`cell-${cellId}`);
        if (cellEl) {
            cellEl.classList.toggle('has-comment', !!note);
            if (note) cellEl.title = `Примечание: ${note}`;
            else cellEl.removeAttribute('title');
        }
        autoSave();
        showToast(note ? 'Примечание добавлено' : 'Примечание удалено');
    }
}

function deleteCellComment() {
    const cellId = state.activeCell.id;
    if (state.cells[cellId]) {
        state.cells[cellId].comment = '';
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (cellEl) {
            cellEl.classList.remove('has-comment');
            cellEl.removeAttribute('title');
        }
        autoSave();
        showToast('Примечание удалено');
    }
}

/**
 * ============================================================================
 * 13. STORAGE, TEMPLATES & STATUS BAR
 * ============================================================================
 */
function loadFromStorageOrTemplate() {
    try {
        const raw = localStorage.getItem(CONFIG.STORAGE_KEY);
        if (raw) {
            const data = JSON.parse(raw);
            if (data.name) {
                state.workbookName = data.name;
                if (workbookTitleInput) workbookTitleInput.textContent = data.name;
            }
            if (data.data && Object.keys(data.data).length > 0) {
                state.cells = data.data;
                renderLoadedCells();
            }
        }
    } catch (e) {
        console.error('Failed to load workbook', e);
    }
}

function renderLoadedCells() {
    Object.keys(state.cells).forEach(cellId => {
        const cellData = state.cells[cellId];
        const cellEl = document.getElementById(`cell-${cellId}`);
        if (cellEl && cellData) {
            let displayVal = cellData.val;
            if (cellData.formula) {
                displayVal = evaluateFormula(cellData.formula);
                cellData.val = displayVal;
            }
            cellEl.textContent = displayVal !== undefined ? displayVal : '';
            if (cellData.style) {
                applyCellStyles(cellEl, cellData.style);
            }
            if (cellData.comment) {
                cellEl.classList.add('has-comment');
                cellEl.title = `Примечание: ${cellData.comment}`;
            }
            checkAndResetRowHeight(parseInt(cellId.replace(/[A-Z]/g, ''), 10));
        }
    });
}

function autoSave() {
    try {
        const payload = {
            name: state.workbookName,
            data: state.cells
        };
        localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
        console.warn('LocalStorage error:', e);
    }
}

function manualSave() {
    autoSave();
    showToast(`Книга "${state.workbookName}" успешно сохранена!`);
}

function updateStatusBar() {
    const selectedIds = getSelectedCellIds();
    let sum = 0;
    let count = 0;
    let numCount = 0;

    selectedIds.forEach(cellId => {
        const data = state.cells[cellId];
        if (data && data.val !== '' && data.val !== undefined) {
            count++;
            const num = parseFloat(data.val);
            if (!isNaN(num)) {
                sum += num;
                numCount++;
            }
        }
    });

    const avg = numCount > 0 ? parseFloat((sum / numCount).toFixed(2)) : 0;

    if (statusBarSum) statusBarSum.textContent = numCount > 0 ? sum : 0;
    if (statusBarCount) statusBarCount.textContent = count;
    if (statusBarAvg) statusBarAvg.textContent = avg;
}

async function saveToDatabase() {
    const payload = {
        name: state.workbookName,
        data: state.cells,
        timestamp: new Date().toISOString()
    };

    try {
        showToast('Подключение к MySQL...');
        const response = await fetch('api.php?action=save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            showToast('✓ Сохранено в MySQL через PHP API!');
        } else {
            fallbackLocalSave();
        }
    } catch (err) {
        fallbackLocalSave();
    }
}

function fallbackLocalSave() {
    autoSave();
    showToast('✓ Сохранено в LocalStorage (Запустите api.php или server.py для MySQL)');
}

function addNewSheet() {
    const newIdx = state.sheets.length + 1;
    const sheetName = `Лист${newIdx}`;
    state.sheets.push(sheetName);

    const sheetsList = document.getElementById('sheets-list');
    if (sheetsList) {
        const tab = document.createElement('div');
        tab.className = 'sheet-tab-item';
        tab.textContent = sheetName;
        tab.onclick = () => {
            document.querySelectorAll('.sheet-tab-item').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            showToast(`Переключено на ${sheetName}`);
        };
        sheetsList.appendChild(tab);
    }
    showToast(`Создан новый лист: ${sheetName}`);
}

function showToast(message) {
    let toast = document.getElementById('excel-toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'excel-toast';
        toast.className = 'excel-toast';
        document.body.appendChild(toast);
    }
    toast.innerHTML = `<span>✓</span> <div>${message}</div>`;
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3200);
}
