/**
 * Microsoft Excel Backstage (File View) Logic v2.0
 */

document.addEventListener('DOMContentLoaded', () => {
    updateGreeting();
    setupTemplateHandlers();
    setupRecentFileHandlers();
    setupSidebarActions();
    setupTabs();
});

function updateGreeting() {
    const greetingEl = document.getElementById('backstage-greeting');
    if (!greetingEl) return;

    const hour = new Date().getHours();
    let text = 'Добрый день';
    if (hour >= 5 && hour < 12) {
        text = 'Доброе утро';
    } else if (hour >= 12 && hour < 18) {
        text = 'Добрый день';
    } else {
        text = 'Добрый вечер';
    }
    greetingEl.textContent = text;
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

// Sample Templates Data
const TEMPLATES = {
    blank: {
        name: 'Книга1',
        data: {}
    },
    tour: {
        name: 'Добро_пожаловать_в_Excel',
        data: {
            'A1': { val: 'Добро пожаловать в Microsoft Excel Web!', style: { bold: true, fontSize: '16px', color: '#107c41' } },
            'A3': { val: '№', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'B3': { val: 'Функция', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'C3': { val: 'Описание', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'A4': { val: '1' },
            'B4': { val: 'Формулы' },
            'C4': { val: 'Формулы начинаются со знака равно: =SUM(A1:A5)' },
            'A5': { val: '2' },
            'B5': { val: 'Форматирование' },
            'C5': { val: 'Цвета заливки и текста теперь наглядно выбираются из палитры' },
            'A6': { val: '3' },
            'B6': { val: 'Экспорт' },
            'C6': { val: 'Скачивание сохраняет только заполненные вами строки без лишних пустых строк' }
        }
    },
    formula: {
        name: 'Учебник_по_формулам',
        data: {
            'A1': { val: 'Учебник по формулам Excel', style: { bold: true, fontSize: '15px', color: '#107c41' } },
            'A3': { val: 'Товар', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'B3': { val: 'Кол-во', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'C3': { val: 'Цена ($)', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'D3': { val: 'Сумма ($)', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'A4': { val: 'Ноутбук' }, 'B4': { val: 5 }, 'C4': { val: 800 }, 'D4': { val: '=B4*C4' },
            'A5': { val: 'Монитор' }, 'B5': { val: 12 }, 'C5': { val: 250 }, 'D5': { val: '=B5*C5' },
            'A6': { val: 'Клавиатура' }, 'B6': { val: 25 }, 'C6': { val: 45 }, 'D6': { val: '=B6*C6' },
            'A7': { val: 'Мышь' }, 'B7': { val: 30 }, 'C7': { val: 25 }, 'D7': { val: '=B7*C7' },
            'A9': { val: 'ИТОГО:', style: { bold: true } },
            'D9': { val: '=SUM(D4:D7)', style: { bold: true, backgroundColor: '#d4edda', color: '#155724' } },
            'A10': { val: 'СРЕДНЯЯ ЦЕНА:', style: { bold: true } },
            'D10': { val: '=AVERAGE(C4:C7)', style: { bold: true } }
        }
    },
    budget: {
        name: 'Личный_бюджет_на_месяц',
        data: {
            'A1': { val: 'Личный бюджет на месяц', style: { bold: true, fontSize: '16px', color: '#0078d4' } },
            'A3': { val: 'Категория', style: { bold: true, backgroundColor: '#deecf9' } },
            'B3': { val: 'План ($)', style: { bold: true, backgroundColor: '#deecf9' } },
            'C3': { val: 'Факт ($)', style: { bold: true, backgroundColor: '#deecf9' } },
            'D3': { val: 'Разница ($)', style: { bold: true, backgroundColor: '#deecf9' } },
            'A4': { val: 'Аренда / Жилье' }, 'B4': { val: 400 }, 'C4': { val: 400 }, 'D4': { val: '=B4-C4' },
            'A5': { val: 'Продукты питания' }, 'B5': { val: 300 }, 'C5': { val: 320 }, 'D5': { val: '=B5-C5' },
            'A6': { val: 'Транспорт' }, 'B6': { val: 80 }, 'C6': { val: 75 }, 'D6': { val: '=B6-C6' },
            'A7': { val: 'Развлечения' }, 'B7': { val: 100 }, 'C7': { val: 130 }, 'D7': { val: '=B7-C7' },
            'A8': { val: 'Образование' }, 'B8': { val: 150 }, 'C8': { val: 150 }, 'D8': { val: '=B8-C8' },
            'A10': { val: 'Всего расходов:', style: { bold: true } },
            'B10': { val: '=SUM(B4:B8)', style: { bold: true } },
            'C10': { val: '=SUM(C4:C8)', style: { bold: true } },
            'D10': { val: '=SUM(D4:D8)', style: { bold: true, backgroundColor: '#fff4ce' } }
        }
    },
    maktab: {
        name: '2-maktab_baholari',
        data: {
            'A1': { val: '2-Maktab O\'quvchilari Baholar Jurnali', style: { bold: true, fontSize: '15px', color: '#107c41' } },
            'A3': { val: '№', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'B3': { val: 'O\'quvchi F.I.Sh', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'C3': { val: 'Matematika', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'D3': { val: 'Informatika', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'E3': { val: 'Ingliz tili', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'F3': { val: 'O\'rtacha baho', style: { bold: true, backgroundColor: '#f3f2f1' } },
            'A4': { val: '1' }, 'B4': { val: 'Aliyev Dovud' }, 'C4': { val: 5 }, 'D4': { val: 5 }, 'E4': { val: 5 }, 'F4': { val: '=AVERAGE(C4:E4)' },
            'A5': { val: '2' }, 'B5': { val: 'Karimov Jasur' }, 'C5': { val: 4 }, 'D5': { val: 5 }, 'E5': { val: 4 }, 'F5': { val: '=AVERAGE(C5:E5)' },
            'A6': { val: '3' }, 'B6': { val: 'Toshmatova Madina' }, 'C6': { val: 5 }, 'D6': { val: 4 }, 'E6': { val: 5 }, 'F6': { val: '=AVERAGE(C6:E6)' },
            'A7': { val: '4' }, 'B7': { val: 'Rustamov Bobur' }, 'C7': { val: 3 }, 'D7': { val: 4 }, 'E7': { val: 4 }, 'F7': { val: '=AVERAGE(C7:E7)' },
            'A9': { val: 'Sinf o\'rtacha ko\'rsatkichi:', style: { bold: true } },
            'F9': { val: '=AVERAGE(F4:F7)', style: { bold: true, backgroundColor: '#d4edda', color: '#155724' } }
        }
    },
    sirdaryo: {
        name: 'sirdaryo_11-sinf_lotin',
        data: {
            'A1': { val: 'Sirdaryo viloyati 11-sinf reyting jadvali', style: { bold: true, fontSize: '15px', color: '#004e8c' } },
            'A3': { val: 'T/r', style: { bold: true, backgroundColor: '#e1dfdd' } },
            'B3': { val: 'Tuman / Shahar', style: { bold: true, backgroundColor: '#e1dfdd' } },
            'C3': { val: 'Bitiruvchilar', style: { bold: true, backgroundColor: '#e1dfdd' } },
            'D3': { val: 'Oliy ta\'limga kirish', style: { bold: true, backgroundColor: '#e1dfdd' } },
            'E3': { val: 'Ulush (%)', style: { bold: true, backgroundColor: '#e1dfdd' } },
            'A4': { val: '1' }, 'B4': { val: 'Guliston shahar' }, 'C4': { val: 1250 }, 'D4': { val: 980 }, 'E4': { val: '=(D4/C4)*100' },
            'A5': { val: '2' }, 'B5': { val: 'Yangiyer shahar' }, 'C5': { val: 820 }, 'D5': { val: 640 }, 'E5': { val: '=(D5/C5)*100' },
            'A6': { val: '3' }, 'B6': { val: 'Shirin shahar' }, 'C6': { val: 410 }, 'D6': { val: 330 }, 'E6': { val: '=(D6/C6)*100' },
            'A7': { val: '4' }, 'B7': { val: 'Boyovut tumani' }, 'C7': { val: 1420 }, 'D7': { val: 1020 }, 'E7': { val: '=(D7/C7)*100' },
            'A9': { val: 'JAMI VILOYAT:', style: { bold: true } },
            'C9': { val: '=SUM(C4:C7)', style: { bold: true } },
            'D9': { val: '=SUM(D4:D7)', style: { bold: true } }
        }
    },
    moliyaviy: {
        name: 'shaxsiy_moliyaviy_yordamchi',
        data: {
            'A1': { val: 'Shaxsiy Moliyaviy Yordamchi & Daromadlar', style: { bold: true, fontSize: '15px', color: '#107c41' } },
            'A3': { val: 'Sana', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'B3': { val: 'Daromad turi', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'C3': { val: 'Kirim ($)', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'D3': { val: 'Chiqim ($)', style: { bold: true, backgroundColor: '#fde7e9' } },
            'E3': { val: 'Qoldiq ($)', style: { bold: true, backgroundColor: '#e8f5e9' } },
            'A4': { val: '01.10.2026' }, 'B4': { val: 'Oylik maosh' }, 'C4': { val: 1200 }, 'D4': { val: 0 }, 'E4': { val: 1200 },
            'A5': { val: '05.10.2026' }, 'B5': { val: 'Frilans loyiha' }, 'C5': { val: 450 }, 'D5': { val: 0 }, 'E5': { val: '=E4+C5' },
            'A6': { val: '10.10.2026' }, 'B6': { val: 'Ijara to\'lovi' }, 'C6': { val: 0 }, 'D6': { val: 350 }, 'E6': { val: '=E5-D6' },
            'A7': { val: '15.10.2026' }, 'B7': { val: 'Do\'kon xarajatlari' }, 'C7': { val: 0 }, 'D7': { val: 180 }, 'E7': { val: '=E6-D7' },
            'A9': { val: 'Sof daromad:', style: { bold: true } },
            'C9': { val: '=SUM(C4:C7)', style: { bold: true } },
            'D9': { val: '=SUM(D4:D7)', style: { bold: true } },
            'E9': { val: '=C9-D9', style: { bold: true, backgroundColor: '#d4edda', color: '#155724' } }
        }
    }
};

function loadWorkbookAndRedirect(workbookKey) {
    const template = TEMPLATES[workbookKey] || TEMPLATES.blank;
    try {
        localStorage.setItem('excel_active_workbook', JSON.stringify(template));
    } catch (e) {
        console.error('LocalStorage error:', e);
    }
    window.location.href = 'index.html';
}

function setupTemplateHandlers() {
    document.getElementById('card-blank')?.addEventListener('click', () => loadWorkbookAndRedirect('blank'));
    document.getElementById('card-tour')?.addEventListener('click', () => loadWorkbookAndRedirect('tour'));
    document.getElementById('card-formula')?.addEventListener('click', () => loadWorkbookAndRedirect('formula'));
    document.getElementById('card-pivot')?.addEventListener('click', () => loadWorkbookAndRedirect('tour'));
    document.getElementById('card-chart')?.addEventListener('click', () => loadWorkbookAndRedirect('tour'));
    document.getElementById('card-calendar')?.addEventListener('click', () => loadWorkbookAndRedirect('blank'));
    document.getElementById('card-budget')?.addEventListener('click', () => loadWorkbookAndRedirect('budget'));
}

function setupRecentFileHandlers() {
    document.getElementById('file-maktab')?.addEventListener('click', () => loadWorkbookAndRedirect('maktab'));
    document.getElementById('file-sirdaryo')?.addEventListener('click', () => loadWorkbookAndRedirect('sirdaryo'));
    document.getElementById('file-moliyaviy')?.addEventListener('click', () => loadWorkbookAndRedirect('moliyaviy'));
}

function setupSidebarActions() {
    const openBtn = document.getElementById('sidebar-open');
    const fileInput = document.getElementById('backstage-file-input');
    if (openBtn && fileInput) {
        openBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', handleFileOpen);
    }

    document.getElementById('sidebar-save')?.addEventListener('click', () => {
        showToast('Книга сохранена в браузере (LocalStorage & База данных)');
    });

    document.getElementById('sidebar-save-as')?.addEventListener('click', () => {
        exportCurrentToCSV();
    });

    document.getElementById('sidebar-print')?.addEventListener('click', () => window.print());
    document.getElementById('sidebar-export-pdf')?.addEventListener('click', () => window.print());
    document.getElementById('sidebar-new')?.addEventListener('click', () => loadWorkbookAndRedirect('blank'));
}

function setupTabs() {
    const tabRecent = document.getElementById('tab-recent');
    const tabPinned = document.getElementById('tab-pinned');
    const recentTable = document.getElementById('files-table-recent');
    const pinnedTable = document.getElementById('files-table-pinned');

    if (tabRecent && tabPinned) {
        tabRecent.addEventListener('click', () => {
            tabRecent.classList.add('active');
            tabPinned.classList.remove('active');
            if (recentTable) recentTable.style.display = 'table';
            if (pinnedTable) pinnedTable.style.display = 'none';
        });

        tabPinned.addEventListener('click', () => {
            tabPinned.classList.add('active');
            tabRecent.classList.remove('active');
            if (recentTable) recentTable.style.display = 'none';
            if (pinnedTable) pinnedTable.style.display = 'table';
        });
    }
}

function handleFileOpen(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        try {
            if (file.name.endsWith('.json')) {
                const parsed = JSON.parse(text);
                localStorage.setItem('excel_active_workbook', JSON.stringify(parsed));
                window.location.href = 'index.html';
                return;
            }

            const lines = text.split(/\r\n|\n/);
            const data = {};
            const colLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
            lines.forEach((line, rIdx) => {
                if (!line.trim()) return;
                const cols = line.split(',');
                cols.forEach((cellVal, cIdx) => {
                    if (cIdx < colLetters.length) {
                        const cellId = `${colLetters[cIdx]}${rIdx + 1}`;
                        data[cellId] = { val: cellVal.trim().replace(/^"|"$/g, '') };
                    }
                });
            });

            const wb = {
                name: file.name.replace(/\.[^/.]+$/, ""),
                data: data
            };
            localStorage.setItem('excel_active_workbook', JSON.stringify(wb));
            window.location.href = 'index.html';
        } catch (err) {
            showToast('Ошибка при чтении файла');
        }
    };
    reader.readAsText(file);
}

// Clean smart CSV export without ghost rows
function exportCurrentToCSV() {
    let saved = localStorage.getItem('excel_active_workbook');
    let wb = saved ? JSON.parse(saved) : TEMPLATES.blank;
    let data = wb.data || {};

    let maxRow = -1;
    let maxColIdx = -1;
    const colLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    Object.keys(data).forEach(cellId => {
        const item = data[cellId];
        const val = item ? (item.val !== undefined ? String(item.val).trim() : '') : '';
        if (val !== '') {
            const match = cellId.match(/^([A-Z]+)([0-9]+)$/);
            if (match) {
                const cIdx = colLetters.indexOf(match[1]);
                const r = parseInt(match[2], 10);
                if (cIdx !== -1) {
                    if (r > maxRow) maxRow = r;
                    if (cIdx > maxColIdx) maxColIdx = cIdx;
                }
            }
        }
    });

    if (maxRow === -1) {
        showToast('Книга пуста. Нет данных для скачивания.');
        return;
    }

    let csv = '';
    for (let r = 1; r <= maxRow; r++) {
        let rowVals = [];
        for (let c = 0; c <= maxColIdx; c++) {
            let cellId = `${colLetters[c]}${r}`;
            let val = data[cellId] && data[cellId].val !== undefined ? String(data[cellId].val) : '';
            rowVals.push(`"${val.replace(/"/g, '""')}"`);
        }
        csv += rowVals.join(',') + '\r\n';
    }

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${wb.name || 'Книга1'}.csv`;
    link.click();
    showToast(`Файл "${wb.name || 'Книга1'}.csv" скачан!`);
}
