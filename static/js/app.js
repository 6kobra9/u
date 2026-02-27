// Глобальні змінні
let currentTable = null;
let currentTableInfo = null;
let currentMenuItem = null;
let editingRow = null;

// Ініціалізація при завантаженні сторінки
document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
    
    const exportBtn = document.getElementById('exportExcelBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', exportTableToExcel);
    }
});


function setupEventListeners() {
     
    document.getElementById('addRowBtn').addEventListener('click', () => openEditModal(null));
    document.getElementById('closeEditModal').addEventListener('click', closeEditModal);
    document.getElementById('editForm').addEventListener('submit', handleSave);

    const viewTableBtn = document.getElementById('viewTableBtn');
    const viewTreeBtn = document.getElementById('viewTreeBtn');
    if (viewTableBtn) viewTableBtn.addEventListener('click', showTableView);
    if (viewTreeBtn) viewTreeBtn.addEventListener('click', showTreeView);

    // Обробка кліку на пункти меню
    document.querySelectorAll('.nav-item, .nav-link').forEach(item => {
        item.addEventListener('click', function(e) {
            // Прибираємо активний стан з усіх пунктів
            document.querySelectorAll('.nav-item').forEach(nav => {
                nav.classList.remove('active');
            });
            // Додаємо активний стан до поточного пункту
            const navItem = this.closest('.nav-item') || this;
            if (navItem) {
                navItem.classList.add('active');
            }
        });
    });

    // Обробка кліку на підпункти меню (як у верхньому, так і в боковому меню)
    const sidePanel = document.getElementById('sidePanel');
    const sidePanelBackdrop = document.getElementById('sidePanelBackdrop');
    const menuToggle = document.getElementById('menuToggle');
    const sidePanelClose = document.getElementById('sidePanelClose');

    function openSidePanel() {
        if (sidePanel) sidePanel.classList.add('open');
        if (sidePanelBackdrop) sidePanelBackdrop.classList.add('open');
    }

    function closeSidePanel() {
        if (sidePanel) sidePanel.classList.remove('open');
        if (sidePanelBackdrop) sidePanelBackdrop.classList.remove('open');
    }

    if (menuToggle) {
        menuToggle.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            openSidePanel();
        });
    }

    if (sidePanelClose) {
        sidePanelClose.addEventListener('click', function(e) {
            e.preventDefault();
            closeSidePanel();
        });
    }

    if (sidePanelBackdrop) {
        sidePanelBackdrop.addEventListener('click', function() {
            closeSidePanel();
        });
    }

    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            closeSidePanel();
        }
    });

    document.querySelectorAll('.mega-menu-list a').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            if (this.classList.contains('nav-link-disabled') || this.getAttribute('data-enabled') === 'false') {
                return;
            }
            const menuItem = this.textContent.trim();
            handleMenuItemClick(menuItem);
            // Після вибору пункту закриваємо бокову панель (якщо відкрита)
            closeSidePanel();
        });
    });

   
    const closeButtons = document.querySelectorAll('.close');
    closeButtons.forEach(btn => {
        btn.addEventListener('click', function() {
            const modal = this.closest('.modal');
            if (modal) {
                if (modal.id === 'editModal') {
                    closeEditModal();
                } else {
                    closeModal();
                }
            }
        });
    });
    
}

function exportTableToExcel() {
    if (!window.tableData || !window.tableColumns || window.tableColumns.length === 0) {
        showModal('Немає даних для експорту', 'error');
        return;
    }

    if (typeof XLSX === 'undefined') {
        showModal('Експорт недоступний: бібліотека XLSX не завантажена', 'error');
        return;
    }

    const headerCells = Array.from(
        document.querySelectorAll('#tableHead th[data-column]')
    );
    
    if (!headerCells.length) {
        showModal('Немає заголовків таблиці для експорту', 'error');
        return;
    }

    const headers = headerCells.map(th => {
        const key = th.dataset.column;
        // Текст без іконки сорту (беремо перший текстовий вузол)
        let title = th.childNodes[0] && th.childNodes[0].textContent
            ? th.childNodes[0].textContent.trim()
            : th.textContent.trim();
        return { key, title };
    });

    // Формуємо двовимірний масив для XLSX: [рядки][колонки]
    const sheetData = [];

    // Перший рядок – заголовки
    sheetData.push(headers.map(h => h.title));

    // Рядки даних: експортуємо саме те, що зараз відображається (відфільтровані дані), без колонки Дії
    const dataToExport = window.displayedTableData !== undefined ? window.displayedTableData : window.tableData;
    if (!dataToExport || dataToExport.length === 0) {
        showModal('Немає даних для експорту (таблиця порожня або відфільтровано 0 рядків)', 'error');
        return;
    }
    dataToExport.forEach(row => {
        const rowArr = headers.map(h => {
            let value = row[h.key];
            if (value === null || value === undefined) value = '';
            value = String(value);

            // Для телефонів залишаємо як текст, щоб не губився початковий 0
            const isPhone =
                h.key.toLowerCase().includes('phone') ||
                h.title.toLowerCase() === 'телефон';
            if (isPhone) {
                return value; // XLSX збереже це як текст
            }

            return value;
        });
        sheetData.push(rowArr);
    });

    // Створюємо книгу та аркуш
    const wb = XLSX.utils.book_new();
    const wsName = (window.currentMenuItem || 'Лист').substring(0, 31); // Excel: max 31 символ
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Налаштовуємо ширину колонок на основі довжини заголовків та даних
    const colWidths = headers.map((h, colIndex) => {
        let maxLength = h.title.length; // Починаємо з довжини заголовка
        
        // Перевіряємо всі дані в цій колонці
        sheetData.forEach((row, rowIndex) => {
            if (rowIndex > 0) { // Пропускаємо рядок заголовків
                const cellValue = row[colIndex];
                if (cellValue) {
                    const cellLength = String(cellValue).length;
                    if (cellLength > maxLength) {
                        maxLength = cellLength;
                    }
                }
            }
        });
        
        // Додаємо невеликий запас (2-3 символи) та обмежуємо максимальну ширину
        return { wch: Math.min(maxLength + 3, 50) };
    });
    ws['!cols'] = colWidths;
    
   
    const thinBlack = { style: 'thin', color: { rgb: 'FF000000' } };
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;
            
            if (!ws[cellAddress].s) {
                ws[cellAddress].s = {};
            }
            if (R === 0) {
                ws[cellAddress].s.font = { bold: true };
            }
            ws[cellAddress].s.border = {
                top: thinBlack,
                bottom: thinBlack,
                left: thinBlack,
                right: thinBlack
            };
            if (!ws[cellAddress].s.alignment) {
                ws[cellAddress].s.alignment = {};
            }
            ws[cellAddress].s.alignment.wrapText = true;
            ws[cellAddress].s.alignment.vertical = 'top';
        }
    }
    
    XLSX.utils.book_append_sheet(wb, ws, wsName);

    // Генеруємо xlsx як масив байтів
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const date = new Date().toISOString().slice(0, 10);
    const title = (window.currentMenuItem || 'table').replace(/\s+/g, '_');
    a.href = url;
    a.download = `${title}_${date}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}


// Перевірка підключення
async function testConnection() {
    try {
        const response = await fetch('/api/test-connection');
        const data = await response.json();
        
        if (data.success) {
            showModal('Підключення успішно!\n\n' + data.version, 'success');
        } else {
            showModal('Помилка підключення: ' + data.error, 'error');
        }
    } catch (error) {
        showModal('Помилка: ' + error.message, 'error');
    }
}

// Обробка кліку на пункт меню
async function handleMenuItemClick(menuItem) {
    currentMenuItem = menuItem;
    
    // Показуємо панель
    const dataPanel = document.getElementById('dataPanel');
    const panelTitle = document.getElementById('panelTitle');
    const viewToggleWrap = document.getElementById('viewToggleWrap');
    const tableContainer = document.getElementById('tableContainer');
    const treeContainer = document.getElementById('treeContainer');

    panelTitle.textContent = menuItem;
    dataPanel.style.display = 'block';

    // Перемикач "Таблиця / Дерево" тільки для структури
    const isStructureView = menuItem === 'Загальна' || menuItem === 'Структура НРК';
    if (viewToggleWrap) viewToggleWrap.style.display = isStructureView ? '' : 'none';
    if (tableContainer) tableContainer.style.display = '';
    if (treeContainer) treeContainer.style.display = 'none';
    document.getElementById('viewTableBtn')?.classList.add('active');
    document.getElementById('viewTreeBtn')?.classList.remove('active');

   
    if (menuItem === 'Загальна') {
        await loadCustomQuery();
    } else if (menuItem === 'ВО загальні') {
        await loadVOQuery();
    } else if (menuItem === 'ВО нрк') {
        await loadVONRKQuery();
    } else if (menuItem === 'НРК') {
        await loadNRCQuery();
    } else if (menuItem === 'Структура НРК') {
        await loadNRCStructureQuery();
    } else if (menuItem === 'довідник засобів') {
        await loadMaterialsToNewTable();
    } else {
        // Мапінг пунктів меню на таблиці БД
        const menuToTable = {
            'Структура БПЛА': 'uav_structure'
        };
        
        const tableName = menuToTable[menuItem] || menuItem.toLowerCase().replace(/\s+/g, '_');
        await loadTableDataForMenu(tableName);
    }
}

// ——— Деревоподібна схема структури ———
function buildTreeFromFlat(flatData) {
    if (!flatData || !flatData.length) return [];
    const byId = new Map();
    flatData.forEach(row => {
        byId.set(row.id, { id: row.id, parent_id: row.parent_id, name: row.name || '', children: [] });
    });
    const roots = [];
    flatData.forEach(row => {
        const node = byId.get(row.id);
        if (!node) return;
        const parentId = row.parent_id;
        if (parentId == null || parentId === '' || !byId.has(parentId)) {
            roots.push(node);
        } else {
            const parent = byId.get(parentId);
            if (parent) parent.children.push(node);
            else roots.push(node);
        }
    });
    roots.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    roots.forEach(n => sortTreeChildren(n));
    return roots;
}
function sortTreeChildren(node) {
    if (node.children && node.children.length) {
        node.children.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
        node.children.forEach(sortTreeChildren);
    }
}
function renderTreeHtml(nodes) {
    if (!nodes || !nodes.length) return '';
    let html = '<ul class="tree-list">';
    nodes.forEach((node, i) => {
        const hasChildren = node.children && node.children.length > 0;
        const toggleClass = hasChildren ? 'tree-toggle' : 'tree-toggle tree-toggle--empty';
        const nodeClass = hasChildren ? 'tree-node' : 'tree-node tree-node--leaf';
        const chevron = hasChildren ? '<span class="tree-chevron" aria-hidden="true">▶</span>' : '<span class="tree-chevron tree-chevron--leaf" aria-hidden="true"></span>';
        html += `<li class="${nodeClass}" data-id="${node.id}">
          <div class="tree-node-inner">
            <span class="${toggleClass}">${chevron}</span>
            <span class="tree-label">${escapeHtml(node.name)}</span>
          </div>
          ${hasChildren ? `<div class="tree-children">${renderTreeHtml(node.children)}</div>` : ''}
        </li>`;
    });
    html += '</ul>';
    return html;
}
function escapeHtml(s) {
    if (s == null) return '';
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
}
function attachTreeToggleListeners(container) {
    if (!container) return;
    container.querySelectorAll('.tree-node').forEach(nodeEl => {
        const toggle = nodeEl.querySelector('.tree-toggle');
        const children = nodeEl.querySelector('.tree-children');
        if (toggle && children) {
            toggle.addEventListener('click', function (e) {
                e.stopPropagation();
                nodeEl.classList.toggle('tree-node--closed');
            });
        }
        const inner = nodeEl.querySelector('.tree-node-inner');
        if (inner) {
            inner.addEventListener('click', function () {
                if (children) nodeEl.classList.toggle('tree-node--closed');
            });
        }
    });
}
function attachTreeExpandCollapseAll(treeEl) {
    const btnExpand = document.getElementById('treeExpandAll');
    const btnCollapse = document.getElementById('treeCollapseAll');
    if (!treeEl || !btnExpand || !btnCollapse) return;
    btnExpand.onclick = function () {
        treeEl.querySelectorAll('.tree-node').forEach(el => {
            if (el.querySelector('.tree-children')) el.classList.remove('tree-node--closed');
        });
    };
    btnCollapse.onclick = function () {
        treeEl.querySelectorAll('.tree-node').forEach(el => {
            if (el.querySelector('.tree-children')) el.classList.add('tree-node--closed');
        });
    };
}
async function loadStructureTreeView() {
    const treeContainer = document.getElementById('treeContainer');
    const treeEl = document.getElementById('structureTree');
    const treeLoading = document.getElementById('treeLoading');
    const tableContainer = document.getElementById('tableContainer');
    if (!treeContainer || !treeEl) return;
    treeContainer.style.display = '';
    if (tableContainer) tableContainer.style.display = 'none';
    treeEl.innerHTML = '';
    treeLoading.style.display = 'block';
    try {
        const res = await fetch('/api/structure-tree');
        const data = await res.json();
        treeLoading.style.display = 'none';
        if (data.success && data.data && data.data.length) {
            const roots = buildTreeFromFlat(data.data);
            treeEl.innerHTML = renderTreeHtml(roots);
            attachTreeToggleListeners(treeEl);
            attachTreeExpandCollapseAll(treeEl);
        } else {
            treeEl.innerHTML = '<p class="tree-empty">Немає даних для дерева або помилка: ' + (data.error || 'порожній відповідь') + '</p>';
        }
    } catch (e) {
        treeLoading.style.display = 'none';
        treeEl.innerHTML = '<p class="tree-empty">Помилка завантаження: ' + e.message + '</p>';
    }
}
function showTableView() {
    const tableContainer = document.getElementById('tableContainer');
    const treeContainer = document.getElementById('treeContainer');
    if (tableContainer) tableContainer.style.display = '';
    if (treeContainer) treeContainer.style.display = 'none';
    document.getElementById('viewTableBtn')?.classList.add('active');
    document.getElementById('viewTreeBtn')?.classList.remove('active');
}
function showTreeView() {
    document.getElementById('viewTableBtn')?.classList.remove('active');
    document.getElementById('viewTreeBtn')?.classList.add('active');
    loadStructureTreeView();
}

// Завантаження даних для ВО загальні (той самий запит що й ВО НРК)
async function loadVOQuery() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '<tr><td colspan="100" class="loading">Завантаження даних...</td></tr>';
    tableBody.innerHTML = '';

    try {
        const query = `SELECT 
    uu.off_name AS l1,
    u.off_name  AS l2,
    bs.type_bps,
    tt.*
FROM subordination s
LEFT JOIN unit u  ON u.id = s.unit_id
LEFT JOIN unit uu ON uu.id = s.parent_id
LEFT JOIN bps_structure bs ON bs.unit_id = u.id
LEFT JOIN (
    SELECT 
        cp.id,
        cp.firstname,
        cp.middlename,
        cp.surname,
        cp.mil_name,
        cp.phone,
        cp.phone2,
        cp.phone3,
        r.name AS rank,
        p.name AS position,
        u.off_name AS v4,
        uu.off_name AS v4_verh,
        uuu.off_name AS v4_verh2,
        cp.comment
    FROM contact_person cp 
    LEFT JOIN position p ON cp.position_id = p.id
    LEFT JOIN rank r ON r.id = cp.rank_id
    LEFT JOIN unit u ON u.id = cp.unit_id
    LEFT JOIN subordination s ON s.unit_id = u.id
    LEFT JOIN unit uu ON uu.id = s.parent_id
    LEFT JOIN subordination ss ON ss.unit_id = uu.id
    LEFT JOIN unit uuu ON uuu.id = ss.parent_id
    WHERE u.comment NOT LIKE N'%морська%'
) tt ON tt.v4 = u.off_name`;
        
        const response = await fetch('/api/query', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ query: query })
        });

        const data = await response.json();

        if (data.success) {
            // Встановлюємо поточну таблицю для редагування
            currentTable = 'contact_person';
            
            // Отримуємо інформацію про таблицю для форми редагування
            const infoResponse = await fetch(`/api/table-info/contact_person`);
            const infoData = await infoResponse.json();
            if (infoData.success) {
                currentTableInfo = infoData;
            }
            
            // Зберігаємо дані для фільтрації та сортування
            window.tableData = data.data;
            
            // Формуємо порядок і набір видимих колонок (як у ВО НРК)
            const rawColumns = data.columns.filter(col => {
                const name = col.toLowerCase();
                return name !== 'id' && name !== 'v4' && name !== 'v4_verh' && name !== 'comment';
            });

            const preferredOrder = [
                'v4_verh2',   // Структура1
                'l1',         // Структура2
                'l2',         // Структура3
                'type_bps',   // Тип підрозділа
                'surname',    // Прізвище
                'firstname',  // Ім'я
                'middlename', // По-батькові
                'mil_name',   // Позивний
                'phone',      // Телефон
                'phone2',     // Телефон2
                'phone3',     // Телефон3
                'comments',   // Коментар
                'rank',       // Звання
                'position'    // Посада
            ];

            let visibleColumns = [];
            preferredOrder.forEach(colName => {
                const found = rawColumns.find(c => c.toLowerCase() === colName.toLowerCase());
                if (found && !visibleColumns.includes(found)) {
                    visibleColumns.push(found);
                }
            });
            rawColumns.forEach(col => {
                if (!visibleColumns.includes(col)) {
                    visibleColumns.push(col);
                }
            });
            window.tableColumns = visibleColumns;
            
            // Заголовки + колонка для дій з потрібними назвами
            tableHead.innerHTML = '<tr>' + 
                visibleColumns.map(col => {
                    const name = col.toLowerCase();
                    let displayName = col;
                    if (name === 'v4_verh2') {
                        displayName = "Структура1";
                    } else if (name === 'l1') {
                        displayName = "Структура2";
                    } else if (name === 'l2') {
                        displayName = "Структура3";
                    } else if (name === 'type_bps') {
                        displayName = "Наявність НРК";
                    } else if (name === 'surname') {
                        displayName = "Прізвище";
                    } else if (name === 'middlename') {
                        displayName = "По-батькові";
                    } else if (name === 'firstname') {
                        displayName = "Ім'я";
                    } else if (name === 'mil_name') {
                        displayName = "Позивний";
                    } else if (name === 'phone') {
                        displayName = "Телефон";
                    } else if (name === 'phone2') {
                        displayName = "Телефон2";
                    } else if (name === 'phone3') {
                        displayName = "Телефон3";
                    } else if (name === 'comments') {
                        displayName = "Коментар";
                    } else if (name === 'rank') {
                        displayName = "Звання";
                    } else if (name === 'position') {
                        displayName = "Посада";
                    } else if (name === 'comment') {
                        displayName = "Коментар2";
                    }
                    return `<th data-column="${col}">${displayName} <span class="sort-icon">↕</span></th>`;
                }).join('') + 
                '<th>Дії</th></tr>';

            // Додаємо обробники для сортування
            setTimeout(() => {
                document.querySelectorAll('th[data-column]').forEach(th => {
                    th.style.cursor = 'pointer';
                    th.addEventListener('click', () => sortTable(th.dataset.column));
                });
            }, 100);

            // Додаємо рядок пошуку
            addSearchRow();

            // Якщо були збережені фільтри — застосовуємо їх, інакше показуємо всі дані
            if (!applySavedSearchFiltersIfAny()) {
                renderTable(data.data, visibleColumns);
            }
        } else {
            tableHead.innerHTML = '';
            tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
        }
    } catch (error) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
    }
}

// Завантаження даних для ВО НРК
async function loadVONRKQuery() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '<tr><td colspan="100" class="loading">Завантаження даних...</td></tr>';
    tableBody.innerHTML = '';

    try {
        const query = `SELECT 
    uu.off_name AS l1,
    u.off_name  AS l2,
    bs.type_bps,
    tt.*
FROM subordination s
LEFT JOIN unit u  ON u.id = s.unit_id
LEFT JOIN unit uu ON uu.id = s.parent_id
LEFT JOIN bps_structure bs ON bs.unit_id = u.id
LEFT JOIN (
    SELECT 
        cp.id,
        cp.firstname,
        cp.middlename,
        cp.surname,
        cp.mil_name,
        cp.phone,
        cp.phone2,
        cp.phone3,
        r.name AS rank,
        p.name AS position,
        u.off_name AS v4,
        uu.off_name AS v4_verh,
        uuu.off_name AS v4_verh2,
        cp.comment,
        cp.type_bps
    FROM contact_person cp 
    LEFT JOIN position p ON cp.position_id = p.id
    LEFT JOIN rank r ON r.id = cp.rank_id
    LEFT JOIN unit u ON u.id = cp.unit_id
    LEFT JOIN subordination s ON s.unit_id = u.id
    LEFT JOIN unit uu ON uu.id = s.parent_id
    LEFT JOIN subordination ss ON ss.unit_id = uu.id
    LEFT JOIN unit uuu ON uuu.id = ss.parent_id
    WHERE u.comment NOT LIKE N'%морська%'
) tt ON tt.v4 = u.off_name
WHERE bs.type_bps = 'nrk' OR tt.type_bps = 'bps'`;
        
        const response = await fetch('/api/query', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ query: query })
        });

        const data = await response.json();

        if (data.success) {
            // Встановлюємо поточну таблицю для редагування (контакти ВО)
            currentTable = 'contact_person';
            
            // Отримуємо інформацію про таблицю для форми редагування
            const infoResponse = await fetch(`/api/table-info/contact_person`);
            const infoData = await infoResponse.json();
            if (infoData.success) {
                currentTableInfo = infoData;
            }
            
            // Зберігаємо дані для фільтрації та сортування
            window.tableData = data.data;
            
            // Формуємо порядок і набір видимих колонок:
            // 1) приховуємо технічне поле ID, v4, v4_verh, comment (Коментар2) та type_bps (Тип підрозділа)
            // 2) будуємо явний порядок колонок для ВО НРК
            const rawColumns = data.columns.filter(col => {
                const name = col.toLowerCase();
                return name !== 'id' && name !== 'v4' && name !== 'v4_verh' && name !== 'comment' && name !== 'type_bps';
            });

            const preferredOrder = [
                'v4_verh2',   // Структура1
                'l1',         // Структура2
                'l2',         // Структура3
                'surname',    // Прізвище
                'firstname',  // Ім'я
                'middlename', // По-батькові
                'mil_name',   // Позивний
                'phone',      // Телефон
                'phone2',     // Телефон2
                'phone3',     // Телефон3
                'comments',   // Коментар
                'rank',       // Звання
                'position'    // Посада
            ];

            let visibleColumns = [];
            preferredOrder.forEach(colName => {
                const found = rawColumns.find(c => c.toLowerCase() === colName.toLowerCase());
                if (found && !visibleColumns.includes(found)) {
                    visibleColumns.push(found);
                }
            });

            // На випадок наявності інших колонок додаємо їх в кінці
            rawColumns.forEach(col => {
                if (!visibleColumns.includes(col)) {
                    visibleColumns.push(col);
                }
            });
            
            window.tableColumns = visibleColumns;
            
            // Заголовки + колонка для дій з потрібними назвами
            tableHead.innerHTML = '<tr>' + 
                visibleColumns.map(col => {
                    const name = col.toLowerCase();
                    let displayName = col;
                    if (name === 'v4_verh2') {
                        displayName = "Структура1";
                    } else if (name === 'l1') {
                        displayName = "Структура2";
                    } else if (name === 'l2') {
                        displayName = "Структура3";
                    } else if (name === 'type_bps') {
                        displayName = "Тип підрозділа";
                    } else if (name === 'surname') {
                        displayName = "Прізвище";
                    } else if (name === 'middlename') {
                        displayName = "По-батькові";
                    } else if (name === 'firstname') {
                        displayName = "Ім'я";
                    } else if (name === 'mil_name') {
                        displayName = "Позивний";
                    } else if (name === 'phone') {
                        displayName = "Телефон";
                    } else if (name === 'phone2') {
                        displayName = "Телефон2";
                    } else if (name === 'phone3') {
                        displayName = "Телефон3";
                    } else if (name === 'comments') {
                        displayName = "Коментар";
                    } else if (name === 'rank') {
                        displayName = "Звання";
                    } else if (name === 'position') {
                        displayName = "Посада";
                    } else if (name === 'comment') {
                        displayName = "Коментар2";
                    }
                    return `<th data-column="${col}">${displayName} <span class="sort-icon">↕</span></th>`;
                }).join('') + 
                '<th>Дії</th></tr>';

            // Додаємо обробники для сортування
            setTimeout(() => {
                document.querySelectorAll('th[data-column]').forEach(th => {
                    th.style.cursor = 'pointer';
                    th.addEventListener('click', () => sortTable(th.dataset.column));
                });
            }, 100);

            // Додаємо рядок пошуку
            addSearchRow();

            // Якщо були збережені фільтри — застосовуємо їх, інакше показуємо всі дані
            if (!applySavedSearchFiltersIfAny()) {
                renderTable(data.data, visibleColumns);
            }
        } else {
            tableHead.innerHTML = '';
            tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
        }
    } catch (error) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
    }
}

// Завантаження даних для Структура НРК
async function loadNRCStructureQuery() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '<tr><td colspan="100" class="loading">Завантаження даних...</td></tr>';
    tableBody.innerHTML = '';

    try {
        const query = `SELECT 
            s.unit_id,
            uu.off_name AS l1,
            u.off_name  AS l2,
            bs.type_bps,
            bs.unit_structure_id,
            us.name AS unit_structure_name
        FROM subordination s
        LEFT JOIN unit u  ON u.id = s.unit_id
        LEFT JOIN unit uu ON uu.id = s.parent_id
        LEFT JOIN bps_structure bs ON bs.unit_id = u.id
        LEFT JOIN unit_structure us ON bs.unit_structure_id = us.id`;
        
        const response = await fetch('/api/query', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ query: query })
        });

        const data = await response.json();

        if (data.success) {
            // Встановлюємо поточну таблицю для редагування
            currentTable = 'subordination';
            
            // Отримуємо інформацію про таблицю для форми редагування
            const infoResponse = await fetch(`/api/table-info/subordination`);
            const infoData = await infoResponse.json();
            if (infoData.success) {
                currentTableInfo = infoData;
            }
            
            // Зберігаємо дані для фільтрації та сортування
            window.tableData = data.data;
            
            // Приховуємо технічні поля ID та unit_structure_id з відображення, але залишаємо в даних
            const visibleColumns = data.columns.filter(col => {
                const name = col.toLowerCase();
                return name !== 'id' && name !== 'unit_structure_id';
            });
            window.tableColumns = visibleColumns;
            
            // Заголовки + колонка для дій
            tableHead.innerHTML = '<tr>' + 
                visibleColumns.map(col => `<th data-column="${col}">${col} <span class="sort-icon">↕</span></th>`).join('') + 
                '<th>Дії</th></tr>';

            // Додаємо обробники для сортування
            setTimeout(() => {
                document.querySelectorAll('th[data-column]').forEach(th => {
                    th.style.cursor = 'pointer';
                    th.addEventListener('click', () => sortTable(th.dataset.column));
                });
            }, 100);

            // Додаємо рядок пошуку
            addSearchRow();

            // Якщо були збережені фільтри — застосовуємо їх, інакше показуємо всі дані
            if (!applySavedSearchFiltersIfAny()) {
                renderTable(data.data, visibleColumns);
            }
        } else {
            tableHead.innerHTML = '';
            tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
            updateTableFooter(0);
        }
    } catch (error) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
        updateTableFooter(0);
    }
}

// Завантаження даних для НРК
async function loadNRCQuery() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '<tr><td colspan="100" class="loading">Завантаження даних...</td></tr>';
    tableBody.innerHTML = '';

    try {
        const query = `SELECT 
    material.id         AS id,
    material.name       AS material_name,
    material.code       AS code,
    gg.name             AS group_name,
    manufacture.name    AS manufacture_name,
    url
FROM Material AS material
LEFT JOIN material_group_rel AS ff ON ff.material_id = material.id
LEFT JOIN material_group      AS gg ON ff.group_id    = gg.id
LEFT JOIN manufacture         AS manufacture ON material.manufacture_id = manufacture.id`;
        
        const response = await fetch('/api/query', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ query: query })
        });

        const data = await response.json();

        if (data.success) {
            // Встановлюємо поточну таблицю для редагування
            currentTable = 'Material';
            
            // Отримуємо інформацію про таблицю для форми редагування
            const infoResponse = await fetch(`/api/table-info/Material`);
            const infoData = await infoResponse.json();
            if (infoData.success) {
                currentTableInfo = infoData;
            }
            
            // Зберігаємо дані для фільтрації та сортування
            window.tableData = data.data;

            // Приховуємо технічні поля URL, ID та off_name з відображення,
            // але залишаємо їх в даних (для посилань та редагування)
            const visibleColumns = data.columns.filter(col => {
                const colLower = col.toLowerCase();
                return colLower !== 'url' && colLower !== 'id' && colLower !== 'off_name';
            });
            window.tableColumns = visibleColumns;
            
            // Заголовки + колонка для дій (без колонки url)
            tableHead.innerHTML = '<tr>' + 
                visibleColumns.map(col => {
                    let displayName = col;
                    const colLower = col.toLowerCase();
                    if (colLower === 'material_name') {
                        displayName = 'Матеріал';
                    } else if (colLower === 'code') {
                        displayName = 'Код';
                    } else if (colLower === 'group_name') {
                        displayName = 'Група';
                    } else if (colLower === 'manufacture_name') {
                        displayName = 'Виробник';
                    }
                    return `<th data-column="${col}">${displayName} <span class="sort-icon">↕</span></th>`;
                }).join('') + 
                '<th>Дії</th></tr>';

            // Додаємо обробники для сортування
            setTimeout(() => {
                document.querySelectorAll('th[data-column]').forEach(th => {
                    th.style.cursor = 'pointer';
                    th.addEventListener('click', () => sortTable(th.dataset.column));
                });
            }, 100);

            // Додаємо рядок пошуку
            addSearchRow();

            // Сохраняем позицию скролла перед рендерингом (только если это не первая загрузка)
            const tableContainer = document.querySelector('.table-container');
            const scrollPosition = tableContainer && window.savedScrollPosition !== undefined ? window.savedScrollPosition : 0;
            
            // Якщо були збережені фільтри — застосовуємо їх, інакше показуємо всі дані
            if (!applySavedSearchFiltersIfAny()) {
                // Відображаємо дані з посиланнями та кнопками редагування для НРК
                renderNRCTable(data.data, visibleColumns);
            }
            
            // Восстанавливаем позицию скролла после рендеринга
            if (scrollPosition > 0) {
                setTimeout(() => {
                    const container = document.querySelector('.table-container');
                    if (container) {
                        container.scrollTop = scrollPosition;
                    }
                    window.savedScrollPosition = undefined; // Очищаем сохраненную позицию
                }, 50);
            }
        } else {
            tableHead.innerHTML = '';
            tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
            updateTableFooter(0);
        }
    } catch (error) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
        updateTableFooter(0);
    }
}

// Завантаження даних через кастомний SQL запит
async function loadCustomQuery() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '<tr><td colspan="100" class="loading">Завантаження даних...</td></tr>';
    tableBody.innerHTML = '';

    try {
        const query = `SELECT 
                          uu.off_name AS name3,
                          uu.name     AS name4,
                          u.off_name  AS name2,
                          u.name      AS name1,
                          cp.firstname,
                          cp.middlename,
                          cp.surname,
                          cp.phone
                       FROM unit u
                       JOIN subordination s ON u.id = s.unit_id
                       JOIN unit uu ON uu.id = s.parent_id 
                       LEFT JOIN contact_person cp 
                            ON cp.unit_id = u.id AND cp.type_bps = 'bps'`;
        
        const response = await fetch('/api/query', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ query: query })
        });

        const data = await response.json();

        if (data.success) {
            // Зберігаємо дані для фільтрації та сортування
            window.tableData = data.data;
            window.tableColumns = data.columns;
            
            // Мапінг технічних назв колонок у зрозумілі заголовки
            const headerRowHtml = data.columns.map(col => {
                const name = col.toLowerCase();
                let displayName = col;
                if (name === 'name3') {
                    displayName = "Структура1";
                } else if (name === 'name4') {
                    displayName = "Код1";
                } else if (name === 'name2') {
                    displayName = "Структура2";
                } else if (name === 'name1') {
                    displayName = "Код2";
                } else if (name === 'surname') {
                    displayName = "Прізвище";
                } else if (name === 'firstname') {
                    displayName = "Ім'я";
                } else if (name === 'middlename') {
                    displayName = "По-батькові";
                } else if (name === 'phone') {
                    displayName = "Телефон";
                } else if (name === 'phone2') {
                    displayName = "Телефон2";
                } else if (name === 'phone3') {
                    displayName = "Телефон3";
                }
                return `<th data-column="${col}">${displayName} <span class="sort-icon">↕</span></th>`;
            }).join('');
            
            // Заголовки + колонка для дій
            tableHead.innerHTML = '<tr>' + headerRowHtml + '<th>Дії</th></tr>';

            // Додаємо обробники для сортування
            setTimeout(() => {
                document.querySelectorAll('th[data-column]').forEach(th => {
                    th.style.cursor = 'pointer';
                    th.addEventListener('click', () => sortTable(th.dataset.column));
                });
            }, 100);

            // Додаємо рядок пошуку
            addSearchRow();
            
            // Відображаємо дані
            renderTable(data.data, data.columns);
        } else {
            tableHead.innerHTML = '';
            tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
            updateTableFooter(0);
        }
    } catch (error) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
        updateTableFooter(0);
    }
}

// Додавання рядка пошуку
function addSearchRow() {
    const table = document.getElementById('dataTable');
    const thead = document.getElementById('tableHead');
    
    // Удаляем старую строку поиска, если она существует, чтобы обновить уникальные значения
    const existingSearchRow = document.getElementById('searchRow');
    if (existingSearchRow) {
        existingSearchRow.remove();
    }
    
    const searchRow = document.createElement('tr');
    searchRow.id = 'searchRow';
    searchRow.className = 'search-row';
    
    if (window.tableColumns) {
        const noUniqueFilterColumns = ['surname', 'firstname', 'middlename', 'mil_name', 'l2', 'phone', 'phone2', 'phone3'];
        window.tableColumns.forEach(col => {
            const td = document.createElement('td');
            const wrapper = document.createElement('div');
            wrapper.className = 'search-filter-cell';
            const colLower = col.toLowerCase();
            const skipUniqueSelect = noUniqueFilterColumns.includes(colLower);

            if (!skipUniqueSelect) {
                const select = document.createElement('select');
                select.className = 'search-select';
                select.dataset.column = col;
                select.title = 'Фільтр за унікальними значеннями';
                select.addEventListener('change', handleSearch);

                const optionAll = document.createElement('option');
                optionAll.value = '';
                optionAll.textContent = '— Всі —';
                select.appendChild(optionAll);

                if (window.tableData && window.tableData.length > 0) {
                    const hasEmpty = window.tableData.some(row => {
                        const v = row[col];
                        return v === null || v === undefined || v === '';
                    });
                    if (hasEmpty) {
                        const optEmpty = document.createElement('option');
                        optEmpty.value = '__empty__';
                        optEmpty.textContent = 'Пусто';
                        select.appendChild(optEmpty);
                    }
                    const uniqueValues = [...new Set(window.tableData
                        .map(row => row[col])
                        .filter(val => val !== null && val !== undefined && val !== '')
                    )].sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

                    uniqueValues.forEach(value => {
                        const option = document.createElement('option');
                        const strVal = String(value);
                        option.value = strVal;
                        let displayText = strVal;
                        if (colLower === 'type_bps' && strVal.toLowerCase() === 'nrk') {
                            displayText = 'є';
                        } else if (strVal.length > 80) {
                            displayText = strVal.slice(0, 77) + '…';
                        }
                        option.textContent = displayText;
                        select.appendChild(option);
                    });
                }
                wrapper.appendChild(select);
            }

            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'search-input';
            input.placeholder = 'Пошук...';
            input.dataset.column = col;
            input.title = 'Контекстний пошук (входить у текст)';
            input.addEventListener('input', handleSearch);
            wrapper.appendChild(input);

            td.appendChild(wrapper);
            searchRow.appendChild(td);
        });
        
        // Порожня клітинка для колонки дій
        const td = document.createElement('td');
        searchRow.appendChild(td);
    }
    
    thead.appendChild(searchRow);
    
    // Встановлюємо правильну позицію для рядка пошуку після додавання
    const updateSearchRowPosition = () => {
        const firstRow = thead.querySelector('tr:first-child');
        if (firstRow && searchRow) {
            const headerHeight = firstRow.offsetHeight;
            searchRow.style.top = headerHeight + 'px';
            searchRow.style.position = 'sticky';
        }
    };
    
    // Викликаємо одразу та після невеликої затримки
    updateSearchRowPosition();
    setTimeout(updateSearchRowPosition, 100);
    setTimeout(updateSearchRowPosition, 300);
    
    // Оновлюємо позицію при зміні розміру
    if (window.ResizeObserver) {
        const firstRow = thead.querySelector('tr:first-child');
        if (firstRow) {
            const resizeObserver = new ResizeObserver(() => {
                updateSearchRowPosition();
            });
            resizeObserver.observe(firstRow);
        }
    }
}

// Збір значень фільтрів: точний вибір (select, включно "Пусто") + контекстний пошук (input)
function getSearchFilters() {
    const exact = {};
    const contains = {};
    document.querySelectorAll('.search-select').forEach(select => {
        if (select.value.trim()) {
            exact[select.dataset.column] = select.value;
        }
    });
    document.querySelectorAll('.search-input').forEach(input => {
        if (input.value.trim()) {
            contains[input.dataset.column] = input.value.trim().toLowerCase();
        }
    });
    return { exact, contains };
}

// Збережені фільтри між перезавантаженнями таблиці
window.savedSearchFilters = null;

// Застосування збережених фільтрів до поточного рядка пошуку (якщо є)
function applySavedSearchFiltersIfAny() {
    if (!window.savedSearchFilters) {
        return false;
    }

    const { exact, contains } = window.savedSearchFilters;

    // Відновлюємо значення select'ів
    document.querySelectorAll('.search-select').forEach(select => {
        const col = select.dataset.column;
        if (exact && Object.prototype.hasOwnProperty.call(exact, col)) {
            select.value = exact[col];
        }
    });

    // Відновлюємо значення input'ів
    document.querySelectorAll('.search-input').forEach(input => {
        const col = input.dataset.column;
        if (contains && Object.prototype.hasOwnProperty.call(contains, col)) {
            input.value = contains[col];
        }
    });

    // Після відновлення значень одразу застосовуємо фільтрацію
    const filters = getSearchFilters();
    let filteredData = applySearchFilters(window.tableData, filters);

    // Застосовуємо сортування, якщо воно є
    if (window.currentSortColumn) {
        filteredData = sortData(filteredData, window.currentSortColumn, window.sortDirection);
    }

    if (currentMenuItem === 'НРК') {
        renderNRCTable(filteredData, window.tableColumns);
    } else {
        renderTable(filteredData, window.tableColumns);
    }

    return true;
}

// Застосування фільтрів до даних (exact + contains; exact може бути __empty__ для "Пусто")
function applySearchFilters(data, filters) {
    if (!data) return [];
    const { exact, contains } = filters;
    return data.filter(row => {
        for (const col in exact) {
            const value = row[col];
            if (exact[col] === '__empty__') {
                const isEmpty = value === null || value === undefined || String(value).trim() === '';
                if (!isEmpty) return false;
            } else {
                if (value === null || value === undefined) {
                    return false;
                }
                if (String(value) !== exact[col]) return false;
            }
        }
        for (const col in contains) {
            const value = row[col];
            if (value === null || value === undefined) {
                return false;
            }
            if (!String(value).toLowerCase().includes(contains[col])) return false;
        }
        return true;
    });
}

// Обробка пошуку (унікальні значення + контекстний пошук)
function handleSearch(e) {
    const filters = getSearchFilters();
    const filteredData = applySearchFilters(window.tableData, filters);
    
    // Застосовуємо сортування, якщо воно є
    if (window.currentSortColumn) {
        filteredData = sortData(filteredData, window.currentSortColumn, window.sortDirection);
    }
    
    // Використовуємо renderNRCTable для таблиці НРК, інакше renderTable
    if (currentMenuItem === 'НРК') {
        renderNRCTable(filteredData, window.tableColumns);
    } else {
        renderTable(filteredData, window.tableColumns);
    }
}

// Сортування таблиці
let sortDirection = 'asc';
let currentSortColumn = null;

function sortTable(column) {
    if (currentSortColumn === column) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortDirection = 'asc';
        currentSortColumn = column;
    }
    
    window.currentSortColumn = column;
    window.sortDirection = sortDirection;
    
    // Оновлюємо іконки та підсвічування сортування
    document.querySelectorAll('th[data-column]').forEach(th => {
        th.classList.remove('sorted-asc', 'sorted-desc');
        const icon = th.querySelector('.sort-icon');
        if (icon) {
            icon.textContent = '↕';
        }
    });
    
    const currentTh = document.querySelector(`th[data-column="${column}"]`);
    if (currentTh) {
        const currentIcon = currentTh.querySelector('.sort-icon');
        if (currentIcon) {
            currentIcon.textContent = sortDirection === 'asc' ? '↑' : '↓';
        }
        currentTh.classList.add(sortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc');
    }
    
    const filters = getSearchFilters();
    const dataToSort = applySearchFilters(window.tableData, filters);
    const sortedData = sortData(dataToSort, column, sortDirection);
    
    // Використовуємо renderNRCTable для таблиці НРК, інакше renderTable
    if (currentMenuItem === 'НРК') {
        renderNRCTable(sortedData, window.tableColumns);
    } else {
        renderTable(sortedData, window.tableColumns);
    }
}

// Функція сортування даних
function sortData(data, column, direction) {
    return [...data].sort((a, b) => {
        let aVal = a[column];
        let bVal = b[column];
        
        if (aVal === null || aVal === undefined) aVal = '';
        if (bVal === null || bVal === undefined) bVal = '';
        
        aVal = String(aVal).toLowerCase();
        bVal = String(bVal).toLowerCase();
        
        if (direction === 'asc') {
            return aVal > bVal ? 1 : aVal < bVal ? -1 : 0;
        } else {
            return aVal < bVal ? 1 : aVal > bVal ? -1 : 0;
        }
    });
}

// Відображення таблиці (зберігаємо дані для експорту — саме те, що відображається)
function renderTable(data, columns) {
    window.displayedTableData = data || [];
    const tableBody = document.getElementById('tableBody');
    
    if (!data || data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="100" style="text-align: center; padding: 40px; color: #999;">Немає даних</td></tr>';
    } else {
        tableBody.innerHTML = data.map((row, index) => {
            return '<tr data-row-index="' + index + '">' + 
                columns.map(col => {
                    let value = row[col];
                    if (value === null || value === undefined) value = '';
                    let cellContent = value;
                    if (col.toLowerCase() === 'type_bps' && String(value).toLowerCase() === 'nrk') {
                        cellContent = '<span title="НРК">є</span>';
                    }
                    return `<td data-column="${col}">${cellContent}</td>`;
                }).join('') + 
                '<td>' +
                    '<button class="edit-btn" data-row-data=\'' + JSON.stringify(row).replace(/'/g, "&#39;") + '\' title="Редагувати">✏️</button> ' +
                    '<button class="delete-btn" data-row-data=\'' + JSON.stringify(row).replace(/'/g, "&#39;") + '\' title="Видалити">🗑️</button>' +
                '</td>' +
                '</tr>';
        }).join('');
    }
    
    // Оновлюємо футер з кількістю записів
    updateTableFooter(data.length);
    
    // Додаємо обробники для кнопок редагування та видалення
    setTimeout(() => {
        document.querySelectorAll('.edit-btn').forEach(btn => {
            btn.addEventListener('click', async function() {
                try {
                    const rowData = JSON.parse(this.getAttribute('data-row-data'));
                    
                    // Для таблицы ВО (contact_person) загружаем полную запись с ID
                    if ((currentMenuItem === 'ВО загальні' || currentMenuItem === 'ВО нрк') && currentTable === 'contact_person') {
                        const whereConditions = {};
                        
                        // Используем ID если он есть
                        if (rowData.id) {
                            whereConditions.id = rowData.id;
                        } else {
                            // Если нет ID, используем firstname и surname для поиска
                            if (rowData.firstname) whereConditions.firstname = rowData.firstname;
                            if (rowData.surname) whereConditions.surname = rowData.surname;
                        }
                        
                        if (Object.keys(whereConditions).length > 0) {
                            try {
                                const response = await fetch('/api/get-row', {
                                    method: 'POST',
                                    headers: {
                                        'Content-Type': 'application/json'
                                    },
                                    body: JSON.stringify({
                                        table: 'contact_person',
                                        where: whereConditions
                                    })
                                });
                                
                                const fullRowData = await response.json();
                                if (fullRowData.success) {
                                    openEditModal(fullRowData.data);
                                    return;
                                } else {
                                    console.warn('Не вдалося завантажити повну запис:', fullRowData.error);
                                }
                            } catch (error) {
                                console.error('Помилка завантаження повної записи:', error);
                            }
                        }
                    }
                    
                    // Для більшості таблиць, включаючи «Структура НРК», використовуємо дані поточного рядка
                    openEditModal(rowData);
                } catch (e) {
                    console.error('Помилка парсингу даних рядка:', e);
                }
            });
        });
        
        document.querySelectorAll('.delete-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                try {
                    const rowData = JSON.parse(this.getAttribute('data-row-data'));
                    handleDelete(rowData);
                } catch (e) {
                    console.error('Помилка парсингу даних рядка:', e);
                }
            });
        });
    }, 100);
}

// Відображення таблиці НРК з посиланнями та кнопками редагування (зберігаємо дані для експорту)
function renderNRCTable(data, columns) {
    window.displayedTableData = data || [];
    const tableBody = document.getElementById('tableBody');
    
    if (!data || data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="100" style="text-align: center; padding: 40px; color: #999;">Немає даних</td></tr>';
    } else {
        tableBody.innerHTML = data.map((row, index) => {
            // Беремо URL з даних рядка (поле url)
            const linkUrl = row.url || row.URL || '';
            // Екрануємо посилання для безпеки
            const safeLink = linkUrl.replace(/'/g, "&#39;").replace(/"/g, "&quot;");
            
            // Якщо є URL, робимо рядок клікабельним
            const rowClass = linkUrl ? 'clickable-row' : '';
            const rowOnClick = linkUrl ? `onclick="window.open('${safeLink}', '_blank')"` : '';
            const rowStyle = linkUrl ? 'cursor: pointer;' : '';
            
            // Екрануємо дані рядка для JSON
            const rowDataEscaped = JSON.stringify(row).replace(/'/g, "&#39;");
            
            return '<tr data-row-index="' + index + '" class="' + rowClass + '" ' + rowOnClick + ' style="' + rowStyle + '">' + 
                columns.map(col => {
                    let value = row[col];
                    // Для естетичного вигляду замінюємо null/undefined на порожній рядок
                    if (value === null || value === undefined) value = '';
                    return `<td>${value}</td>`;
                }).join('') + 
                '<td>' +
                    (linkUrl ? '<a href="' + safeLink + '" target="_blank" onclick="event.stopPropagation();" style="text-decoration: none; color: #0066cc; font-size: 18px; margin-right: 10px;" title="Відкрити посилання">🔗</a>' : '<span style="display: inline-block; width: 28px;"></span>') +
                    '<button class="edit-btn" data-row-data=\'' + rowDataEscaped + '\' title="Редагувати" onclick="event.stopPropagation();">✏️</button> ' +
                    '<button class="delete-btn" data-row-data=\'' + rowDataEscaped + '\' title="Видалити" onclick="event.stopPropagation();">🗑️</button>' +
                '</td>' +
                '</tr>';
        }).join('');
    }
    
    // Оновлюємо футер з кількістю записів
    updateTableFooter(data.length);
    
    // Додаємо обробники для кнопок редагування та видалення
    setTimeout(() => {
        document.querySelectorAll('.edit-btn').forEach(btn => {
            btn.addEventListener('click', async function(e) {
                e.stopPropagation();
                try {
                    const rowData = JSON.parse(this.getAttribute('data-row-data'));
                    
                    // Для НРК завантажуємо повну запис з БД
                    if (currentMenuItem === 'НРК' && currentTable === 'Material') {
                        const whereConditions = {};
                        
                        // Спочатку намагаємося використати ID, якщо він є
                        if (rowData.id) {
                            whereConditions.id = rowData.id;
                        } else {
                            // Якщо немає ID, використовуємо name та off_name
                            if (rowData.name) whereConditions.name = rowData.name;
                            if (rowData.off_name) whereConditions.off_name = rowData.off_name;
                        }
                        
                        if (Object.keys(whereConditions).length > 0) {
                            try {
                                const response = await fetch('/api/get-row', {
                                    method: 'POST',
                                    headers: {
                                        'Content-Type': 'application/json'
                                    },
                                    body: JSON.stringify({
                                        table: 'Material',
                                        where: whereConditions
                                    })
                                });
                                
                                const fullRowData = await response.json();
                                if (fullRowData.success) {
                                    openEditModal(fullRowData.data);
                                    return;
                                } else {
                                    console.warn('Не вдалося завантажити повну запис:', fullRowData.error);
                                }
                            } catch (error) {
                                console.error('Помилка завантаження повної записи:', error);
                            }
                        }
                    }
                    
                    // Якщо не вдалося завантажити повну запис, використовуємо наявні дані
                    openEditModal(rowData);
                } catch (e) {
                    console.error('Помилка парсингу даних рядка:', e);
                }
            });
        });
        
        document.querySelectorAll('.delete-btn').forEach(btn => {
            btn.addEventListener('click', async function(e) {
                e.stopPropagation();
                try {
                    const rowData = JSON.parse(this.getAttribute('data-row-data'));
                    handleDelete(rowData);
                } catch (e) {
                    console.error('Помилка парсингу даних рядка:', e);
                }
            });
        });
    }, 100);
}

// Оновлення футера таблиці
function updateTableFooter(filteredCount) {
    const recordCountEl = document.getElementById('recordCount');
    const filteredCountEl = document.getElementById('filteredCount');
    
    const totalCount = window.tableData ? window.tableData.length : 0;
    
    if (recordCountEl) {
        recordCountEl.textContent = `Загальна кількість: ${totalCount}`;
    }
    
    if (filteredCountEl) {
        filteredCountEl.textContent = `Відфільтровано: ${filteredCount}`;
    }
}

// Завантаження даних таблиці для меню
async function loadTableDataForMenu(tableName) {
    currentTable = tableName;
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '<tr><td colspan="100" class="loading">Завантаження даних...</td></tr>';
    tableBody.innerHTML = '';

    try {
        // Спочатку отримуємо інформацію про таблицю
        const infoResponse = await fetch(`/api/table-info/${tableName}`);
        const infoData = await infoResponse.json();
        
        if (infoData.success) {
            currentTableInfo = infoData;
        }
        
        // Завантажуємо дані
        const response = await fetch(`/api/table/${tableName}`);
        const data = await response.json();

        if (data.success) {
            // Заголовки + колонка для дій
            tableHead.innerHTML = '<tr>' + 
                data.columns.map(col => `<th data-column="${col}">${col} <span class="sort-icon">↕</span></th>`).join('') + 
                '<th>Дії</th></tr>';
            
            // Додаємо обробники для сортування
            document.querySelectorAll('th[data-column]').forEach(th => {
                th.style.cursor = 'pointer';
                th.addEventListener('click', () => sortTable(th.dataset.column));
            });

            // Дані
            if (data.data.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="100" style="text-align: center; padding: 40px; color: #999;">Немає даних</td></tr>';
            } else {
                tableBody.innerHTML = data.data.map((row, index) => {
                    return '<tr data-row-index="' + index + '">' + 
                        data.columns.map(col => {
                            let value = row[col];
                            if (value === null) value = '<em style="color: #999;">NULL</em>';
                            return `<td>${value}</td>`;
                        }).join('') + 
                        '<td>' +
                            '<button class="edit-btn" data-row-data=\'' + JSON.stringify(row).replace(/'/g, "&#39;") + '\' title="Редагувати">✏️</button> ' +
                            '<button class="delete-btn" data-row-data=\'' + JSON.stringify(row).replace(/'/g, "&#39;") + '\' title="Видалити">🗑️</button>' +
                        '</td>' +
                        '</tr>';
                }).join('');
            }
            
            // Додаємо обробники для кнопок редагування та видалення
            document.querySelectorAll('.edit-btn').forEach(btn => {
                btn.addEventListener('click', function() {
                    const rowData = JSON.parse(this.getAttribute('data-row-data'));
                    openEditModal(rowData);
                });
            });
            
            document.querySelectorAll('.delete-btn').forEach(btn => {
                btn.addEventListener('click', function() {
                    const rowData = JSON.parse(this.getAttribute('data-row-data'));
                    handleDelete(rowData);
                });
            });
        } else {
            tableHead.innerHTML = '';
            tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
        }
    } catch (error) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
    }
}

// Вибір таблиці
async function selectTable(tableName) {
    currentTable = tableName;
    
    // Оновлюємо форму додавання
    updateInsertForm(tableName);
}

// Завантаження інформації про таблицю
async function loadTableInfo(tableName) {
    try {
        const response = await fetch(`/api/table-info/${tableName}`);
        const data = await response.json();

        if (data.success) {
            currentTableInfo = data;
        }
    } catch (error) {
        console.error('Помилка завантаження інформації про таблицю:', error);
    }
}


// Відкриття модального вікна для редагування/додавання
async function openEditModal(rowData) {
    editingRow = rowData;
    const modal = document.getElementById('editModal');
    const modalTitle = document.getElementById('editModalTitle');
    const editFields = document.getElementById('editFields');
    
    modalTitle.textContent = rowData ? 'Редагувати запис' : 'Додати новий запис';
    editFields.innerHTML = '';

    // Спеціальна форма для пункту "Структура НРК"
    if (currentMenuItem === 'Структура НРК') {
        // 1) l1 (readonly)
        const f1 = document.createElement('div');
        f1.className = 'form-group';
        const l1Label = document.createElement('label');
        l1Label.textContent = 'l1';
        const l1Input = document.createElement('input');
        l1Input.type = 'text';
        l1Input.name = 'l1';
        l1Input.readOnly = true;
        l1Input.value = rowData && rowData.l1 != null ? rowData.l1 : '';
        f1.appendChild(l1Label);
        f1.appendChild(l1Input);
        editFields.appendChild(f1);

        // 2) l2 (readonly)
        const f2 = document.createElement('div');
        f2.className = 'form-group';
        const l2Label = document.createElement('label');
        l2Label.textContent = 'l2';
        const l2Input = document.createElement('input');
        l2Input.type = 'text';
        l2Input.name = 'l2';
        l2Input.readOnly = true;
        l2Input.value = rowData && rowData.l2 != null ? rowData.l2 : '';
        f2.appendChild(l2Label);
        f2.appendChild(l2Input);
        editFields.appendChild(f2);

        // 3) type_bps (редактируемое текстовое поле)
        const f3 = document.createElement('div');
        f3.className = 'form-group';
        const typeLabel = document.createElement('label');
        typeLabel.textContent = 'type_bps';
        const typeInput = document.createElement('input');
        typeInput.type = 'text';
        typeInput.name = 'type_bps';
        typeInput.value = rowData && rowData.type_bps != null ? rowData.type_bps : '';
        f3.appendChild(typeLabel);
        f3.appendChild(typeInput);
        editFields.appendChild(f3);

        // 4) unit_structure_id (select по us.name, сохраняем id)
        const f4 = document.createElement('div');
        f4.className = 'form-group';
        const usLabel = document.createElement('label');
        usLabel.textContent = 'Структура підрозділу';
        const usSelect = document.createElement('select');
        usSelect.name = 'unit_structure_id';
        usSelect.className = 'form-select';

        // Пустая опция (по желанию)
        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.textContent = '-- Виберіть структуру --';
        usSelect.appendChild(emptyOption);

        try {
            // Загружаем варианты из unit_structure
            const resp = await fetch('/api/table-data/unit_structure');
            const result = await resp.json();
            if (result.success && Array.isArray(result.data)) {
                result.data.forEach(us => {
                    const opt = document.createElement('option');
                    opt.value = us.id;        // в БД пойдет id
                    opt.textContent = us.name; // в списке показываем name
                    usSelect.appendChild(opt);
                });
            }
        } catch (e) {
            console.error('Помилка завантаження unit_structure для Структура НРК:', e);
        }

        // Текущее значение unit_structure_id из строки
        if (rowData && rowData.unit_structure_id != null) {
            usSelect.value = rowData.unit_structure_id;
        }

        f4.appendChild(usLabel);
        f4.appendChild(usSelect);
        editFields.appendChild(f4);

        modal.style.display = 'block';
        return;
    }

    if (!currentTableInfo || !currentTableInfo.columns) {
        showModal('Помилка: інформація про таблицю не завантажена', 'error');
        return;
    }

    // Загружаем дані для випадаючих списків
    let rankData = [];
    let positionData = [];
    let unitData = [];
    let unitStructureData = [];
    let materialGroupData = [];
    let manufactureData = [];
    const hasUnitId = currentTableInfo.columns.some(c => c.name.toLowerCase() === 'unit_id');

    if (currentTable === 'contact_person') {
        try {
            const rankResponse = await fetch('/api/table-data/rank');
            const rankResult = await rankResponse.json();
            if (rankResult.success) {
                rankData = rankResult.data;
            }
            
            const positionResponse = await fetch('/api/table-data/position');
            const positionResult = await positionResponse.json();
            if (positionResult.success) {
                positionData = positionResult.data;
            }
            
            const unitResponse = await fetch('/api/table-data/unit');
            const unitResult = await unitResponse.json();
            if (unitResult.success) {
                unitData = unitResult.data;
            }
        } catch (error) {
            console.error('Помилка завантаження даних для випадаючих списків:', error);
        }
    } else if (hasUnitId) {
        try {
            const unitResponse = await fetch('/api/table-data/unit');
            const unitResult = await unitResponse.json();
            if (unitResult.success) {
                unitData = unitResult.data;
            }
        } catch (error) {
            console.error('Помилка завантаження даних unit для випадаючого списку:', error);
        }
    }
    
    // Загружаем дані для unit_structure для таблиці subordination
    if (currentTable === 'subordination') {
        try {
            const unitStructureResponse = await fetch('/api/table-data/unit_structure');
            const unitStructureResult = await unitStructureResponse.json();
            if (unitStructureResult.success) {
                unitStructureData = unitStructureResult.data;
            }
        } catch (error) {
            console.error('Помилка завантаження даних для unit_structure:', error);
        }
    }

    // Загружаем дані для довідників груп засобів та виробників для таблиці Material
    if (currentTable === 'Material') {
        try {
            const [groupResp, manufactureResp] = await Promise.all([
                fetch('/api/table-data/material_group'),
                fetch('/api/table-data/manufacture')
            ]);

            const groupResult = await groupResp.json();
            if (groupResult.success) {
                materialGroupData = groupResult.data || [];
            }

            const manufactureResult = await manufactureResp.json();
            if (manufactureResult.success) {
                manufactureData = manufactureResult.data || [];
            }

            // Якщо редагуємо існуючий матеріал і в нас ще немає group_id у rowData,
            // можна в подальшому тут додати окреме завантаження поточної групи з бекенду,
            // якщо це буде потрібно.
        } catch (error) {
            console.error('Помилка завантаження груп засобів / виробників:', error);
        }
    }

    // Создаем поля формы
    for (const col of currentTableInfo.columns) {
        // Пропускаємо автоінкрементні поля при додаванні
        if (!rowData && col.name.toLowerCase() === 'id' && col.type.toLowerCase().includes('int')) {
            continue;
        }
        
        // Пропускаємо поле id при редагуванні
        if (rowData && col.name.toLowerCase() === 'id') {
            continue;
        }
        
        // Пропускаємо поле dt при редагуванні
        if (rowData && col.name.toLowerCase() === 'dt') {
            continue;
        }

        const div = document.createElement('div');
        div.className = 'form-group';
        
        const label = document.createElement('label');
        label.textContent = `${col.name} (${col.type}${col.nullable === 'YES' ? ', nullable' : ''})`;
        
        // Для полів rank_id, position_id, unit_id, unit_structure_id, group_id, manufacture_id створюємо select
        if (col.name.toLowerCase() === 'rank_id' && rankData.length > 0) {
            const select = document.createElement('select');
            select.name = col.name;
            select.required = col.nullable === 'NO' && !col.default && !rowData;
            select.className = 'form-select';
            
            // Добавляем пустую опцию если поле nullable
            if (col.nullable === 'YES') {
                const emptyOption = document.createElement('option');
                emptyOption.value = '';
                emptyOption.textContent = '-- Виберіть звання --';
                select.appendChild(emptyOption);
            }
            
            // Заполняем опции из rankData
            rankData.forEach(rank => {
                const option = document.createElement('option');
                option.value = rank.id;
                option.textContent = rank.name;
                select.appendChild(option);
            });
            
            // Устанавливаем выбранное значение при редактировании
            if (rowData && rowData[col.name] !== null && rowData[col.name] !== undefined) {
                select.value = rowData[col.name];
            }
            
            div.appendChild(label);
            div.appendChild(select);
            editFields.appendChild(div);
        } else if (col.name.toLowerCase() === 'position_id' && positionData.length > 0) {
            const select = document.createElement('select');
            select.name = col.name;
            select.required = col.nullable === 'NO' && !col.default && !rowData;
            select.className = 'form-select';
            
            // Добавляем пустую опцию если поле nullable
            if (col.nullable === 'YES') {
                const emptyOption = document.createElement('option');
                emptyOption.value = '';
                emptyOption.textContent = '-- Виберіть посаду --';
                select.appendChild(emptyOption);
            }
            
            // Заполняем опции из positionData
            positionData.forEach(position => {
                const option = document.createElement('option');
                option.value = position.id;
                option.textContent = position.name;
                select.appendChild(option);
            });
            
            // Устанавливаем выбранное значение при редактировании
            if (rowData && rowData[col.name] !== null && rowData[col.name] !== undefined) {
                select.value = rowData[col.name];
            }
            
            div.appendChild(label);
            div.appendChild(select);
            editFields.appendChild(div);
        } else if (col.name.toLowerCase() === 'unit_id' && unitData.length > 0) {
            const select = document.createElement('select');
            select.name = col.name;
            select.required = col.nullable === 'NO' && !col.default && !rowData;
            select.className = 'form-select';
            
            if (col.nullable === 'YES') {
                const emptyOption = document.createElement('option');
                emptyOption.value = '';
                emptyOption.textContent = '-- Виберіть підрозділ --';
                select.appendChild(emptyOption);
            }
            
            // У всіх формах: список unit.off_name (value = id)
            unitData.forEach(unit => {
                const option = document.createElement('option');
                option.value = unit.id;
                option.textContent = (unit.off_name !== undefined && unit.off_name !== null ? unit.off_name : unit.name || '').toString().trim() || '(порожньо)';
                select.appendChild(option);
            });
             
            if (rowData && rowData[col.name] !== null && rowData[col.name] !== undefined) {
                select.value = rowData[col.name];
            }
            
            div.appendChild(label);
            div.appendChild(select);
            editFields.appendChild(div);
        } else if (col.name.toLowerCase() === 'unit_structure_id' && unitStructureData.length > 0) {
            const select = document.createElement('select');
            select.name = col.name;
            select.required = col.nullable === 'NO' && !col.default && !rowData;
            select.className = 'form-select';
            
            
            if (col.nullable === 'YES') {
                const emptyOption = document.createElement('option');
                emptyOption.value = '';
                emptyOption.textContent = '-- Виберіть структуру підрозділу --';
                select.appendChild(emptyOption);
            }
            
            
            unitStructureData.forEach(unitStructure => {
                const option = document.createElement('option');
                option.value = unitStructure.id;
                option.textContent = unitStructure.name;
                select.appendChild(option);
            });
            
            // Устанавливаем выбранное значение при редактировании
            if (rowData && rowData[col.name] !== null && rowData[col.name] !== undefined) {
                select.value = rowData[col.name];
            }
            
            div.appendChild(label);
            div.appendChild(select);
            editFields.appendChild(div);
        } else if (col.name.toLowerCase() === 'manufacture_id' && manufactureData.length > 0) {
            // Вибір виробника з таблиці manufacture
            const select = document.createElement('select');
            select.name = col.name;
            select.required = col.nullable === 'NO' && !col.default && !rowData;
            select.className = 'form-select';

            if (col.nullable === 'YES') {
                const emptyOption = document.createElement('option');
                emptyOption.value = '';
                emptyOption.textContent = '-- Виберіть виробника --';
                select.appendChild(emptyOption);
            }

            manufactureData.forEach(m => {
                const option = document.createElement('option');
                option.value = m.id;
                option.textContent = m.name;
                select.appendChild(option);
            });

            if (rowData && rowData[col.name] !== null && rowData[col.name] !== undefined) {
                select.value = rowData[col.name];
            }

            div.appendChild(label);
            div.appendChild(select);
            editFields.appendChild(div);
        } else {
            // Для остальных полей создаем обычный input
            const input = document.createElement('input');
            input.type = 'text';
            input.name = col.name;
            input.required = col.nullable === 'NO' && !col.default && !rowData;
            
            // Заповнюємо значеннями при редагуванні
            if (rowData && rowData[col.name] !== null && rowData[col.name] !== undefined) {
                input.value = rowData[col.name];
            } else {
                input.placeholder = col.default ? `За замовчуванням: ${col.default}` : '';
            }

            div.appendChild(label);
            div.appendChild(input);
            editFields.appendChild(div);
        }
    }

    // Окреме поле "Група засобу" для таблиці Material (звʼязок через material_group_rel, а не пряме поле таблиці)
    if (currentTable === 'Material' && materialGroupData.length > 0) {
        const div = document.createElement('div');
        div.className = 'form-group';

        const label = document.createElement('label');
        label.textContent = 'Група засобу';

        const select = document.createElement('select');
        // Використовуємо спеціальне імʼя поля, яке не існує у таблиці Material,
        // щоб не намагатися оновлювати неіснуючу колонку в БД.
        select.name = 'material_group_id';
        select.className = 'form-select';

        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.textContent = '-- Виберіть групу засобу --';
        select.appendChild(emptyOption);

        materialGroupData.forEach(group => {
            const option = document.createElement('option');
            option.value = group.id;
            option.textContent = group.name;
            select.appendChild(option);
        });

        // Встановлюємо поточну групу за назвою, якщо є (group_name у рядку НРК)
        if (rowData && rowData.group_name) {
            const found = materialGroupData.find(g => g.name === rowData.group_name);
            if (found) {
                select.value = String(found.id);
            }
        }

        div.appendChild(label);
        div.appendChild(select);
        editFields.appendChild(div);
    }

    modal.style.display = 'block';
}

// Закриття модального вікна редагування
function closeEditModal() {
    document.getElementById('editModal').style.display = 'none';
    editingRow = null;
    document.getElementById('editForm').reset();
}

// Обробка збереження даних
async function handleSave(e) {
    e.preventDefault();

    if (!currentTable) {
        showModal('Помилка: таблиця не вибрана', 'error');
        return;
    }

    // Окрема логіка збереження для пункту "Структура НРК"
    if (currentMenuItem === 'Структура НРК') {
        try {
            const form = document.getElementById('editForm');
            const formData = new FormData(form);

            const type_bps = formData.get('type_bps');
            const unit_structure_id = formData.get('unit_structure_id');

            if (!editingRow || !editingRow.id) {
                showModal('Помилка: відсутній ідентифікатор unit_id для оновлення структури НРК', 'error');
                return;
            }

            const payload = {
                unit_id: editingRow.id,
                type_bps: type_bps !== '' ? type_bps : null,
                unit_structure_id: unit_structure_id !== '' ? unit_structure_id : null
            };

            const response = await fetch('/api/update-nrc-structure', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            const data = await response.json();

            if (data.success) {
                showModal(data.message || 'Структуру НРК оновлено', 'success');

                // Закриваємо форму після успішного збереження
                closeEditModal();

                // Оновлюємо таблицю, зберігаючи позицію скролу
                const tableContainer = document.querySelector('.table-container');
                const scrollPosition = tableContainer ? tableContainer.scrollTop : 0;
                window.savedScrollPosition = scrollPosition;

                await loadNRCStructureQuery();
            } else {
                showModal('Помилка: ' + (data.error || 'невідома помилка при оновленні структури НРК'), 'error');
            }
        } catch (error) {
            showModal('Помилка: ' + error.message, 'error');
        }
        return;
    }

    const formData = new FormData(e.target);
    const values = {};
    const where = {};

    formData.forEach((value, key) => {
        const trimmed = value.trim();

        // Пропускаємо поле id, воно буде використано з editingRow
        if (key.toLowerCase() === 'id') {
            return;
        }

        // Для НРК (Material) поле material_group_id є службовим для звʼязку,
        // воно не існує безпосередньо в таблиці Material, тому не відправляємо його у INSERT/UPDATE
        if (currentTable === 'Material' && key.toLowerCase() === 'material_group_id') {
            return;
        }

        if (editingRow) {
            // При редагуванні ми хочемо мати можливість очищати поля,
            // тому навіть порожні значення відправляємо в запит.
            // Порожній рядок інтерпретуємо як NULL на рівні БД.
            values[key] = trimmed === '' ? null : trimmed;
        } else {
            // При створенні нового запису пропускаємо порожні поля
            if (trimmed !== '') {
                values[key] = trimmed;
            }
        }
    });

    try {
        let response;
        if (editingRow) {
            // Оновлення існуючого запису
            // Використовуємо ID з editingRow для WHERE умови
            if (editingRow.id) {
                where.id = editingRow.id;
            } else if (currentTable === 'Material') {
                // Якщо немає ID, використовуємо name та off_name для ідентифікації
                if (editingRow.name) where.name = editingRow.name;
                if (editingRow.off_name) where.off_name = editingRow.off_name;
            }
            
            // Перевіряємо, чи є умови для WHERE
            if (Object.keys(where).length === 0) {
                showModal('Помилка: не вдалося визначити запис для оновлення', 'error');
                return;
            }
            
            // Для поля dt автоматично встановлюємо поточну дату/час при редагуванні
            if (currentTableInfo && currentTableInfo.columns) {
                const dtColumn = currentTableInfo.columns.find(col => col.name.toLowerCase() === 'dt');
                if (dtColumn) {
                    const now = new Date();
                    const year = now.getFullYear();
                    const month = String(now.getMonth() + 1).padStart(2, '0');
                    const day = String(now.getDate()).padStart(2, '0');
                    const hours = String(now.getHours()).padStart(2, '0');
                    const minutes = String(now.getMinutes()).padStart(2, '0');
                    const seconds = String(now.getSeconds()).padStart(2, '0');
                    values[dtColumn.name] = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
                }
            }
            
            response = await fetch('/api/update', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    table: currentTable,
                    values: values,
                    where: where
                })
            });
        } else {
            // Додавання нового запису
            response = await fetch('/api/insert', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    table: currentTable,
                    values: values
                })
            });
        }

        const data = await response.json();

        if (data.success) {
            // Зберігаємо поточні фільтри, щоб не скидати їх після оновлення даних
            window.savedSearchFilters = getSearchFilters();

            // Якщо редагуємо засіб (Material) — окремо оновлюємо його групу через material_group_rel
            if (editingRow && currentTable === 'Material') {
                try {
                    const formEl = document.getElementById('editForm');
                    const groupSelect =
                        formEl.querySelector('select[name="group_id"]') ||
                        formEl.querySelector('select[name="material_group_id"]');

                    if (groupSelect && editingRow.id) {
                        const groupVal = groupSelect.value;
                        await fetch('/api/update-material-group', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json'
                            },
                            body: JSON.stringify({
                                material_id: editingRow.id,
                                group_id: groupVal === '' ? null : groupVal
                            })
                        });
                    }
                } catch (e) {
                    console.error('Помилка оновлення групи засобу:', e);
                }
            }

            // Успішне збереження без показу повідомлення
            closeEditModal();
            
            // Сохраняем позицию скролла перед обновлением
            const tableContainer = document.querySelector('.table-container');
            const scrollPosition = tableContainer ? tableContainer.scrollTop : 0;
            window.savedScrollPosition = scrollPosition; // Сохраняем в глобальной переменной
            
            // Оновлюємо дані таблиці
            if (currentMenuItem === 'ВО загальні') {
                await loadVOQuery();
            } else if (currentMenuItem === 'ВО нрк') {
                await loadVONRKQuery();
            } else if (currentMenuItem === 'Загальна') {
                await loadCustomQuery();
            } else if (currentMenuItem === 'НРК') {
                await loadNRCQuery();
            } else if (currentMenuItem === 'Структура НРК') {
                await loadNRCStructureQuery();
            } else if (currentTable) {
                await loadTableDataForMenu(currentTable);
            } else {
                // Якщо це запит без таблиці, перезавантажуємо поточний запит
                if (currentMenuItem === 'НРК') {
                    await loadNRCQuery();
                } else if (currentMenuItem === 'Структура НРК') {
                    await loadNRCStructureQuery();
                } else if (currentMenuItem === 'ВО нрк') {
                    await loadVONRKQuery();
                }
            }
        } else {
            showModal('Помилка: ' + data.error, 'error');
        }
    } catch (error) {
        showModal('Помилка: ' + error.message, 'error');
    }
}

// Обробка видалення даних
async function handleDelete(rowData) {
    if (!currentTable) {
        showModal('Помилка: таблиця не вибрана', 'error');
        return;
    }

    // Підтвердження видалення
    if (!confirm('Ви впевнені, що хочете видалити цей запис?')) {
        return;
    }

    try {
        // Формуємо умови WHERE для видалення
        const whereConditions = {};
        
        // Для таблиці Material використовуємо id, або name + off_name якщо id немає
        if (currentTable === 'Material') {
            if (rowData.id) {
                whereConditions.id = rowData.id;
            } else {
                if (rowData.name) whereConditions.name = rowData.name;
                if (rowData.off_name) whereConditions.off_name = rowData.off_name;
            }
        } else {
            // Для інших таблиць використовуємо id якщо він є
            if (rowData.id) {
                whereConditions.id = rowData.id;
            } else {
                // Якщо немає id, намагаємося використати перше доступне поле
                const firstKey = Object.keys(rowData)[0];
                if (firstKey) {
                    whereConditions[firstKey] = rowData[firstKey];
                }
            }
        }

        if (Object.keys(whereConditions).length === 0) {
            showModal('Помилка: недостатньо даних для видалення', 'error');
            return;
        }

        const response = await fetch('/api/delete', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                table: currentTable,
                where: whereConditions
            })
        });

        const data = await response.json();

        if (data.success) {
            showModal(data.message || 'Дані успішно видалено', 'success');
            
            // Сохраняем позицию скролла перед обновлением
            const tableContainer = document.querySelector('.table-container');
            const scrollPosition = tableContainer ? tableContainer.scrollTop : 0;
            window.savedScrollPosition = scrollPosition; // Сохраняем в глобальной переменной
            
            // Оновлюємо дані таблиці
            if (currentMenuItem === 'ВО загальні') {
                await loadVOQuery();
            } else if (currentMenuItem === 'ВО нрк') {
                await loadVONRKQuery();
            } else if (currentMenuItem === 'Загальна') {
                await loadCustomQuery();
            } else if (currentMenuItem === 'НРК') {
                await loadNRCQuery();
            } else if (currentMenuItem === 'Структура НРК') {
                await loadNRCStructureQuery();
            } else if (currentTable) {
                await loadTableDataForMenu(currentTable);
            }
        } else {
            showModal('Помилка: ' + (data.error || 'Не вдалося видалити дані'), 'error');
        }
    } catch (error) {
        showModal('Помилка: ' + error.message, 'error');
    }
}

// Обробка додавання даних
async function handleInsert(e) {
    e.preventDefault();

    if (!currentTable) {
        showModal('Виберіть таблицю', 'error');
        return;
    }

    const formData = new FormData(e.target);
    const values = {};

    formData.forEach((value, key) => {
        if (value.trim() !== '') {
            values[key] = value.trim();
        }
    });

    if (Object.keys(values).length === 0) {
        showModal('Заповніть хоча б одне поле', 'error');
        return;
    }

    try {
        const response = await fetch('/api/insert', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                table: currentTable,
                values: values
            })
        });

        const data = await response.json();

        if (data.success) {
            showModal(data.message, 'success');
            e.target.reset();
            // Оновлюємо інформацію про таблицю
            await loadTableInfo(currentTable);
        } else {
            showModal('Помилка: ' + data.error, 'error');
        }
    } catch (error) {
        showModal('Помилка: ' + error.message, 'error');
    }
}

// Завантаження довідника засобів в нову таблицю з Excel файлу
async function loadMaterialsToNewTable() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    tableHead.innerHTML = '';
    tableBody.innerHTML = '<tr><td colspan="100" class="loading">Очікування вибору файлу...</td></tr>';

    try {
        // Створюємо input для вибору файлу
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.xlsx,.xls';
        fileInput.style.display = 'none';
        document.body.appendChild(fileInput);

        // Відкриваємо діалог вибору файлу
        fileInput.click();

        fileInput.onchange = async function(e) {
            const file = e.target.files[0];
            if (!file) {
                tableBody.innerHTML = '<tr><td colspan="100" style="color: #999; padding: 20px;">Операцію скасовано</td></tr>';
                document.body.removeChild(fileInput);
                return;
            }

            // Перевіряємо розширення файлу
            const fileName = file.name;
            const fileExtension = fileName.split('.').pop().toLowerCase();
            if (!['xlsx', 'xls'].includes(fileExtension)) {
                showModal('Помилка: Оберіть файл Excel (.xlsx або .xls)', 'error');
                tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: Оберіть файл Excel (.xlsx або .xls)</td></tr>';
                document.body.removeChild(fileInput);
                return;
            }

            tableBody.innerHTML = '<tr><td colspan="100" class="loading">Завантаження файлу...</td></tr>';

            // Створюємо FormData для відправки файлу
            const formData = new FormData();
            formData.append('file', file);

            // Отримуємо ім'я таблиці з імені файлу (без розширення)
            const tableName = fileName.replace(/\.[^/.]+$/, '').replace(/[^a-zA-Z0-9_]/g, '_');

            // Викликаємо API для завантаження Excel файлу
            const response = await fetch(`/api/upload-excel?table_name=${encodeURIComponent(tableName)}`, {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (data.success) {
                const successMessage = `Дані успішно завантажено в таблицю "${tableName}"!\n\n` +
                    `Завантажено записів: ${data.rows_loaded}\n` +
                    `Кількість колонок: ${data.columns_count}`;
                showModal(successMessage, 'success');
                tableBody.innerHTML = '<tr><td colspan="100" style="color: #28a745; padding: 20px; text-align: center;">' +
                    successMessage.replace(/\n/g, '<br>') + '</td></tr>';
            } else {
                showModal('Помилка: ' + data.error, 'error');
                tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + data.error + '</td></tr>';
            }

            document.body.removeChild(fileInput);
        };
    } catch (error) {
        showModal('Помилка: ' + error.message, 'error');
        tableBody.innerHTML = '<tr><td colspan="100" style="color: #dc3545; padding: 20px;">Помилка: ' + error.message + '</td></tr>';
    }
}

// Показати модальне вікно
function showModal(message, type = 'success') {
    const modal = document.getElementById('modal');
    const modalMessage = document.getElementById('modalMessage');
    
    modalMessage.className = type === 'success' ? 'message-success' : 'message-error';
    modalMessage.textContent = message;
    modal.style.display = 'block';
}

// Закрити модальне вікно
function closeModal() {
    document.getElementById('modal').style.display = 'none';
}
