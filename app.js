/**
         * yasicalc Core Logic
         * - Single State -> Render Loop
         * - Cycle Detection
         * - LocalStorage Persistence
         * - Phase 2: Sign Switching & Sheet Reordering
         */

        const Utils = {
            generateId: () => Math.random().toString(36).substr(2, 9),
            formatNumber: (num) => new Intl.NumberFormat('ja-JP').format(num)
        };

        const INITIAL_STATE = {
            unit: '円',
            books: [
                {
                    id: 'book_1',
                    name: 'Book 1',
                    sheets: [
                        {
                            id: 'sheet_1',
                            name: 'Sheet 1',
                            totalViewMode: 'total',
                            splitCount: 1,
                            rows: [
                                { id: 'row_1', label: '項目 A', type: 'value', value: 0, refSheetId: null, operator: 1 },
                                { id: 'row_2', label: '項目 B', type: 'value', value: 0, refSheetId: null, operator: 1 }
                            ]
                        }
                    ]
                }
            ],
            view: 'sheet', // 'sheet' or 'summary'
            activeBookId: 'book_1',
            activeSheetId: 'sheet_1'
        };

        class App {
            constructor() {
                this.state = this.loadState() || JSON.parse(JSON.stringify(INITIAL_STATE));
                // Migrate old state if needed (add operator if missing)
                this.migrateState();
                this.cache = {};
                this.dragSrcEl = null; // For D&D
                this.sidebarOpen = false;
                this.init();
            }

            migrateState() {
                if (!this.state.unit) {
                    this.state.unit = '円';
                }
                if (!this.state.view || (this.state.view !== 'sheet' && this.state.view !== 'summary')) {
                    this.state.view = 'sheet';
                }

                // Legacy migration: state.sheets -> state.books[0].sheets
                if (!Array.isArray(this.state.books)) {
                    const legacySheets = Array.isArray(this.state.sheets) ? this.state.sheets : [];
                    const bookId = Utils.generateId();
                    this.state.books = [{
                        id: bookId,
                        name: 'Book 1',
                        sheets: legacySheets.length ? legacySheets : [this.createDefaultSheet('Sheet 1')]
                    }];
                    this.state.activeBookId = bookId;
                }

                if (!this.state.books.length) {
                    this.state.books = [this.createDefaultBook('Book 1')];
                }

                this.state.books = this.state.books.map((book, bookIndex) => {
                    const normalizedBook = {
                        id: book?.id || Utils.generateId(),
                        name: book?.name || `Book ${bookIndex + 1}`,
                        sheets: Array.isArray(book?.sheets) ? book.sheets : []
                    };

                    if (!normalizedBook.sheets.length) {
                        normalizedBook.sheets.push(this.createDefaultSheet('Sheet 1'));
                    }

                    normalizedBook.sheets = normalizedBook.sheets.map((sheet, sheetIndex) => {
                        const normalizedSheet = {
                            id: sheet?.id || Utils.generateId(),
                            name: sheet?.name || `Sheet ${sheetIndex + 1}`,
                            totalViewMode: sheet?.totalViewMode === 'perPerson' ? 'perPerson' : 'total',
                            splitCount: this.normalizeSplitCount(sheet?.splitCount),
                            rows: Array.isArray(sheet?.rows) ? sheet.rows : []
                        };

                        normalizedSheet.rows = normalizedSheet.rows.map((row) => ({
                            id: row?.id || Utils.generateId(),
                            label: row?.label ?? '',
                            type: row?.type === 'ref' ? 'ref' : 'value',
                            value: row?.value ?? 0,
                            refSheetId: row?.refSheetId ?? null,
                            operator: row?.operator === -1 ? -1 : 1
                        }));

                        return normalizedSheet;
                    });

                    // Remove invalid references inside each book
                    const sheetIds = new Set(normalizedBook.sheets.map((s) => s.id));
                    normalizedBook.sheets.forEach((sheet) => {
                        sheet.rows.forEach((row) => {
                            if (row.type === 'ref' && row.refSheetId && !sheetIds.has(row.refSheetId)) {
                                row.refSheetId = null;
                            }
                        });
                    });

                    return normalizedBook;
                });

                // Pointer repair
                const fallbackBook = this.state.books[0];
                if (!this.state.activeBookId || !this.getBook(this.state.activeBookId)) {
                    this.state.activeBookId = fallbackBook.id;
                }

                const activeBook = this.getActiveBook();
                const activeSheets = activeBook?.sheets || [];
                if (!this.state.activeSheetId || !activeSheets.find((s) => s.id === this.state.activeSheetId)) {
                    this.state.activeSheetId = activeSheets[0]?.id || null;
                }

                delete this.state.sheets;
            }

            createDefaultRow(label = '新規項目') {
                return {
                    id: Utils.generateId(),
                    label,
                    type: 'value',
                    value: 0,
                    refSheetId: null,
                    operator: 1
                };
            }

            createDefaultSheet(name = 'Sheet 1') {
                return {
                    id: Utils.generateId(),
                    name,
                    totalViewMode: 'total',
                    splitCount: 1,
                    rows: [this.createDefaultRow('項目 A')]
                };
            }

            createDefaultBook(name = 'Book 1') {
                return {
                    id: Utils.generateId(),
                    name,
                    sheets: [this.createDefaultSheet('Sheet 1')]
                };
            }

            getBook(bookId) {
                return this.state.books.find((book) => book.id === bookId) || null;
            }

            getSheet(sheetId, bookId = this.state.activeBookId) {
                return this.getBook(bookId)?.sheets.find((sheet) => sheet.id === sheetId) || null;
            }

            getActiveBook() {
                return this.getBook(this.state.activeBookId);
            }

            getActiveSheets() {
                return this.getActiveBook()?.sheets || [];
            }

            init() {
                // Initial Bindings
                const unitInput = document.getElementById('unitInput');
                unitInput.value = this.state.unit;
                this.adjustUnitInputWidth(this.state.unit);
                unitInput.addEventListener('input', (e) => {
                    this.state.unit = e.target.value;
                    this.adjustUnitInputWidth(e.target.value);
                    this.saveAndRender();
                });

                this.applySidebarState();
                window.addEventListener('resize', () => {
                    if (!window.matchMedia('(max-width: 720px)').matches && this.sidebarOpen) {
                        this.toggleSidebar(false);
                    }
                    this.adjustUnitInputWidth(unitInput.value);
                });

                this.render();
            }

            adjustUnitInputWidth(value) {
                const unitInput = document.getElementById('unitInput');
                if (!unitInput) return;

                const text = String(value || '').trim() || '円';
                const minWidth = window.matchMedia('(max-width: 720px)').matches ? 72 : 64;
                const width = Math.min(96, Math.max(minWidth, 24 + text.length * 16));
                unitInput.style.width = `${width}px`;
            }

            // --- Mobile Sidebar ---

            applySidebarState() {
                document.body.classList.toggle('sidebar-open', this.sidebarOpen);
            }

            toggleSidebar(force) {
                if (typeof force === 'boolean') {
                    this.sidebarOpen = force;
                } else {
                    this.sidebarOpen = !this.sidebarOpen;
                }
                this.applySidebarState();
            }

            closeSidebar() {
                if (this.sidebarOpen) {
                    this.toggleSidebar(false);
                }
            }

            maybeCloseSidebar() {
                if (window.matchMedia('(max-width: 720px)').matches) {
                    this.closeSidebar();
                }
            }

            // --- State Management ---

            saveState() {
                localStorage.setItem('yasicalc_data', JSON.stringify(this.state));
            }

            loadState() {
                const data = localStorage.getItem('yasicalc_data');
                return data ? JSON.parse(data) : null;
            }

            saveAndRender() {
                this.saveState();
                const unitInput = document.getElementById('unitInput');
                if (unitInput && unitInput.value !== this.state.unit) {
                    unitInput.value = this.state.unit;
                    this.adjustUnitInputWidth(this.state.unit);
                }
                this.render();
            }

            // --- Calculation Engine ---

            normalizeSplitCount(value) {
                const numeric = Number.parseInt(String(value ?? '').replace(/[^\d]/g, ''), 10);
                return Number.isFinite(numeric) && numeric > 0 ? numeric : 1;
            }

            calculateSheetTotal(sheetId, visited = new Set(), bookId = this.state.activeBookId) {
                if (visited.has(sheetId)) {
                    return { value: 0, error: 'LOOP' };
                }

                const book = this.getBook(bookId);
                if (!book) return { value: 0, error: 'MISSING_BOOK' };

                const sheet = book.sheets.find(s => s.id === sheetId);
                if (!sheet) return { value: 0, error: 'MISSING' };

                let total = 0;
                let hasError = null;

                visited.add(sheetId);

                for (const row of sheet.rows) {
                    const op = row.operator !== undefined ? row.operator : 1;

                    if (row.type === 'value') {
                        total += (Number(row.value) || 0) * op;
                    } else if (row.type === 'ref') {
                        if (!row.refSheetId) continue;

                        // Recursive call
                        const result = this.calculateSheetTotal(row.refSheetId, new Set(visited), bookId);

                        if (result.error) {
                            hasError = result.error;
                        } else {
                            total += result.value * op;
                        }
                    }
                }

                visited.delete(sheetId);
                return { value: total, error: hasError };
            }

            getSheetTotalViewMode(sheet) {
                return sheet?.totalViewMode === 'perPerson' ? 'perPerson' : 'total';
            }

            getSheetTotalLabel(mode) {
                return mode === 'perPerson' ? '一人当たり' : '合計';
            }

            getSheetTotalDisplayResult(sheetId, bookId = this.state.activeBookId) {
                const sheet = this.getSheet(sheetId, bookId);
                const mode = this.getSheetTotalViewMode(sheet);
                const splitCount = this.normalizeSplitCount(sheet?.splitCount);
                const totalRes = this.calculateSheetTotal(sheetId, new Set(), bookId);

                if (totalRes.error || mode !== 'perPerson') {
                    return {
                        ...totalRes,
                        mode,
                        splitCount
                    };
                }

                return {
                    value: totalRes.value / splitCount,
                    error: null,
                    mode,
                    splitCount
                };
            }

            // --- Actions ---

            addBook() {
                const newBook = this.createDefaultBook(`Book ${this.state.books.length + 1}`);
                this.state.books.push(newBook);
                this.state.activeBookId = newBook.id;
                this.state.activeSheetId = newBook.sheets[0].id;
                this.state.view = 'sheet';
                this.saveAndRender();
                this.maybeCloseSidebar();
            }

            selectBook(bookId) {
                const book = this.getBook(bookId);
                if (!book) return;

                if (!book.sheets.length) {
                    book.sheets.push(this.createDefaultSheet('Sheet 1'));
                }

                this.state.activeBookId = book.id;
                this.state.activeSheetId = book.sheets[0].id;
                this.state.view = 'sheet';
                this.saveAndRender();
                this.maybeCloseSidebar();
            }

            renameBook(bookId = this.state.activeBookId) {
                const book = this.getBook(bookId);
                if (!book) return;

                const nextName = prompt('ブック名を入力してください。', book.name);
                if (nextName === null) return;

                const trimmed = nextName.trim();
                if (!trimmed) {
                    alert('ブック名は空にできません。');
                    return;
                }

                if (trimmed === book.name) return;
                book.name = trimmed;
                this.saveAndRender();
            }

            deleteBook(bookId = this.state.activeBookId) {
                if (this.state.books.length <= 1) {
                    alert('ブックは1つ以上必要です。');
                    return;
                }

                const book = this.getBook(bookId);
                if (!book) return;

                if (!confirm(`「${book.name}」を削除しても本当によろしいですか？`)) return;

                const index = this.state.books.findIndex((b) => b.id === bookId);
                this.state.books.splice(index, 1);

                const nextBook = this.state.books[Math.min(index, this.state.books.length - 1)];
                if (!nextBook.sheets.length) {
                    nextBook.sheets.push(this.createDefaultSheet('Sheet 1'));
                }

                this.state.activeBookId = nextBook.id;
                this.state.activeSheetId = nextBook.sheets[0].id;
                this.state.view = 'sheet';
                this.saveAndRender();
            }

            addSheet() {
                const book = this.getActiveBook();
                if (!book) return;

                const newId = Utils.generateId();
                const newSheet = {
                    id: newId,
                    name: `New Sheet`,
                    totalViewMode: 'total',
                    splitCount: 1,
                    rows: [
                        { id: Utils.generateId(), label: '新規項目', type: 'value', value: 0, refSheetId: null, operator: 1 }
                    ]
                };
                book.sheets.push(newSheet);
                this.state.activeSheetId = newId;
                this.state.view = 'sheet';
                this.saveAndRender();
                this.maybeCloseSidebar();
            }

            duplicateSheet(sheetId) {
                const book = this.getActiveBook();
                if (!book) return;

                const sheetToDuplicate = book.sheets.find(s => s.id === sheetId);
                if (!sheetToDuplicate) return;

                const newSheetId = Utils.generateId();
                // 行のidも新しく生成してディープコピーする
                const newRows = sheetToDuplicate.rows.map(row => ({
                    ...row,
                    id: Utils.generateId()
                }));

                const newSheet = {
                    id: newSheetId,
                    name: `${sheetToDuplicate.name} (コピー)`,
                    totalViewMode: this.getSheetTotalViewMode(sheetToDuplicate),
                    splitCount: this.normalizeSplitCount(sheetToDuplicate.splitCount),
                    rows: newRows
                };

                // コピー元シートの直後に挿入する
                const index = book.sheets.findIndex(s => s.id === sheetId);
                book.sheets.splice(index + 1, 0, newSheet);

                this.state.activeSheetId = newSheetId;
                this.saveAndRender();
            }

            deleteSheet(sheetId) {
                if (!confirm('このシートを削除しても本当によろしいですか？')) return;

                const book = this.getActiveBook();
                if (!book) return;

                book.sheets = book.sheets.filter(s => s.id !== sheetId);

                // Remove references to this sheet
                book.sheets.forEach(s => {
                    s.rows.forEach(r => {
                        if (r.refSheetId === sheetId) {
                            r.refSheetId = null;
                        }
                    });
                });

                if (this.state.activeSheetId === sheetId) {
                    this.state.activeSheetId = book.sheets[0]?.id || null;
                    if (!this.state.activeSheetId && book.sheets.length === 0) {
                        const fallbackSheet = this.createDefaultSheet('Sheet 1');
                        book.sheets.push(fallbackSheet);
                        this.state.activeSheetId = fallbackSheet.id;
                    }
                }
                this.saveAndRender();
            }

            selectSheet(sheetId) {
                this.state.activeSheetId = sheetId;
                this.state.view = 'sheet';
                this.render();
                this.maybeCloseSidebar();
            }

            selectSummary() {
                this.state.view = 'summary';
                this.render();
                this.maybeCloseSidebar();
            }

            updateSheetName(id, newName) {
                const sheet = this.getActiveSheets().find(s => s.id === id);
                if (sheet) {
                    sheet.name = newName;
                    this.saveAndRender();
                }
            }

            updateSheetTotalViewMode(sheetId, mode) {
                const sheet = this.getSheet(sheetId);
                if (!sheet) return;

                sheet.totalViewMode = mode === 'perPerson' ? 'perPerson' : 'total';
                this.saveAndRender();
            }

            updateSheetSplitCount(sheetId, value) {
                const sheet = this.getSheet(sheetId);
                if (!sheet) return;

                sheet.splitCount = this.normalizeSplitCount(value);
                this.saveAndRender();
            }

            addRow(sheetId) {
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);
                if (sheet) {
                    sheet.rows.push({
                        id: Utils.generateId(),
                        label: '',
                        type: 'value',
                        value: 0,
                        refSheetId: null,
                        operator: 1
                    });
                    this.saveAndRender();
                }
            }

            duplicateRow(sheetId, rowId) {
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);
                if (sheet) {
                    const rowToDuplicate = sheet.rows.find(r => r.id === rowId);
                    if (rowToDuplicate) {
                        const newRow = {
                            ...rowToDuplicate,
                            id: Utils.generateId()
                        };
                        const rowIndex = sheet.rows.findIndex(r => r.id === rowId);
                        sheet.rows.splice(rowIndex + 1, 0, newRow);
                        this.saveAndRender();
                    }
                }
            }

            deleteRow(sheetId, rowId) {
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);
                if (sheet) {
                    sheet.rows = sheet.rows.filter(r => r.id !== rowId);
                    this.saveAndRender();
                }
            }

            updateRow(sheetId, rowId, field, value) {
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);
                if (!sheet) return;
                const row = sheet.rows.find(r => r.id === rowId);
                if (row) {
                    row[field] = value;
                    this.saveAndRender();
                }
            }

            normalizeNumericInputChars(value) {
                return String(value ?? '')
                    // Convert full-width numerals and symbols used on mobile JP keyboards.
                    .replace(/[０-９]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
                    .replace(/[．。]/g, '.')
                    .replace(/，/g, ',')
                    .replace(/＋/g, '+')
                    .replace(/[－−ー]/g, '-');
            }

            normalizeAmountInput(value) {
                const cleaned = this.normalizeNumericInputChars(value)
                    .trim()
                    .replace(/[,\s]/g, '')
                    .replace(/^[-+]/, '');

                const dotIndex = cleaned.indexOf('.');
                if (dotIndex === -1) {
                    return cleaned.replace(/[^\d]/g, '');
                }

                const integerPart = cleaned.slice(0, dotIndex).replace(/[^\d]/g, '');
                const fractionPart = cleaned.slice(dotIndex + 1).replace(/[^\d]/g, '');

                if (!integerPart && !fractionPart) return '';
                if (!fractionPart) return integerPart || '0';

                return `${integerPart || '0'}.${fractionPart}`;
            }

            formatAmountInput(value) {
                const normalized = this.normalizeAmountInput(value);
                if (!normalized || normalized === '.') return '';

                if (!/^\d*(\.\d*)?$/.test(normalized)) {
                    return normalized;
                }

                const [intRaw, fracRaw] = normalized.split('.');
                const intDigits = (intRaw || '0').replace(/^0+(?=\d)/, '');
                const intWithCommas = intDigits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

                return fracRaw !== undefined ? `${intWithCommas}.${fracRaw}` : intWithCommas;
            }

            formatValueDisplayByOperator(value, op) {
                const normalized = this.normalizeAmountInput(value);
                const formatted = this.formatAmountInput(normalized);
                const numericValue = Number(normalized || 0);
                const showMinus = op === -1 && formatted !== '' && numericValue !== 0;
                return showMinus ? `-${formatted}` : formatted;
            }

            getCaretPosFromTokenCount(displayValue, tokenCount, preferAfterMinus = false) {
                if (tokenCount <= 0) {
                    if (preferAfterMinus && displayValue.startsWith('-')) return 1;
                    return 0;
                }

                let seen = 0;
                for (let i = 0; i < displayValue.length; i++) {
                    const ch = displayValue[i];
                    if ((ch >= '0' && ch <= '9') || ch === '.') {
                        seen += 1;
                        if (seen >= tokenCount) return i + 1;
                    }
                }
                return displayValue.length;
            }

            handleValueInput(inputEl, op, sheetId = null, rowId = null) {
                if (!inputEl) return;

                const rawValue = inputEl.value ?? '';
                const caret = typeof inputEl.selectionStart === 'number' ? inputEl.selectionStart : rawValue.length;
                const rawPrefix = rawValue.slice(0, caret);
                const tokenCount = this.normalizeAmountInput(rawPrefix).length;
                const formattedDisplay = this.formatValueDisplayByOperator(rawValue, op);

                // Selection/caret drags on mobile can fire input events.
                // If visual value is unchanged, keep browser-managed selection as-is.
                if (formattedDisplay === rawValue) {
                    if (sheetId && rowId) {
                        this.updateValueRow(sheetId, rowId, rawValue, false);
                    }
                    return;
                }

                inputEl.value = formattedDisplay;

                if (typeof inputEl.setSelectionRange === 'function') {
                    const nextCaret = this.getCaretPosFromTokenCount(
                        formattedDisplay,
                        tokenCount,
                        rawPrefix.startsWith('-')
                    );
                    inputEl.setSelectionRange(nextCaret, nextCaret);
                }

                if (sheetId && rowId) {
                    // Keep state in sync even when `change` does not fire on some mobile browsers.
                    this.updateValueRow(sheetId, rowId, inputEl.value, false);
                }
            }

            updateValueRow(sheetId, rowId, value, shouldRender = true) {
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);
                if (!sheet) return;

                const row = sheet.rows.find(r => r.id === rowId);
                if (!row) return;

                const normalized = this.normalizeAmountInput(value);
                row.value = normalized;

                if (shouldRender) {
                    this.saveAndRender();
                    return;
                }

                this.saveState();
                this.refreshSheetTotalDisplay(sheetId);
            }

            getAmountMarkup(value) {
                return `
                    <span class="total-value">${Utils.formatNumber(value)}</span>
                    <span class="total-unit">${this.state.unit}</span>
                `;
            }

            getSheetTotalMarkup(totalRes) {
                if (totalRes.error) {
                    return `
                        <span class="total-value text-error" style="font-size:1.5rem">循環参照</span>
                        <span class="total-unit">${this.state.unit}</span>
                    `;
                }
                return this.getAmountMarkup(totalRes.value);
            }

            refreshSheetTotalDisplay(sheetId) {
                if (this.state.view !== 'sheet' || this.state.activeSheetId !== sheetId) return;
                const totalAmountEl = document.querySelector('.sheet-footer-under-title .total-amount');
                if (!totalAmountEl) return;
                const totalRes = this.getSheetTotalDisplayResult(sheetId);
                totalAmountEl.innerHTML = this.getSheetTotalMarkup(totalRes);
            }

            toggleOperator(sheetId, rowId) {
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);
                if (!sheet) return;
                const row = sheet.rows.find(r => r.id === rowId);
                if (row) {
                    row.operator = (row.operator === 1) ? -1 : 1;
                    this.saveAndRender();
                }
            }

            // Drag & Drop for Sheets
            handleDragStart(e, index) {
                this.dragSrcIndex = index;
                e.dataTransfer.effectAllowed = 'move';
                e.target.style.opacity = '0.4';
            }

            handleDragOver(e) {
                if (e.preventDefault) {
                    e.preventDefault(); // Necessary. Allows us to drop.
                }
                e.dataTransfer.dropEffect = 'move';
                return false;
            }

            handleDragEnter(e) {
                e.target.closest('li').classList.add('over');
            }

            handleDragLeave(e) {
                e.target.closest('li').classList.remove('over');
            }

            handleDrop(e, dropIndex) {
                if (e.stopPropagation) {
                    e.stopPropagation(); // stops the browser from redirecting.
                }

                if (this.dragSrcIndex !== dropIndex) {
                    const sheets = this.getActiveSheets();
                    const [movedSheet] = sheets.splice(this.dragSrcIndex, 1);
                    sheets.splice(dropIndex, 0, movedSheet);
                    this.saveAndRender();
                }
                return false;
            }

            handleDragEnd(e) {
                e.target.style.opacity = '1';
                document.querySelectorAll('.sheet-item').forEach(item => {
                    item.classList.remove('over');
                });
            }


            // Export/Import
            exportData() {
                const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(this.state));
                const downloadAnchorNode = document.createElement('a');
                downloadAnchorNode.setAttribute("href", dataStr);
                downloadAnchorNode.setAttribute("download", "yasicalc_data.json");
                document.body.appendChild(downloadAnchorNode);
                downloadAnchorNode.click();
                downloadAnchorNode.remove();
            }

            importData(input) {
                const file = input.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const json = JSON.parse(e.target.result);
                        if (json.books || json.sheets) {
                            this.state = json;
                            this.migrateState(); // Ensure compatibility
                            this.saveAndRender();
                            alert('読み込みが完了しました。');
                        } else {
                            alert('データの形式が正しくありません。');
                        }
                    } catch (err) {
                        alert('ファイルが破損しています。');
                    }
                };
                reader.readAsText(file);
                input.value = '';
            }

            resetData() {
                if (confirm('すべてのデータを消去して初期状態に戻します。本当によろしいですか？')) {
                    this.state = JSON.parse(JSON.stringify(INITIAL_STATE));
                    this.saveAndRender();
                }
            }

            showHelp() {
                document.getElementById('helpModal').classList.add('show');
            }

            hideHelp() {
                document.getElementById('helpModal').classList.remove('show');
            }

            // --- Rendering ---

            render() {
                this.renderSidebar();
                this.renderMain();
            }

            renderSidebar() {
                const bookSelect = document.getElementById('bookSelect');
                const renameBookBtn = document.getElementById('renameBookBtn');
                const deleteBookBtn = document.getElementById('deleteBookBtn');
                const list = document.getElementById('sheetList');
                list.innerHTML = '';

                const books = this.state.books || [];
                if (bookSelect) {
                    bookSelect.innerHTML = books.map((book) =>
                        `<option value="${book.id}">${book.name}</option>`
                    ).join('');
                    bookSelect.value = this.state.activeBookId || books[0]?.id || '';
                }

                if (deleteBookBtn) {
                    deleteBookBtn.disabled = books.length <= 1;
                }

                if (renameBookBtn) {
                    renameBookBtn.disabled = books.length === 0;
                }

                const sheets = this.getActiveSheets();
                sheets.forEach((sheet, index) => {
                    const li = document.createElement('li');

                    li.className = `sheet-item ${this.state.view === 'sheet' && this.state.activeSheetId === sheet.id ? 'active' : ''}`;
                    li.draggable = true;

                    // D&D Events
                    li.ondragstart = (e) => this.handleDragStart(e, index);
                    li.ondragover = (e) => this.handleDragOver(e);
                    li.ondragenter = (e) => this.handleDragEnter(e);
                    li.ondragleave = (e) => this.handleDragLeave(e);
                    li.ondrop = (e) => this.handleDrop(e, index);
                    li.ondragend = (e) => this.handleDragEnd(e);

                    // Click to select (checking if not dragging would be ideal but simple click works fine here usually)
                    li.onclick = (e) => {
                        // Prevent selection when interacting with specific sub-elements if needed
                        this.selectSheet(sheet.id);
                    };

                    const total = this.calculateSheetTotal(sheet.id);
                    const totalDisplay = total.error ? 'Err' : Utils.formatNumber(total.value);

                    li.innerHTML = `
                        <div style="display:flex; align-items:center; gap:var(--space-2);">
                            <span style="color:var(--text-tertiary); cursor:grab;">⋮⋮</span>
                            <span class="sheet-name">${sheet.name}</span>
                        </div>
                        <span class="sheet-total">${totalDisplay} <span style="font-size:0.7em">${this.state.unit}</span></span>
                    `;
                    list.appendChild(li);
                });

                const summaryBtn = document.getElementById('summaryTab');
                if (this.state.view === 'summary') {
                    summaryBtn.classList.add('active');
                } else {
                    summaryBtn.classList.remove('active');
                }
            }

            renderMain() {
                const container = document.getElementById('mainEditor');
                container.innerHTML = '';

                if (this.state.view === 'summary') {
                    this.renderSummary(container);
                } else {
                    this.renderSheetEditor(container);
                }
            }

            renderSummary(container) {
                const activeBook = this.getActiveBook();
                const sheets = this.getActiveSheets();
                if (!activeBook) return;

                const grandTotal = sheets.reduce((acc, sheet) => {
                    const res = this.calculateSheetTotal(sheet.id);
                    return acc + (res.error ? 0 : res.value);
                }, 0);

                let html = `
                    <div class="sheet-header">
                        <h1 style="margin:0; font-size: 2rem;">${activeBook.name} / まとめで見る</h1>
                    </div>
                    <div class="summary-table-container">
                        <table class="summary-table">
                            <thead>
                                <tr>
                                    <th>シート名</th>
                                    <th style="text-align:right;">合計</th>
                                </tr>
                            </thead>
                            <tbody>
                `;

                sheets.forEach(sheet => {
                    const res = this.calculateSheetTotal(sheet.id);
                    const val = res.error ? `<span class="text-error">Error (${res.error})</span>` : Utils.formatNumber(res.value);
                    html += `
                        <tr>
                            <td style="font-weight: 500;">${sheet.name}</td>
                            <td style="text-align:right; font-weight: 600; font-family: var(--font-main);">
                                ${val} <span style="font-size:0.8em; color: var(--text-secondary)">${this.state.unit}</span>
                            </td>
                        </tr>
                    `;
                });

                html += `
                            </tbody>
                        </table>
                    </div>
                    <div class="sheet-footer">
                        <div>
                            <div class="total-label">総合計</div>
                            <div style="font-size: 0.85rem; color: var(--text-tertiary);">このブック内の合算</div>
                        </div>
                        <div class="total-amount">
                            ${this.getAmountMarkup(grandTotal)}
                        </div>
                    </div>
                `;

                container.innerHTML = html;
            }

            renderSheetEditor(container) {
                const sheetId = this.state.activeSheetId;
                const sheet = this.getActiveSheets().find(s => s.id === sheetId);

                if (!sheet) return;

                // Header
                const header = document.createElement('div');
                header.className = 'sheet-header sheet-header-tight';
                header.innerHTML = `
                    <input type="text" class="sheet-title-input" value="${sheet.name}" placeholder="シート名を入力..."
                        onchange="app.updateSheetName('${sheetId}', this.value)">
                    <div class="sheet-actions">
                        <button onclick="app.duplicateSheet('${sheetId}')" title="シートを複製">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="8" width="10" height="14" rx="2" ry="2"></rect><path d="M5 18V6a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v2"></path></svg>
                        </button>
                        <button onclick="app.deleteSheet('${sheetId}')" class="btn-delete" title="シートを削除">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                        </button>
                    </div>
                `;
                container.appendChild(header);

                // Rows
                const rowsContainer = document.createElement('div');
                rowsContainer.className = 'rows-container';

                sheet.rows.forEach(row => {
                    const rowEl = document.createElement('div');
                    rowEl.className = `row-item row-type-${row.type}`;

                    const otherSheets = this.getActiveSheets().filter(s => s.id !== sheetId);
                    const op = row.operator !== undefined ? row.operator : 1;

                    let inputHtml = '';
                    if (row.type === 'value') {
                        const valueInputClass = op === -1
                            ? 'row-input row-value-input text-right is-minus'
                            : 'row-input row-value-input text-right';
                        const displayValue = this.formatValueDisplayByOperator(row.value, op);
                        const valueInputGroupClass = op === -1 ? 'value-input-group is-minus' : 'value-input-group';
                        inputHtml = `
                            <div class="${valueInputGroupClass}">
                                <input type="text" inputmode="decimal" class="${valueInputClass}" value="${displayValue}" 
                                    oninput="app.handleValueInput(this, ${op}, '${sheetId}', '${row.id}')"
                                    onblur="app.updateValueRow('${sheetId}', '${row.id}', this.value)"
                                    onchange="app.updateValueRow('${sheetId}', '${row.id}', this.value)">
                                <span class="value-unit-inline">${this.state.unit}</span>
                            </div>
                        `;
                    } else {
                        // Ref Selector
                        const options = otherSheets.map(s =>
                            `<option value="${s.id}" ${row.refSheetId === s.id ? 'selected' : ''}>${s.name}</option>`
                        ).join('');
                        let refAmountHtml = `<div class="ref-amount is-empty">未選択 <span>${this.state.unit}</span></div>`;

                        if (row.refSheetId) {
                            const refResult = this.calculateSheetTotal(row.refSheetId);
                            if (refResult.error) {
                                refAmountHtml = `<div class="ref-amount is-error">Err (${refResult.error})</div>`;
                            } else {
                                const refAmount = refResult.value * op;
                                const refAmountClass = refAmount < 0 ? 'is-minus' : 'is-plus';
                                refAmountHtml = `
                                    <div class="ref-amount ${refAmountClass}">
                                        ${Utils.formatNumber(refAmount)} <span>${this.state.unit}</span>
                                    </div>
                                `;
                            }
                        }

                        inputHtml = `
                            <div class="ref-input-group">
                                <select class="row-select w-full" onchange="app.updateRow('${sheetId}', '${row.id}', 'refSheetId', this.value)">
                                    <option value="">-- 選択 --</option>
                                    ${options}
                                </select>
                                ${refAmountHtml}
                            </div>
                        `;
                    }

                    const operatorBtnStyle = op === 1
                        ? 'color: var(--success); font-weight:700;'
                        : 'color: var(--error); font-weight:700;';

                    rowEl.innerHTML = `
                        <div class="row-drag-handle" onclick="app.toggleOperator('${sheetId}', '${row.id}')">
                            <span style="${operatorBtnStyle}">${op === 1 ? '＋' : '－'}</span>
                        </div>
                        <input type="text" class="row-input" style="font-weight:500;" placeholder="項目名" value="${row.label}"
                            onchange="app.updateRow('${sheetId}', '${row.id}', 'label', this.value)">
                        
                        <select class="row-select" onchange="app.updateRow('${sheetId}', '${row.id}', 'type', this.value)">
                            <option value="value" ${row.type === 'value' ? 'selected' : ''}>金額</option>
                            <option value="ref" ${row.type === 'ref' ? 'selected' : ''}>参照</option>
                        </select>

                        <div class="row-input-area">
                            ${inputHtml}
                        </div>

                        <div class="row-tools">
                            <div class="row-actions">
                                <button class="action-row-btn" onclick="app.duplicateRow('${sheetId}', '${row.id}')" title="行を複製">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="7" y="11" width="14" height="8" rx="2" ry="2"></rect><path d="M3 15V9a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v2"></path></svg>
                                </button>
                                <button class="delete-row-btn" onclick="app.deleteRow('${sheetId}', '${row.id}')" title="行を削除">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                                </button>
                            </div>
                        </div>
                    `;
                    rowsContainer.appendChild(rowEl);
                });

                // Add Row Button
                const addBtn = document.createElement('button');
                addBtn.className = 'add-row-btn';
                addBtn.innerText = '+ 行を追加';
                addBtn.onclick = () => app.addRow(sheetId);
                rowsContainer.appendChild(addBtn);

                // Footer
                const totalMode = this.getSheetTotalViewMode(sheet);
                const totalRes = this.getSheetTotalDisplayResult(sheetId);
                const splitCount = this.normalizeSplitCount(sheet.splitCount);
                const footer = document.createElement('div');
                footer.className = 'sheet-footer sheet-footer-under-title';
                footer.innerHTML = `
                    <div class="sheet-total-meta">
                        <div class="total-display-controls">
                            <div class="total-mode-toggle" role="group" aria-label="合計表示の切り替え">
                                <button type="button" class="${totalMode === 'total' ? 'active' : ''}" onclick="app.updateSheetTotalViewMode('${sheetId}', 'total')">合計</button>
                                <button type="button" class="${totalMode === 'perPerson' ? 'active' : ''}" onclick="app.updateSheetTotalViewMode('${sheetId}', 'perPerson')">一人当たり</button>
                            </div>
                            <label class="split-count-control ${totalMode === 'perPerson' ? '' : 'is-hidden'}">
                                <span>人数</span>
                                <input type="number" inputmode="numeric" min="1" step="1" value="${splitCount}"
                                    onchange="app.updateSheetSplitCount('${sheetId}', this.value)">
                                <span>人</span>
                            </label>
                        </div>
                        <div class="sheet-total-summary">
                            <div class="total-label">${this.getSheetTotalLabel(totalMode)}</div>
                            <div class="total-amount">
                                ${this.getSheetTotalMarkup(totalRes)}
                            </div>
                        </div>
                    </div>
                `;
                container.appendChild(footer);
                container.appendChild(rowsContainer);
            }
        }

        // Initialize App
        const app = new App();

