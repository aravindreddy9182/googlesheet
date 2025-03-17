document.addEventListener('DOMContentLoaded', function() {
    // Constants
    let ROWS = 50;
    let COLS = 26; // A-Z
    
    // State variables
    let activeCell = null;
    let cellData = {};
    let cellFormatting = {};
    let cellValidation = {};
    let selectionStart = null;
    let selectionEnd = null;
    let isResizing = false;
    let dragStartCell = null;
    let chartInstance = null;
    
    // Initialize the grid
    initializeGrid();
    setupEventListeners();
    
    // Button click handlers
    document.getElementById('btnAddRow').addEventListener('click', addRow);
    document.getElementById('btnDeleteRow').addEventListener('click', deleteRow);
    document.getElementById('btnAddColumn').addEventListener('click', addColumn);
    document.getElementById('btnDeleteColumn').addEventListener('click', deleteColumn);
    document.getElementById('btnBold').addEventListener('click', toggleBold);
    document.getElementById('btnItalic').addEventListener('click', toggleItalic);
    document.getElementById('btnColor').addEventListener('click', function() {
        document.getElementById('colorPicker').click();
    });
    
    document.getElementById('colorPicker').addEventListener('change', function(e) {
        if (activeCell) {
            const cellId = activeCell.dataset.id;
            if (!cellFormatting[cellId]) cellFormatting[cellId] = {};
            cellFormatting[cellId].color = e.target.value;
            applyFormatting(activeCell);
        }
    });
    
    document.getElementById('fontFamily').addEventListener('change', function(e) {
        if (activeCell) {
            const cellId = activeCell.dataset.id;
            if (!cellFormatting[cellId]) cellFormatting[cellId] = {};
            cellFormatting[cellId].fontFamily = e.target.value;
            applyFormatting(activeCell);
        }
    });
    
    document.getElementById('fontSize').addEventListener('change', function(e) {
        if (activeCell) {
            const cellId = activeCell.dataset.id;
            if (!cellFormatting[cellId]) cellFormatting[cellId] = {};
            cellFormatting[cellId].fontSize = e.target.value + 'px';
            applyFormatting(activeCell);
        }
    });
    
    // Function dialog handlers
    document.getElementById('btnFunction').addEventListener('click', function() {
        showDialog('functionDialog');
    });
    
    document.getElementById('cancelFunction').addEventListener('click', function() {
        hideDialog('functionDialog');
    });
    
    document.getElementById('applyFunction').addEventListener('click', function() {
        const functionType = document.getElementById('functionType').value;
        const range = document.getElementById('functionRange').value;
        
        if (activeCell && range) {
            const formula = `=${functionType}(${range})`;
            setCellValue(activeCell, formula);
            document.getElementById('formulaInput').value = formula;
            calculateFormulas();
        }
        
        hideDialog('functionDialog');
    });
    
    // Find and Replace dialog handlers
    document.getElementById('btnFindReplace').addEventListener('click', function() {
        showDialog('findReplaceDialog');
    });
    
    document.getElementById('cancelFindReplace').addEventListener('click', function() {
        hideDialog('findReplaceDialog');
    });
    
    document.getElementById('applyFindReplace').addEventListener('click', function() {
        const findText = document.getElementById('findText').value;
        const replaceText = document.getElementById('replaceText').value;
        const range = document.getElementById('searchRange').value;
        
        if (findText && range) {
            findAndReplace(findText, replaceText, range);
        }
        
        hideDialog('findReplaceDialog');
    });
    
    // Data Validation dialog handlers
    document.getElementById('btnDataValidation').addEventListener('click', function() {
        showDialog('dataValidationDialog');
    });
    
    document.getElementById('cancelValidation').addEventListener('click', function() {
        hideDialog('dataValidationDialog');
    });
    
    document.getElementById('applyValidation').addEventListener('click', function() {
        const validationType = document.getElementById('validationType').value;
        const range = document.getElementById('validationRange').value;
        
        if (range) {
            applyDataValidation(validationType, range);
        }
        
        hideDialog('dataValidationDialog');
    });
    
    // Remove duplicates button handler
    document.getElementById('btnRemoveDuplicates').addEventListener('click', function() {
        if (selectionStart && selectionEnd) {
            const range = getCellRangeNotation(selectionStart, selectionEnd);
            removeDuplicates(range);
        } else {
            alert('Please select a range first');
        }
    });
    
    // Save and Load handlers
    document.getElementById('btnSave').addEventListener('click', saveSpreadsheet);
    document.getElementById('btnLoad').addEventListener('click', () => document.getElementById('loadInput').click());
    document.getElementById('loadInput').addEventListener('change', loadSpreadsheet);
    
    // Chart handlers
    document.getElementById('btnCreateChart').addEventListener('click', () => showDialog('chartDialog'));
    document.getElementById('cancelChart').addEventListener('click', () => hideDialog('chartDialog'));
    document.getElementById('applyChart').addEventListener('click', createChart);
    document.getElementById('closeChart').addEventListener('click', () => {
        document.getElementById('chartContainer').classList.add('hidden');
    });
    
    // Formula input handler
    document.getElementById('formulaInput').addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && activeCell) {
            setCellValue(activeCell, e.target.value);
            calculateFormulas();
        }
    });
    
    function initializeGrid() {
        const thead = document.querySelector('#grid thead tr');
        const tbody = document.querySelector('#grid tbody');
        
        // Create column headers (A, B, C, ...)
        for (let i = 0; i < COLS; i++) {
            const th = document.createElement('th');
            th.className = 'column-header';
            th.textContent = String.fromCharCode(65 + i);
            
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'column-resize-handle';
            resizeHandle.dataset.col = i;
            th.appendChild(resizeHandle);
            
            thead.appendChild(th);
        }
        
        // Create rows and cells
        for (let i = 0; i < ROWS; i++) {
            const tr = document.createElement('tr');
            
            const th = document.createElement('th');
            th.className = 'row-header';
            th.textContent = i + 1;
            
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'row-resize-handle';
            resizeHandle.dataset.row = i;
            th.appendChild(resizeHandle);
            
            tr.appendChild(th);
            
            for (let j = 0; j < COLS; j++) {
                const td = document.createElement('td');
                td.className = 'cell';
                td.contentEditable = true;
                
                const colLetter = String.fromCharCode(65 + j);
                const rowNumber = i + 1;
                const cellId = `${colLetter}${rowNumber}`;
                td.dataset.id = cellId;
                td.dataset.row = i;
                td.dataset.col = j;
                
                tr.appendChild(td);
            }
            
            tbody.appendChild(tr);
        }
    }
    
    function setupEventListeners() {
        const cells = document.querySelectorAll('.cell');
        cells.forEach(cell => {
            cell.addEventListener('click', function(e) {
                selectCell(this);
            });
            
            cell.addEventListener('input', function(e) {
                const value = this.textContent;
                setCellValue(this, value);
            });
            
            cell.addEventListener('keydown', function(e) {
                handleCellKeydown(e, this);
            });
            
            cell.addEventListener('mousedown', function(e) {
                if (e.button === 0) {
                    selectionStart = this;
                    selectionEnd = this;
                    
                    if (e.offsetX > this.offsetWidth - 8 && e.offsetY > this.offsetHeight - 8) {
                        dragStartCell = this;
                    }
                }
            });
            
            cell.addEventListener('mousemove', function(e) {
                if (e.buttons === 1 && selectionStart) {
                    clearSelectedCells();
                    selectionEnd = this;
                    highlightSelectedRange(selectionStart, selectionEnd);
                }
            });
        });
        
        const colResizeHandles = document.querySelectorAll('.column-resize-handle');
        colResizeHandles.forEach(handle => {
            handle.addEventListener('mousedown', function(e) {
                isResizing = true;
                const colIndex = parseInt(this.dataset.col);
                const cells = document.querySelectorAll(`.cell[data-col="${colIndex}"]`);
                const header = document.querySelectorAll('th.column-header')[colIndex];
                
                const initialWidth = cells[0].offsetWidth;
                const startX = e.clientX;
                
                const handleMouseMove = function(e) {
                    if (isResizing) {
                        const width = initialWidth + (e.clientX - startX);
                        if (width >= 40) {
                            cells.forEach(cell => cell.style.width = `${width}px`);
                            header.style.width = `${width}px`;
                        }
                    }
                };
                
                const handleMouseUp = function() {
                    isResizing = false;
                    document.removeEventListener('mousemove', handleMouseMove);
                    document.removeEventListener('mouseup', handleMouseUp);
                };
                
                document.addEventListener('mousemove', handleMouseMove);
                document.addEventListener('mouseup', handleMouseUp);
                e.preventDefault();
                e.stopPropagation();
            });
        });
        
        const rowResizeHandles = document.querySelectorAll('.row-resize-handle');
        rowResizeHandles.forEach(handle => {
            handle.addEventListener('mousedown', function(e) {
                isResizing = true;
                const rowIndex = parseInt(this.dataset.row);
                const row = document.querySelector(`#grid tbody tr:nth-child(${rowIndex + 1})`);
                
                const initialHeight = row.offsetHeight;
                const startY = e.clientY;
                
                const handleMouseMove = function(e) {
                    if (isResizing) {
                        const height = initialHeight + (e.clientY - startY);
                        if (height >= 20) {
                            row.style.height = `${height}px`;
                        }
                    }
                };
                
                const handleMouseUp = function() {
                    isResizing = false;
                    document.removeEventListener('mousemove', handleMouseMove);
                    document.removeEventListener('mouseup', handleMouseUp);
                };
                
                document.addEventListener('mousemove', handleMouseMove);
                document.addEventListener('mouseup', handleMouseUp);
                e.preventDefault();
                e.stopPropagation();
            });
        });
        
        document.addEventListener('mouseup', function(e) {
            if (dragStartCell && selectionEnd && dragStartCell !== selectionEnd) {
                fillCells(dragStartCell, selectionEnd);
            }
            dragStartCell = null;
        });
    }
    
    function selectCell(cell) {
        clearSelectedCells();
        activeCell = cell;
        cell.classList.add('selected');
        document.getElementById('cellAddress').textContent = cell.dataset.id;
        
        const cellId = cell.dataset.id;
        let formula = cellId in cellData ? (cellData[cellId].formula || cellData[cellId].value || '') : '';
        document.getElementById('formulaInput').value = formula;
        
        updateFormattingControls(cell);
    }
    
    function clearSelectedCells() {
        document.querySelectorAll('.cell.selected').forEach(cell => cell.classList.remove('selected'));
        document.querySelectorAll('.cell.in-range').forEach(cell => cell.classList.remove('in-range'));
    }
    
    function highlightSelectedRange(start, end) {
        const startRow = parseInt(start.dataset.row);
        const startCol = parseInt(start.dataset.col);
        const endRow = parseInt(end.dataset.row);
        const endCol = parseInt(end.dataset.col);
        
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        
        for (let i = minRow; i <= maxRow; i++) {
            for (let j = minCol; j <= maxCol; j++) {
                const cell = document.querySelector(`.cell[data-row="${i}"][data-col="${j}"]`);
                if (cell) {
                    if (cell === start) {
                        cell.classList.add('selected');
                    } else {
                        cell.classList.add('in-range');
                    }
                }
            }
        }
    }
    
    function handleCellKeydown(event, cell) {
        const row = parseInt(cell.dataset.row);
        const col = parseInt(cell.dataset.col);
        
        switch (event.key) {
            case 'Tab':
                event.preventDefault();
                const nextCell = document.querySelector(`.cell[data-row="${row}"][data-col="${col + 1}"]`);
                if (nextCell) selectCell(nextCell);
                break;
            case 'Enter':
                event.preventDefault();
                const cellBelow = document.querySelector(`.cell[data-row="${row + 1}"][data-col="${col}"]`);
                if (cellBelow) selectCell(cellBelow);
                break;
            case 'ArrowUp':
                event.preventDefault();
                const cellAbove = document.querySelector(`.cell[data-row="${row - 1}"][data-col="${col}"]`);
                if (cellAbove) selectCell(cellAbove);
                break;
            case 'ArrowDown':
                event.preventDefault();
                const cellDown = document.querySelector(`.cell[data-row="${row + 1}"][data-col="${col}"]`);
                if (cellDown) selectCell(cellDown);
                break;
            case 'ArrowLeft':
                if (cell.textContent.length === 0 || window.getSelection().anchorOffset === 0) {
                    event.preventDefault();
                    const cellLeft = document.querySelector(`.cell[data-row="${row}"][data-col="${col - 1}"]`);
                    if (cellLeft) selectCell(cellLeft);
                }
                break;
            case 'ArrowRight':
                if (cell.textContent.length === 0 || window.getSelection().anchorOffset === cell.textContent.length) {
                    event.preventDefault();
                    const cellRight = document.querySelector(`.cell[data-row="${row}"][data-col="${col + 1}"]`);
                    if (cellRight) selectCell(cellRight);
                }
                break;
        }
    }
    
    function setCellValue(cell, value) {
        const cellId = cell.dataset.id;
        if (!cellData[cellId]) cellData[cellId] = {};
        
        if (value && typeof value === 'string' && value.startsWith('=')) {
            cellData[cellId].formula = value;
        } else {
            cellData[cellId].formula = null;
            cellData[cellId].value = value;
            
            if (cellValidation[cellId]) {
                const validationResult = validateCellData(value, cellValidation[cellId]);
                if (!validationResult.valid) {
                    alert(validationResult.message);
                    cell.style.backgroundColor = '#ffcccc';
                    return;
                } else {
                    cell.style.backgroundColor = '';
                }
            }
            cell.textContent = value;
        }
    }
    
    function calculateFormulas() {
        for (const cellId in cellData) {
            const data = cellData[cellId];
            if (data.formula) {
                try {
                    const result = calculateFormula(data.formula);
                    data.value = result;
                    const cell = document.querySelector(`.cell[data-id="${cellId}"]`);
                    if (cell) cell.textContent = result;
                } catch (error) {
                    console.error(`Error in ${cellId}:`, error);
                    const cell = document.querySelector(`.cell[data-id="${cellId}"]`);
                    if (cell) cell.textContent = '#ERROR!';
                }
            }
        }
    }
    
    function calculateFormula(formula) {
        const expression = formula.substring(1);
        if (expression.includes('(') && expression.includes(')')) {
            const functionMatch = expression.match(/([A-Z]+)\((.*)\)/);
            if (functionMatch) {
                const functionName = functionMatch[1];
                const params = functionMatch[2];
                
                if (params.includes(':')) {
                    const range = getCellsInRange(params);
                    switch (functionName) {
                        case 'SUM': return calculateSum(range);
                        case 'AVERAGE': return calculateAverage(range);
                        case 'MAX': return calculateMax(range);
                        case 'MIN': return calculateMin(range);
                        case 'COUNT': return calculateCount(range);
                        case 'TRIM': return applyTrim(params);
                        case 'UPPER': return applyUpper(params);
                        case 'LOWER': return applyLower(params);
                        case 'POWER':
                            const [base, exp] = params.split(',').map(p => getCellValue(p.trim()));
                            return Math.pow(base, exp);
                        case 'ROUND':
                            const [num, dec] = params.split(',').map(p => getCellValue(p.trim()));
                            return Number(num.toFixed(dec || 0));
                        case 'IF':
                            const [cond, valTrue, valFalse] = params.split(',').map(p => p.trim());
                            const condResult = eval(cond.replace(/(\$?[A-Z]+\$?[0-9]+)/g, match => getCellValue(match.replace(/\$/g, ''))));
                            return condResult ? getCellValue(valTrue) : getCellValue(valFalse);
                        case 'ISNUMBER':
                            return !isNaN(getCellValue(params)) && typeof getCellValue(params) !== 'string';
                        case 'ISTEXT':
                            return typeof getCellValue(params) === 'string' && isNaN(getCellValue(params));
                        default:
                            throw new Error(`Unknown function: ${functionName}`);
                    }
                } else {
                    const cellValue = getCellValue(params);
                    switch (functionName) {
                        case 'TRIM': return applyTrim(cellValue);
                        case 'UPPER': return applyUpper(cellValue);
                        case 'LOWER': return applyLower(cellValue);
                        case 'ISNUMBER': return !isNaN(cellValue) && typeof cellValue !== 'string';
                        case 'ISTEXT': return typeof cellValue === 'string' && isNaN(cellValue);
                        default: throw new Error(`Unknown function: ${functionName}`);
                    }
                }
            }
        }
        
        const cellRefRegex = /(\$?[A-Z]+)(\$?[0-9]+)/g;
        let calcExpression = expression;
        const refs = expression.match(cellRefRegex) || [];
        for (const ref of refs) {
            const value = getCellValue(ref.replace(/\$/g, ''));
            calcExpression = calcExpression.replace(ref, value);
        }
        return eval(calcExpression);
    }
    
    function getCellsInRange(rangeNotation) {
        const [start, end] = rangeNotation.split(':');
        const startCol = start.match(/[A-Z]+/)[0].charCodeAt(0) - 65;
        const startRow = parseInt(start.match(/\d+/)[0]) - 1;
        const endCol = end.match(/[A-Z]+/)[0].charCodeAt(0) - 65;
        const endRow = parseInt(end.match(/\d+/)[0]) - 1;
        
        const result = [];
        for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
                const cellId = `${String.fromCharCode(65 + col)}${row + 1}`;
                result.push({
                    id: cellId,
                    value: cellData[cellId] ? cellData[cellId].value : null
                });
            }
        }
        return result;
    }
    
    function getCellValue(cellId) {
        if (!isNaN(cellId)) return parseFloat(cellId);
        if (cellData[cellId] && cellData[cellId].value !== undefined) {
            const value = cellData[cellId].value;
            return !isNaN(value) ? parseFloat(value) : value;
        }
        return 0;
    }
    
    function calculateSum(range) {
        return range.reduce((sum, cell) => {
            const value = parseFloat(cell.value);
            return sum + (isNaN(value) ? 0 : value);
        }, 0);
    }
    
    function calculateAverage(range) {
        const sum = calculateSum(range);
        const count = range.filter(cell => cell.value !== null && !isNaN(cell.value)).length;
        return count > 0 ? sum / count : 0;
    }
    
    function calculateMax(range) {
        return Math.max(...range.map(cell => parseFloat(cell.value) || Number.NEGATIVE_INFINITY));
    }
    
    function calculateMin(range) {
        return Math.min(...range.map(cell => parseFloat(cell.value) || Number.POSITIVE_INFINITY));
    }
    
    function calculateCount(range) {
        return range.filter(cell => cell.value !== null && !isNaN(cell.value)).length;
    }
    
    function applyTrim(text) {
        return typeof text === 'string' ? text.trim() : text;
    }
    
    function applyUpper(text) {
        return typeof text === 'string' ? text.toUpperCase() : text;
    }
    
    function applyLower(text) {
        return typeof text === 'string' ? text.toLowerCase() : text;
    }
    
    function findAndReplace(findText, replaceText, rangeNotation) {
        const cells = getCellsInRange(rangeNotation);
        cells.forEach(cell => {
            const cellElement = document.querySelector(`.cell[data-id="${cell.id}"]`);
            if (cellElement && cellData[cell.id] && typeof cellData[cell.id].value === 'string') {
                const newValue = cellData[cell.id].value.replace(new RegExp(findText, 'g'), replaceText);
                setCellValue(cellElement, newValue);
            }
        });
    }
    
    function removeDuplicates(rangeNotation) {
        const [start, end] = rangeNotation.split(':');
        const startCol = start.charCodeAt(0) - 65;
        const startRow = parseInt(start.substring(1)) - 1;
        const endCol = end.charCodeAt(0) - 65;
        const endRow = parseInt(end.substring(1)) - 1;
        
        const rows = [];
        for (let row = startRow; row <= endRow; row++) {
            const rowData = [];
            for (let col = startCol; col <= endCol; col++) {
                const cellId = `${String.fromCharCode(65 + col)}${row + 1}`;
                rowData.push({
                    id: cellId,
                    value: cellData[cellId] ? cellData[cellId].value : null
                });
            }
            rows.push(rowData);
        }
        
        const uniqueRows = [];
        const seen = new Set();
        rows.forEach(row => {
            const key = row.map(cell => cell.value).join(',');
            if (!seen.has(key)) {
                seen.add(key);
                uniqueRows.push(row);
            }
        });
        
        for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
                const cellId = `${String.fromCharCode(65 + col)}${row + 1}`;
                const cellElement = document.querySelector(`.cell[data-id="${cellId}"]`);
                if (cellElement) setCellValue(cellElement, '');
            }
        }
        
        uniqueRows.forEach((row, i) => {
            const rowIndex = startRow + i;
            row.forEach((cell, j) => {
                const colIndex = startCol + j;
                const cellId = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                const cellElement = document.querySelector(`.cell[data-id="${cellId}"]`);
                if (cellElement) setCellValue(cellElement, cell.value);
            });
        });
        
        alert(`Removed ${rows.length - uniqueRows.length} duplicate rows`);
    }
    
    function applyDataValidation(validationType, rangeNotation) {
        const cells = getCellsInRange(rangeNotation);
        cells.forEach(cell => {
            cellValidation[cell.id] = validationType;
            if (cellData[cell.id] && cellData[cell.id].value) {
                const validationResult = validateCellData(cellData[cell.id].value, validationType);
                const cellElement = document.querySelector(`.cell[data-id="${cell.id}"]`);
                if (!validationResult.valid && cellElement) {
                    cellElement.style.backgroundColor = '#ffcccc';
                }
            }
        });
        alert(`Data validation applied to ${cells.length} cells`);
    }
    
    function validateCellData(value, validationType) {
        if (value === null || value === undefined || value === '') return { valid: true };
        
        switch (validationType) {
            case 'number':
                if (isNaN(value)) return { valid: false, message: 'This cell only accepts numbers' };
                break;
            case 'text':
                if (!isNaN(value) && typeof value !== 'string') return { valid: false, message: 'This cell only accepts text' };
                break;
            case 'date':
                const dateRegex = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
                if (!dateRegex.test(value)) return { valid: false, message: 'This cell only accepts dates (MM/DD/YYYY)' };
                break;
        }
        return { valid: true };
    }
    
    function toggleBold() {
        if (activeCell) {
            const cellId = activeCell.dataset.id;
            if (!cellFormatting[cellId]) cellFormatting[cellId] = {};
            cellFormatting[cellId].bold = !cellFormatting[cellId].bold;
            applyFormatting(activeCell);
        }
    }
    
    function toggleItalic() {
        if (activeCell) {
            const cellId = activeCell.dataset.id;
            if (!cellFormatting[cellId]) cellFormatting[cellId] = {};
            cellFormatting[cellId].italic = !cellFormatting[cellId].italic;
            applyFormatting(activeCell);
        }
    }
    
    function applyFormatting(cell) {
        const cellId = cell.dataset.id;
        if (cellFormatting[cellId]) {
            const fmt = cellFormatting[cellId];
            cell.style.fontWeight = fmt.bold ? 'bold' : 'normal';
            cell.style.fontStyle = fmt.italic ? 'italic' : 'normal';
            if (fmt.color) cell.style.color = fmt.color;
            if (fmt.fontFamily) cell.style.fontFamily = fmt.fontFamily;
            if (fmt.fontSize) cell.style.fontSize = fmt.fontSize;
        }
    }
    
    function updateFormattingControls(cell) {
        const cellId = cell.dataset.id;
        const formatting = cellFormatting[cellId] || {};
        
        document.getElementById('btnBold').style.backgroundColor = formatting.bold ? '#e0e0e0' : 'transparent';
        document.getElementById('btnItalic').style.backgroundColor = formatting.italic ? '#e0e0e0' : 'transparent';
        document.getElementById('fontFamily').value = formatting.fontFamily || 'Arial';
        document.getElementById('fontSize').value = formatting.fontSize ? formatting.fontSize.replace('px', '') : '12';
    }
    
    function addRow() {
        const tbody = document.querySelector('#grid tbody');
        const newRowIndex = tbody.childElementCount;
        const tr = document.createElement('tr');
        
        const th = document.createElement('th');
        th.className = 'row-header';
        th.textContent = newRowIndex + 1;
        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'row-resize-handle';
        resizeHandle.dataset.row = newRowIndex;
        th.appendChild(resizeHandle);
        tr.appendChild(th);
        
        for (let j = 0; j < COLS; j++) {
            const td = document.createElement('td');
            td.className = 'cell';
            td.contentEditable = true;
            const colLetter = String.fromCharCode(65 + j);
            const cellId = `${colLetter}${newRowIndex + 1}`;
            td.dataset.id = cellId;
            td.dataset.row = newRowIndex;
            td.dataset.col = j;
            
            td.addEventListener('click', function(e) { selectCell(this); });
            td.addEventListener('input', function(e) { setCellValue(this, this.textContent); });
            td.addEventListener('keydown', function(e) { handleCellKeydown(e, this); });
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
        ROWS++;
    }
    
    function deleteRow() {
        if (activeCell && document.querySelector('#grid tbody').childElementCount > 1) {
            const rowIndex = parseInt(activeCell.dataset.row);
            const tbody = document.querySelector('#grid tbody');
            tbody.deleteRow(rowIndex);
            
            for (let i = rowIndex; i < tbody.childElementCount; i++) {
                const row = tbody.children[i];
                row.firstChild.textContent = i + 1;
                row.firstChild.querySelector('.row-resize-handle').dataset.row = i;
                const cells = row.querySelectorAll('.cell');
                cells.forEach((cell, j) => {
                    const colLetter = String.fromCharCode(65 + j);
                    const cellId = `${colLetter}${i + 1}`;
                    cell.dataset.id = cellId;
                    cell.dataset.row = i;
                });
            }
            const newRow = rowIndex < tbody.childElementCount ? rowIndex : rowIndex - 1;
            if (newRow >= 0) selectCell(document.querySelector(`.cell[data-row="${newRow}"][data-col="0"]`));
            ROWS--;
        }
    }
    
    function addColumn() {
        const thead = document.querySelector('#grid thead tr');
        const tbody = document.querySelector('#grid tbody');
        const newColIndex = thead.childElementCount - 1;
        
        const th = document.createElement('th');
        th.className = 'column-header';
        th.textContent = String.fromCharCode(65 + newColIndex);
        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'column-resize-handle';
        resizeHandle.dataset.col = newColIndex;
        th.appendChild(resizeHandle);
        thead.appendChild(th);
        
        for (let i = 0; i < tbody.childElementCount; i++) {
            const row = tbody.children[i];
            const td = document.createElement('td');
            td.className = 'cell';
            td.contentEditable = true;
            const colLetter = String.fromCharCode(65 + newColIndex);
            const cellId = `${colLetter}${i + 1}`;
            td.dataset.id = cellId;
            td.dataset.row = i;
            td.dataset.col = newColIndex;
            
            td.addEventListener('click', function(e) { selectCell(this); });
            td.addEventListener('input', function(e) { setCellValue(this, this.textContent); });
            td.addEventListener('keydown', function(e) { handleCellKeydown(e, this); });
            row.appendChild(td);
        }
        COLS++;
    }
    
    function deleteColumn() {
        if (activeCell && document.querySelector('#grid thead tr').childElementCount > 2) {
            const colIndex = parseInt(activeCell.dataset.col);
            const thead = document.querySelector('#grid thead tr');
            const tbody = document.querySelector('#grid tbody');
            
            thead.deleteCell(colIndex + 1);
            for (let i = 0; i < tbody.childElementCount; i++) {
                tbody.children[i].deleteCell(colIndex + 1);
            }
            
            for (let j = colIndex; j < thead.childElementCount - 1; j++) {
                const header = thead.children[j + 1];
                header.textContent = String.fromCharCode(65 + j);
                header.querySelector('.column-resize-handle').dataset.col = j;
                for (let i = 0; i < tbody.childElementCount; i++) {
                    const cell = tbody.children[i].children[j + 1];
                    const cellId = `${String.fromCharCode(65 + j)}${i + 1}`;
                    cell.dataset.id = cellId;
                    cell.dataset.col = j;
                }
            }
            const newCol = colIndex < thead.childElementCount - 2 ? colIndex : colIndex - 1;
            if (newCol >= 0) {
                const rowIndex = parseInt(activeCell.dataset.row);
                selectCell(document.querySelector(`.cell[data-row="${rowIndex}"][data-col="${newCol}"]`));
            }
            COLS--;
        }
    }
    
    function fillCells(startCell, endCell) {
        const startRow = parseInt(startCell.dataset.row);
        const startCol = parseInt(startCell.dataset.col);
        const endRow = parseInt(endCell.dataset.row);
        const endCol = parseInt(endCell.dataset.col);
        
        const sourceValue = cellData[startCell.dataset.id]?.value || '';
        const sourceFormula = cellData[startCell.dataset.id]?.formula || null;
        
        for (let i = Math.min(startRow, endRow); i <= Math.max(startRow, endRow); i++) {
            for (let j = Math.min(startCol, endCol); j <= Math.max(startCol, endCol); j++) {
                if (i === startRow && j === startCol) continue;
                const cell = document.querySelector(`.cell[data-row="${i}"][data-col="${j}"]`);
                if (cell) {
                    if (sourceFormula) {
                        const adjustedFormula = adjustFormula(sourceFormula, startRow, startCol, i, j);
                        setCellValue(cell, adjustedFormula);
                    } else {
                        setCellValue(cell, sourceValue);
                    }
                }
            }
        }
        calculateFormulas();
    }
    
    function adjustFormula(formula, fromRow, fromCol, toRow, toCol) {
        const cellRefRegex = /(\$?[A-Z]+)(\$?[0-9]+)/g;
        return formula.replace(cellRefRegex, (match, colRef, rowRef) => {
            const isColAbsolute = colRef.startsWith('$');
            const isRowAbsolute = rowRef.startsWith('$');
            let col = colRef.replace('$', '').charCodeAt(0) - 65;
            let row = parseInt(rowRef.replace('$', '')) - 1;
            
            if (!isColAbsolute) col += toCol - fromCol;
            if (!isRowAbsolute) row += toRow - fromRow;
            return `${String.fromCharCode(65 + col)}${row + 1}`;
        });
    }
    
    function getCellRangeNotation(startCell, endCell) {
        return `${startCell.dataset.id}:${endCell.dataset.id}`;
    }
    
    function saveSpreadsheet() {
        const data = {
            cellData,
            cellFormatting,
            cellValidation,
            dimensions: {
                rows: document.querySelector('#grid tbody').childElementCount,
                cols: document.querySelector('#grid thead tr').childElementCount - 1
            }
        };
        const json = JSON.stringify(data);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `spreadsheet-${new Date().toISOString()}.json`;
        a.click();
        URL.revokeObjectURL(url);
    }
    
    function loadSpreadsheet(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = function(event) {
            const data = JSON.parse(event.target.result);
            const tbody = document.querySelector('#grid tbody');
            const thead = document.querySelector('#grid thead tr');
            while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
            while (thead.childElementCount > 1) thead.removeChild(thead.lastChild);
            
            ROWS = data.dimensions.rows;
            COLS = data.dimensions.cols;
            
            initializeGrid();
            cellData = data.cellData;
            cellFormatting = data.cellFormatting;
            cellValidation = data.cellValidation;
            
            for (const cellId in cellData) {
                const cell = document.querySelector(`.cell[data-id="${cellId}"]`);
                if (cell) {
                    setCellValue(cell, cellData[cellId].value);
                    applyFormatting(cell);
                }
            }
            calculateFormulas();
            setupEventListeners();
        };
        reader.readAsText(file);
    }
    
    function createChart() {
        const range = document.getElementById('chartRange').value;
        const chartType = document.getElementById('chartType').value;
        const cells = getCellsInRange(range);
        
        const labels = [];
        const data = [];
        let currentRow = -1;
        
        cells.forEach(cell => {
            const row = parseInt(cell.id.match(/\d+$/)[0]) - 1;
            if (row !== currentRow) {
                currentRow = row;
                if (row === 0) {
                    labels.push(cell.value || '');
                } else {
                    data.push([]);
                }
            }
            if (row > 0) {
                data[row - 1].push(parseFloat(cell.value) || 0);
            }
        });
        
        const ctx = document.getElementById('chartCanvas').getContext('2d');
        if (chartInstance) chartInstance.destroy();
        
        chartInstance = new Chart(ctx, {
            type: chartType,
            data: {
                labels: labels,
                datasets: data.map((rowData, i) => ({
                    label: `Series ${i + 1}`,
                    data: rowData,
                    backgroundColor: `rgba(${Math.random()*255}, ${Math.random()*255}, ${Math.random()*255}, 0.5)`
                }))
            },
            options: {
                scales: {
                    y: { beginAtZero: true }
                }
            }
        });
        
        const chartContainer = document.getElementById('chartContainer');
        chartContainer.style.left = '50%';
        chartContainer.style.top = '50%';
        chartContainer.style.transform = 'translate(-50%, -50%)';
        chartContainer.classList.remove('hidden');
        
        hideDialog('chartDialog');
    }
    
    function showDialog(dialogId) {
        document.getElementById(dialogId).classList.remove('hidden');
        document.getElementById('overlay').classList.remove('hidden');
    }
    
    function hideDialog(dialogId) {
        document.getElementById(dialogId).classList.add('hidden');
        document.getElementById('overlay').classList.add('hidden');
    }
});
