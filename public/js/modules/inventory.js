// 庫存相關功能模組
import { SIZES, tableIdToInventoryType, inventoryTypeToTableId, UNIFORM_TYPES } from './config.js';
import { saveToLocalStorage, loadFromLocalStorage, showAlert } from './utils.js';
import { updateAllocationRatios } from './ui.js';
import { demandData } from './students.js';

// 庫存資料結構
export let inventoryData = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

// 新增: 可分配數量與預留數量的手動調整資料
export let manualAdjustments = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

/**
 * 初始化庫存表格
 */
export function initInventoryTables() {
    const tableIds = [
        'shortSleeveShirtTable',
        'shortSleevePantsTable',
        'longSleeveShirtTable',
        'longSleevePantsTable'
    ];

    tableIds.forEach(tableId => {
        const table = document.getElementById(tableId);
        if (!table) return;

        // 清空表格
        const tbody = table.querySelector('tbody');
        tbody.innerHTML = '';

        // 為每個尺寸添加行
        SIZES.forEach(size => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${size}</td>
                <td><input type="number" class="form-control total-inventory" min="0" value="0"></td>
            `;
            tbody.appendChild(row);

            // 為輸入欄位添加事件監聽器
            const totalInput = row.querySelector('.total-inventory');

            // 更新庫存數據
            totalInput.addEventListener('change', () => {
                updateInventoryData(tableId, size, 'total', parseInt(totalInput.value) || 0);
            });
        });

        // 加載保存的數據
        loadInventoryData(tableId);
    });
}

/**
 * 更新庫存資料
 * @param {string} tableId - 表格ID
 * @param {string} size - 尺寸
 * @param {string} field - 欄位 (total)
 * @param {number} value - 數值
 */
export function updateInventoryData(tableId, size, field, value) {
    const type = tableIdToInventoryType(tableId);
    if (!type) return;

    // 確保該類型的庫存物件存在
    if (!inventoryData[type]) {
        inventoryData[type] = {};
    }

    // 確保該尺寸的庫存物件存在
    if (!inventoryData[type][size]) {
        inventoryData[type][size] = { total: 0 };
    }
    
    // 檢查是否更新總數，如果是則清除該尺寸的手動調整值
    if (field === 'total' && manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
        // 清除該尺寸的手動調整值
        delete manualAdjustments[type][size];
        // 保存到本地儲存
        saveToLocalStorage('manualAdjustments', manualAdjustments);
    }

    // 更新指定欄位
    inventoryData[type][size][field] = value;

    // 更新UI
    updateInventoryUI(tableId, size);

    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    
    // 更新分配比率
    updateAllocationRatios();
}

/**
 * 更新庫存UI
 * @param {string} tableId - 表格ID
 * @param {string} size - 尺寸
 */
export function updateInventoryUI(tableId, size) {
    const table = document.getElementById(tableId);
    if (!table) return;
    
    const type = tableIdToInventoryType(tableId);
    if (!type || !inventoryData[type] || !inventoryData[type][size]) return;
    
    // 尋找對應的行
    const rows = table.querySelectorAll('tbody tr');
    let targetRow = null;
    
    for (const row of rows) {
        const sizeCell = row.querySelector('td:first-child');
        if (sizeCell && sizeCell.textContent === size) {
            targetRow = row;
            break;
        }
    }
    
    if (!targetRow) return;
    
    const data = inventoryData[type][size];
    
    // 更新UI顯示
    const totalInput = targetRow.querySelector('.total-inventory');
    
    if (totalInput) totalInput.value = data.total;
}

/**
 * 加載庫存資料
 * @param {string} tableId - 表格ID
 */
export function loadInventoryData(tableId) {
    const savedData = loadFromLocalStorage('inventoryData', null);
    if (!savedData) return;

    const type = tableIdToInventoryType(tableId);
    if (!type || !savedData[type]) return;

    // 更新內存中的庫存資料
    inventoryData = savedData;

    // 更新UI
    SIZES.forEach(size => {
        if (savedData[type][size]) {
            updateInventoryUI(tableId, size);
        }
    });
}

/**
 * 計算預留數量 - 新方法
 * @param {Object} inventoryData - 庫存資料
 * @param {Object} demandData - 需求資料
 * @param {Object} manualAdjustments - 手動調整資料
 */
export function calculateReservedQuantities(inventoryData, demandData, manualAdjustments) {
    // 處理每種制服類型
    for (const type in inventoryData) {
        if (!inventoryData.hasOwnProperty(type)) continue;
        
        // 計算總庫存
        let totalInventory = 0;
        for (const size in inventoryData[type]) {
            if (inventoryData[type].hasOwnProperty(size)) {
                totalInventory += inventoryData[type][size].total || 0;
            }
        }
        
        // 獲取需求量
        const demand = demandData[type]?.totalDemand || 0;
        
        // 計算分配比例
        let ratio = 0;
        if (totalInventory > 0 && demand > 0) {
            ratio = demand / totalInventory;
        }
        
        // 計算每個尺寸
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const totalInventory = inventoryData[type][size].total || 0;
            let allocatable = 0;
            
            // 檢查是否有手動調整
            if (manualAdjustments && 
                manualAdjustments[type] && 
                manualAdjustments[type][size] !== undefined) {
                // 使用手動調整的值
                allocatable = manualAdjustments[type][size];
            } else {
                // 使用比例計算（改為四捨五入）
                allocatable = Math.round(totalInventory * ratio);
            }
            
            // 確保可分配數量不超過總庫存
            allocatable = Math.min(allocatable, totalInventory);
            
            // 更新庫存資料
            inventoryData[type][size].allocatable = allocatable;
            inventoryData[type][size].reserved = totalInventory - allocatable;
        }
    }
    
    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    
    return inventoryData;
}

/**
 * 計算庫存類型的總數
 * @param {Object} typeInventory - 某類型的庫存資料
 * @returns {number} 總庫存數量
 */
export function calculateTotalInventory(typeInventory) {
    let total = 0;
    for (const size in typeInventory) {
        if (typeInventory.hasOwnProperty(size)) {
            total += typeInventory[size].total || 0;
        }
    }
    return total;
}

/**
 * 更新調整頁面的尺寸表格
 * @param {string} type - 制服類型
 */
export function updateSizeTable(type) {
    const tableId = `${type}AdjustTable`;
    const table = document.getElementById(tableId);
    if (!table) return;

    const tbody = table.querySelector('tbody');
    tbody.innerHTML = '';

    // 獲取分配比例
    const ratioElem = document.getElementById(`${type}Ratio`);
    let ratio = 0;
    if (ratioElem) {
        const percentText = ratioElem.textContent || '0%';
        ratio = parseFloat(percentText) / 100 || 0;
    }

    // 初始化總數計數器
    let totalInventory = 0;
    let totalCalculatedAllocatable = 0;
    let totalManualAdjustment = 0;
    let totalAdjustedAllocatable = 0;
    let totalReserved = 0;

    // 為每個尺寸添加行
    for (const size of SIZES) {
        if (!inventoryData[type][size]) continue;
        
        const data = inventoryData[type][size];
        const total = data.total || 0;
        
        // 使用比例計算可分配數（改為四捨五入）
        const calculatedAllocatable = Math.round(total * ratio);
        
        // 檢查是否有手動調整，如果沒有則使用計算的可分配數
        let manualAdjustment = calculatedAllocatable;
        let adjustedAllocatable = calculatedAllocatable;
        
        if (manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
            manualAdjustment = manualAdjustments[type][size];
            adjustedAllocatable = manualAdjustment;
        }
        
        const reserved = total - adjustedAllocatable;
        
        // 更新總數
        totalInventory += total;
        totalCalculatedAllocatable += calculatedAllocatable;
        totalManualAdjustment += manualAdjustment;
        totalAdjustedAllocatable += adjustedAllocatable;
        totalReserved += reserved;

        // 創建行
        const row = document.createElement('tr');
        row.dataset.size = size;
        row.innerHTML = `
            <td>${size}</td>
            <td>${total}</td>
            <td>${calculatedAllocatable}</td>
            <td class="adjustment-cell">
                <div class="input-group adjustment-group">
                    <div class="input-group-prepend">
                        <button type="button" class="btn btn-outline-primary decrement-btn">−</button>
                    </div>
                    <input type="number" class="form-control manual-adjustment" 
                        value="${manualAdjustment}" 
                        min="0" max="${total}" step="1" 
                        data-original="${calculatedAllocatable}">
                    <div class="input-group-append">
                        <button type="button" class="btn btn-outline-primary increment-btn">+</button>
                    </div>
                </div>
            </td>
            <td class="adjusted-allocatable">${adjustedAllocatable}</td>
            <td class="reserved">${reserved}</td>
        `;
        
        tbody.appendChild(row);

        // 添加事件監聽器
        const input = row.querySelector('.manual-adjustment');
        const incrementBtn = row.querySelector('.increment-btn');
        const decrementBtn = row.querySelector('.decrement-btn');

        // 數字輸入框事件監聽器
        input.addEventListener('input', function() {
            const value = parseInt(this.value) || 0;
            const max = parseInt(this.getAttribute('max')) || 0;
            
            // 確保值不超過最大值
            if (value > max) {
                this.value = max;
                return;
            }
            
            // 更新手動調整數據
            if (!manualAdjustments[type]) {
                manualAdjustments[type] = {};
            }
            manualAdjustments[type][size] = value;
            
            // 更新顯示
            const row = this.closest('tr');
            row.querySelector('.adjusted-allocatable').textContent = value;
            row.querySelector('.reserved').textContent = total - value;
            
            // 直接更新總計行（立即反映變更）
            manualUpdateTotalRow(type, table);
            
            // 保存手動調整到本地儲存
            saveToLocalStorage('manualAdjustments', manualAdjustments);
        });

        // 增加按鈕事件監聽器
        incrementBtn.addEventListener('click', function() {
            const currentValue = parseInt(input.value) || 0;
            const max = parseInt(input.getAttribute('max')) || 0;
            
            if (currentValue < max) {
                input.value = currentValue + 1;
                // 觸發 input 事件以更新數據
                input.dispatchEvent(new Event('input'));
            }
        });

        // 減少按鈕事件監聽器
        decrementBtn.addEventListener('click', function() {
            const currentValue = parseInt(input.value) || 0;
            
            if (currentValue > 0) {
                input.value = currentValue - 1;
                // 觸發 input 事件以更新數據
                input.dispatchEvent(new Event('input'));
            }
        });
    }

    // 添加總數列
    const totalRow = document.createElement('tr');
    totalRow.classList.add('table-info', 'total-row');
    totalRow.innerHTML = `
        <td><strong>總計</strong></td>
        <td><strong>${totalInventory}</strong></td>
        <td><strong>${totalCalculatedAllocatable}</strong></td>
        <td><strong>${totalManualAdjustment}</strong></td>
        <td><strong>${totalAdjustedAllocatable}</strong></td>
        <td><strong>${totalReserved}</strong></td>
    `;
    tbody.appendChild(totalRow);
}

/**
 * 手動更新總計行（直接從表格中讀取數據）
 * @param {string} type - 制服類型
 * @param {HTMLElement} table - 表格元素
 */
function manualUpdateTotalRow(type, table) {
    if (!table) {
        table = document.getElementById(`${type}AdjustTable`);
        if (!table) return;
    }
    
    const rows = Array.from(table.querySelectorAll('tbody tr:not(.total-row)'));
    if (rows.length === 0) return;
    
    let totalInventory = 0;
    let totalCalculatedAllocatable = 0;
    let totalManualAdjustment = 0;
    let totalAdjustedAllocatable = 0;
    let totalReserved = 0;
    
    rows.forEach(row => {
        const cells = row.cells;
        totalInventory += parseInt(cells[1].textContent) || 0;
        totalCalculatedAllocatable += parseInt(cells[2].textContent) || 0;
        
        // 獲取手動調整值
        const manualInput = cells[3].querySelector('input');
        if (manualInput) {
            totalManualAdjustment += parseInt(manualInput.value) || 0;
        }
        
        totalAdjustedAllocatable += parseInt(cells[4].textContent) || 0;
        totalReserved += parseInt(cells[5].textContent) || 0;
    });
    
    // 更新總計行
    const totalRow = table.querySelector('.total-row');
    if (totalRow) {
        totalRow.innerHTML = `
            <td><strong>總計</strong></td>
            <td><strong>${totalInventory}</strong></td>
            <td><strong>${totalCalculatedAllocatable}</strong></td>
            <td><strong>${totalManualAdjustment}</strong></td>
            <td><strong>${totalAdjustedAllocatable}</strong></td>
            <td><strong>${totalReserved}</strong></td>
        `;
    }
}

/**
 * 更新總計行 - 用於向後兼容
 * @param {string} type - 制服類型
 */
function updateTotalRow(type) {
    manualUpdateTotalRow(type);
}

/**
 * 保存手動調整
 */
export function saveManualAdjustments() {
    let hasWarning = false;
    
    // 遍歷所有制服類型並更新庫存資料
    for (const type in manualAdjustments) {
        if (!manualAdjustments.hasOwnProperty(type)) continue;
        
        let typeTotal = 0; // 追蹤每種制服類型的總可分配數
        
        // 遍歷所有尺寸，包括沒有手動調整的尺寸
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const total = inventoryData[type][size].total || 0;
            let allocatable;
            
            // 檢查該尺寸是否有手動調整
            if (manualAdjustments[type] && manualAdjustments[type][size] !== undefined) {
                // 使用手動調整的值
                allocatable = Math.min(manualAdjustments[type][size], total);
            } else {
                // 使用原來計算的值（比例計算）
                const ratioElem = document.getElementById(`${type}Ratio`);
                let ratio = 0;
                if (ratioElem) {
                    const percentText = ratioElem.textContent || '0%';
                    ratio = parseFloat(percentText) / 100 || 0;
                }
                allocatable = Math.round(total * ratio);
            }
            
            // 更新庫存資料
            inventoryData[type][size].allocatable = allocatable;
            inventoryData[type][size].reserved = total - allocatable;
            
            // 累加可分配數
            typeTotal += allocatable;
        }
        
        // 更新對應表格的UI
        const tableId = inventoryTypeToTableId(type);
        if (tableId) {
            for (const size in inventoryData[type]) {
                updateInventoryUI(tableId, size);
            }
        }

        // 檢查當前制服類型的可分配數是否小於需求量
        const demand = demandData[type]?.totalDemand || 0;
        if (typeTotal < demand) {
            const uniformTypeName = UNIFORM_TYPES[type];
            const message = `警告：${uniformTypeName}的可分配數總和(${typeTotal})小於需求量(${demand})，可能會導致部分學生無法分配到制服。`;
            console.warn(message);
            hasWarning = true;
        }

        // 更新表格
        updateSizeTable(type);
    }
    
    // 保存到本地儲存
    saveToLocalStorage('inventoryData', inventoryData);
    saveToLocalStorage('manualAdjustments', manualAdjustments);
    
    // 更新分配比率
    updateAllocationRatios();
    
    // 顯示訊息
    if (hasWarning) {
        showAlert('手動調整已保存，但有部分制服類型的可分配數小於需求量，請注意檢查。', 'warning', 5000);
    } else {
        showAlert('手動調整已保存', 'success');
    }
}

/**
 * 重置手動調整
 */
export function resetManualAdjustments() {
    // 清空手動調整資料
    manualAdjustments = {
        shortSleeveShirt: {},
        shortSleevePants: {},
        longSleeveShirt: {},
        longSleevePants: {}
    };
    
    // 保存到本地儲存
    saveToLocalStorage('manualAdjustments', manualAdjustments);
    
    // 重新載入所有調整表格
    updateSizeTable('shortSleeveShirt');
    updateSizeTable('shortSleevePants');
    updateSizeTable('longSleeveShirt');
    updateSizeTable('longSleevePants');
    
    // 更新分配比率
    updateAllocationRatios();
    
    // 顯示成功訊息
    showAlert('手動調整已重置', 'info');
}

/**
 * 載入庫存和手動調整資料
 */
export function loadInventoryAndAdjustments() {
    // 載入庫存資料
    const savedInventory = loadFromLocalStorage('inventoryData', null);
    if (savedInventory) {
        inventoryData = savedInventory;
    }
    
    // 載入手動調整資料
    const savedAdjustments = loadFromLocalStorage('manualAdjustments', null);
    if (savedAdjustments) {
        manualAdjustments = savedAdjustments;
    }
} 