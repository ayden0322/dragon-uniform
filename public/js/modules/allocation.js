// 分配相關功能模組
import { saveToLocalStorage, loadFromLocalStorage, showAlert, downloadExcel } from './utils.js';
import { SIZES, UNIFORM_TYPES, formatSize, currentSizeDisplayMode, SIZE_DISPLAY_MODES, getFemaleChestAdjustment, getCurrentSchoolConfig } from './config.js';
import { inventoryData, calculateTotalInventory, updateInventoryUI, manualAdjustments, initInventoryFeatures, saveManualAdjustments, saveManualAdjustmentsSilent, calculateReservedQuantities } from './inventory.js';
import { studentData, sortedStudentData, demandData, updateStudentAllocationUI, updateAdjustmentPage } from './students.js';
import { updateAllocationRatios, formatSizeWithAdjustment } from './ui.js';

// 記錄最後分配狀態
const lastAllocationStatus = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

// 新的統一褲子腰圍映射表
const waistToPantsSizeMap_V2 = [
    { min: 20, max: 21, baseSize: "XS/34", adjustments: [ // 腰圍20-21 | XS/34(男&女 當褲長 >=34 則尺碼+1 號)
        { genders: ['男', '女'], pantsLengthThreshold: 34, sizeAdjustment: +1, mark: '↑' }
    ]},
    { min: 22, max: 24, baseSize: "S/36", adjustments: [ // 腰圍22-24 | S/36(男 :當褲長 >=37 則尺碼+1 號)(女 :當褲長 >=38 則尺碼+1 號)
        { genders: ['男'], pantsLengthThreshold: 37, sizeAdjustment: +1, mark: '↑' },
        { genders: ['女'], pantsLengthThreshold: 38, sizeAdjustment: +1, mark: '↑' }
    ]},
    { min: 25, max: 27, baseSize: "M/38", adjustments: [ // 腰圍25-27 | M/38(男 :當褲長 >=38 則尺碼+1 號)(女 :當褲長 >=39 則尺碼+1 號)
        { genders: ['男'], pantsLengthThreshold: 38, sizeAdjustment: +1, mark: '↑' },
        { genders: ['女'], pantsLengthThreshold: 39, sizeAdjustment: +1, mark: '↑' }
    ]},
    { min: 28, max: 30, baseSize: "L/40", adjustments: [ // 腰圍28-30 | L/40(男 :當褲長 >=40 則尺碼+1 號) (女 :當褲長 >=41 則尺碼+1 號)
        { genders: ['男'], pantsLengthThreshold: 40, sizeAdjustment: +1, mark: '↑' },
        { genders: ['女'], pantsLengthThreshold: 41, sizeAdjustment: +1, mark: '↑' } 
    ]},
    { min: 31, max: 33, baseSize: "XL/42", adjustments: [ // 腰圍31-33 | XL/42(男 :當褲長 >=42 則尺碼+1 號) 
        { genders: ['男'], pantsLengthThreshold: 42, sizeAdjustment: +1, mark: '↑' }
        // 女性此區間依照使用者規則，無特定褲長調整
    ]},
    { min: 34, max: 36, baseSize: "2L/44", adjustments: [ // 腰圍34-36 | 2L/44(男&女 當褲長 >=42 則尺碼+1 號)
        { genders: ['男', '女'], pantsLengthThreshold: 42, sizeAdjustment: +1, mark: '↑' }
    ]},
    { min: 37, max: 38, baseSize: "3L/46", adjustments: [ // 腰圍37-38 | 3L/46(男&女 當褲長 >=42 則尺碼+1 號)
        { genders: ['男', '女'], pantsLengthThreshold: 42, sizeAdjustment: +1, mark: '↑' }
    ]},
    { min: 39, max: 40, baseSize: "4L/48", adjustments: [ // 腰圍39-40 | 4L/48(男&女 當褲長 >=42 則尺碼+1 號)
        { genders: ['男', '女'], pantsLengthThreshold: 42, sizeAdjustment: +1, mark: '↑' }
    ]},
    // 腰圍41-43 | 5L/50（庫存不足時可用 6L/52 替代）
    { min: 41, max: 43, baseSize: "5L/50", alternativeSize: "6L/52", adjustments: [] }, 
    // 腰圍44-46 | 7L/54（庫存不足時可用 8L/56 替代）
    { min: 44, max: 46, baseSize: "7L/54", alternativeSize: "8L/56", adjustments: [] }  
];

// 本地排序學生數據，用於避免修改導入的常量
let _localSortedStudentData = [];

// 分配統計資料
export let allocationStats = {
    shortSleeveShirt: {
        allocated: 0,
        exact: 0,
        different: 0,
        failed: 0,
        special: 0,
        pantsSizeAdjusted: 0
    },
    shortSleevePants: {
        allocated: 0,
        exact: 0,
        different: 0,
        failed: 0,
        special: 0,
        pantsSizeAdjusted: 0
    },
    longSleeveShirt: {
        allocated: 0,
        exact: 0,
        different: 0,
        failed: 0,
        special: 0,
        pantsSizeAdjusted: 0
    },
    longSleevePants: {
        allocated: 0,
        exact: 0,
        different: 0,
        failed: 0,
        special: 0,
        pantsSizeAdjusted: 0
    }
};

// 分配進行中狀態標誌
let allocationInProgress = false;

/**
 * 設置分配進行中狀態
 * @param {boolean} status - 是否正在分配中
 */
function setAllocationInProgress(status) {
    allocationInProgress = status;
    
    // 如果需要，可以根據狀態更新UI元素
    const allocateButton = document.getElementById('allocateButton');
    if (allocateButton) {
        allocateButton.disabled = status;
        allocateButton.textContent = status ? '分配中...' : '開始分配';
    }
}

/**
 * 排序學生資料，用於制服分配
 * @returns {Array} 排序後的學生資料
 */
function sortStudents() {
    // 深拷貝學生資料
    const sortedData = [...studentData];
    
    // 按有效胸圍（胸圍和腰圍的較大值）排序
    sortedData.sort((a, b) => {
        // 計算有效胸圍
        const aEffectiveChest = calculateEffectiveChest(a);
        const bEffectiveChest = calculateEffectiveChest(b);
        
        // 按有效胸圍降序排序
        if (aEffectiveChest !== bEffectiveChest) {
            return bEffectiveChest - aEffectiveChest;
        }
        
        // 相同胸圍的情況下，按班級和座號排序
        const classComparison = String(a.class || '').localeCompare(String(b.class || ''));
        if (classComparison !== 0) {
            return classComparison;
        }
        
        // 班級相同的情況下，按座號排序
        const aNumber = parseInt(a.number) || 9999;
        const bNumber = parseInt(b.number) || 9999;
        return aNumber - bNumber;
    });
    
    return sortedData;
}

/**
 * 開始制服分配
 */
export async function startAllocation() {
    console.log('開始制服分配過程');
    
    // 新增：在重置分配前儲存當前所有調整資料
    try {
        console.log('保存當前的所有調整資料（無通知模式）');
        // 調用無通知版本的保存手動調整函數
        saveManualAdjustmentsSilent();
        // 確保庫存資料也被保存
        saveToLocalStorage('inventoryData', inventoryData);
        console.log('所有調整資料已保存');
    } catch (error) {
        console.warn('保存調整資料時發生錯誤:', error);
        // 不中斷流程，繼續進行分配
    }
    
    try {
        // 獲取排序後的學生數據
        _localSortedStudentData = sortStudents();
        
        // 重置分配結果
        resetAllocation();
        
        // 檢查是否有學生資料
        if (!_localSortedStudentData || _localSortedStudentData.length === 0) {
            console.warn('沒有學生資料可供分配');
            throw new Error('沒有學生資料可供分配');
        }
        
        // 檢查是否有庫存資料
        if (!inventoryData) {
            console.warn('沒有庫存資料可供分配');
            throw new Error('沒有庫存資料可供分配');
        }
        
        // 在開始分配前，重新計算所有制服類型的預留數量
        console.log('在開始分配前重新計算所有制服類型的預留數量');
        calculateReservedQuantities(inventoryData, demandData, manualAdjustments);
        console.log('所有制服類型的預留數量計算完成');
        
        console.log(`開始為 ${_localSortedStudentData.length} 名學生分配制服`);
        
        // 顯示所有制服類型的庫存和需求情況
        console.log('------------------------------');
        console.log('分配前庫存與需求對比:');
        
        const uniformTypes = [
            { type: 'shortSleeveShirt', name: '短衣', field: 'shortSleeveShirtCount' },
            { type: 'shortSleevePants', name: '短褲', field: 'shortSleevePantsCount' },
            { type: 'longSleeveShirt', name: '長衣', field: 'longSleeveShirtCount' },
            { type: 'longSleevePants', name: '長褲', field: 'longSleevePantsCount' }
        ];
        
        for (const { type, name, field } of uniformTypes) {
            // 檢查是否有此類型的庫存
            if (!inventoryData[type]) {
                console.warn(`沒有 ${name} 庫存數據`);
                continue;
            }
            
            // 計算總需求量
            let totalDemand = 0;
            _localSortedStudentData.forEach(student => {
                totalDemand += student[field] || 1;
            });
            
            // 計算總可分配數量
            let totalAllocatable = 0;
            let totalInventory = 0;
            let totalReserved = 0;
            
            for (const size in inventoryData[type]) {
                const total = inventoryData[type][size].total || 0;
                const allocatable = inventoryData[type][size].allocatable || 0;
                const reserved = inventoryData[type][size].reserved || 0;
                
                totalInventory += total;
                totalAllocatable += allocatable;
                totalReserved += reserved;
            }
            
            console.log(`${name}: 總庫存=${totalInventory}, 預留=${totalReserved}, 可分配=${totalAllocatable}, 需求=${totalDemand}`);
            
            // 如果總可分配數小於總需求量，發出警告
            if (totalAllocatable < totalDemand) {
                console.warn(`警告：${name}可分配數總和(${totalAllocatable})小於需求量(${totalDemand})，差額${totalDemand - totalAllocatable}件`);
            }
        }
        console.log('------------------------------');
        
        // 設置分配中狀態
        setAllocationInProgress(true);
        
        let allocationSuccess = true;
        
        // 依序分配各種制服，每個函數單獨處理錯誤
        try {
            await allocateShortSleeveShirts();
            console.log('短衣分配完成');
        } catch (error) {
            console.error('短衣分配過程發生錯誤:', error);
            allocationSuccess = false;
        }
        
        try {
            await allocateShortSleevePants();
            console.log('短褲分配完成');
        } catch (error) {
            console.error('短褲分配過程發生錯誤:', error);
            allocationSuccess = false;
        }
        
        try {
            await allocateLongSleeveShirts();
            console.log('長衣分配完成');
        } catch (error) {
            console.error('長衣分配過程發生錯誤:', error);
            allocationSuccess = false;
        }
        
        try {
            await allocateLongSleevePants();
            console.log('長褲分配完成');
        } catch (error) {
            console.error('長褲分配過程發生錯誤:', error);
            allocationSuccess = false;
        }
        
        // 保存分配結果
        saveData();
        
        // 更新分配結果頁面
        updateAllocationResults();
        
        if (allocationSuccess) {
            console.log('制服分配完成');
        } else {
            console.warn('制服分配部分完成，有些類型的分配過程發生錯誤');
        }
        
        // 檢查是否應重新分配
        const shouldReallocate = checkPantsLengthDeficiency();
        
        // 重置分配中狀態
        setAllocationInProgress(false);
        
        return shouldReallocate;
    } catch (error) {
        console.error('分配過程發生錯誤:', error);
        showAlert('分配過程發生錯誤: ' + error.message, 'error');
        
        // 重置分配中狀態
        setAllocationInProgress(false);
        return false;
    }
}

/**
 * 重置制服分配結果
 */
export function resetAllocation() {
    console.log('重置分配結果');
    // 重設學生分配結果
    if (studentData && studentData.length) {
        studentData.forEach(student => {
            // 清除分配的尺寸結果
            student.allocatedShirtSize = '';
            student.allocatedPantsSize = '';
            student.allocatedLongShirtSize = '';
            student.allocatedLongPantsSize = '';
            
            // 清除特殊分配標記
            student.isSpecialShirtAllocation = false;
            student.isSpecialPantsAllocation = false;
            student.isSpecialLongShirtAllocation = false;
            student.isSpecialLongPantsAllocation = false;
            
            // 清除褲長調整標記
            student.isPantsLengthAdjusted = false;
            
            // 清除分配失敗原因
            student.allocationFailReason = {};
            
            // 移除計算的臨時尺寸
            delete student._calculatedShirtSize;
            delete student._calculatedPantsSize;
        });
    }
    
    // 重置庫存分配結果
    for (const type in inventoryData) {
        if (!inventoryData.hasOwnProperty(type)) continue;
        
        let typeTotalAllocatable = 0;
        
        console.log(`重置 ${UNIFORM_TYPES[type]} 庫存分配結果:`);
        
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const total = inventoryData[type][size].total || 0;
            const reserved = inventoryData[type][size].reserved || 0;
            
            // 記錄重置前的狀態
            console.log(`  重置前 ${size}: 總數=${total}, 預留=${reserved}, 已分配=${inventoryData[type][size].allocated || 0}, 可分配=${inventoryData[type][size].allocatable || 0}`);
            
            // 重置已分配數量為0
            inventoryData[type][size].allocated = 0;
            // 重置可分配數量為總數減去預留數
            inventoryData[type][size].allocatable = total - reserved;
            
            typeTotalAllocatable += inventoryData[type][size].allocatable;
            
            console.log(`  重置後 ${size}: 總數=${total}, 預留=${reserved}, 可分配=${inventoryData[type][size].allocatable}`);
        }
        
        console.log(`${UNIFORM_TYPES[type]} 重置完成，總可分配數量=${typeTotalAllocatable}`);
    }
    
    // 檢查預留比例是否正確
    console.log('檢查各類型制服的預留比例:');
    for (const type in inventoryData) {
        let totalInventory = 0;
        let totalAllocatable = 0;
        
        for (const size in inventoryData[type]) {
            totalInventory += inventoryData[type][size].total || 0;
            totalAllocatable += inventoryData[type][size].allocatable || 0;
        }
        
        const reservedRatio = totalInventory > 0 ? (totalInventory - totalAllocatable) / totalInventory : 0;
        console.log(`${UNIFORM_TYPES[type]}: 總庫存=${totalInventory}, 總可分配=${totalAllocatable}, 實際預留比例=${(reservedRatio * 100).toFixed(1)}%`);
    }
    
    // 重置最後分配狀態記錄
    for (const type in lastAllocationStatus) {
        lastAllocationStatus[type] = {};
    }
    console.log('已重置最後分配狀態記錄');
    
    // 重置統計數據
    resetAllocationStats();
}

/**
 * 分配短袖上衣
 */
export async function allocateShortSleeveShirts() {
    return allocateShortShirts('shortSleeveShirt', 'allocatedShirtSize', 'isSpecialShirtAllocation');
}

/**
 * 分配短袖褲子
 */
export async function allocateShortSleevePants() {
    console.log('開始分配短褲 (全新邏輯)');
    
    // 在分配前重新計算預留數量，確保使用最新的計算邏輯
    console.log('重新計算短褲預留數量，以確保使用最新的計算邏輯');
    
    // 創建一個臨時對象，只包含短褲的庫存數據
    const tempInventoryData = {
        shortSleevePants: JSON.parse(JSON.stringify(inventoryData.shortSleevePants))
    };
    
    // 重新計算預留數量
    const updatedInventory = calculateReservedQuantities(tempInventoryData, demandData, manualAdjustments);
    
    // 更新短褲庫存數據
    inventoryData.shortSleevePants = updatedInventory.shortSleevePants;
    
    // 輸出更新後的預留數量信息
    console.log('短褲預留數量重新計算完成');
    
    return allocatePantsNewLogic(
        _localSortedStudentData, 
        'shortSleevePants', 
        'allocatedPantsSize', 
        'pantsAdjustmentMark', 
        'shortSleevePantsCount',
        inventoryData.shortSleevePants
    );
}

/**
 * 分配長袖上衣
 */
export async function allocateLongSleeveShirts() {
    // 確保使用正確的庫存類型
    console.log('開始分配長衣');
    return allocateLongShirts('longSleeveShirt', 'allocatedLongShirtSize', 'isSpecialLongShirtAllocation');
}

/**
 * 分配長袖褲子
 */
export async function allocateLongSleevePants() {
    console.log('開始分配長褲 (全新邏輯)');
    
    // 在分配前重新計算預留數量，確保使用最新的計算邏輯
    console.log('重新計算長褲預留數量，以確保使用最新的計算邏輯');
    
    // 創建一個臨時對象，只包含長褲的庫存數據
    const tempInventoryData = {
        longSleevePants: JSON.parse(JSON.stringify(inventoryData.longSleevePants))
    };
    
    // 重新計算預留數量
    const updatedInventory = calculateReservedQuantities(tempInventoryData, demandData, manualAdjustments);
    
    // 更新長褲庫存數據
    inventoryData.longSleevePants = updatedInventory.longSleevePants;
    
    // 輸出更新後的預留數量信息
    console.log('長褲預留數量重新計算完成');
    
    return allocatePantsNewLogic(
        _localSortedStudentData, 
        'longSleevePants', 
        'allocatedLongPantsSize', 
        'longPantsAdjustmentMark',
        'longSleevePantsCount',
        inventoryData.longSleevePants
    );
}

/**
 * 計算有效胸圍（考慮性別和腰圍）
 * @param {Object} student - 學生資料
 * @returns {number} - 計算後的有效胸圍
 */
function calculateEffectiveChest(student) {
    let effectiveChest = student.chest || 0;
    // 女生胸圍調整已關閉，使用原始胸圍值
    // 如果腰圍大於胸圍，使用腰圍
    const waist = student.waist || 0;
    if (waist > effectiveChest) {
        effectiveChest = waist;
    }
    return effectiveChest;
}

/**
 * 檢查尺寸與上衣尺寸的階級差異
 * @param {string} pantsSize - 褲子尺寸
 * @param {string} shirtSize - 上衣尺寸
 * @returns {boolean} - 是否符合階級差異要求
 */
function isAcceptableSizeDifference(pantsSize, shirtSize) {
    // 如果沒有上衣尺寸，則返回true（允許分配）
    if (!shirtSize) {
        console.log(`isAcceptableSizeDifference: 上衣尺寸為空，自動允許分配褲子尺寸 ${pantsSize}`);
        return true;
    }
    
    const shirtIndex = SIZES.indexOf(shirtSize);
    const pantsIndex = SIZES.indexOf(pantsSize);
    
    // 計算差異並檢查是否可接受
    const diff = Math.abs(shirtIndex - pantsIndex);
    const isAcceptable = diff <= 1;
    
    console.log(`isAcceptableSizeDifference: 檢查褲子尺寸 ${pantsSize}(索引=${pantsIndex}) 與上衣尺寸 ${shirtSize}(索引=${shirtIndex}) 的差異為 ${diff}，結果: ${isAcceptable ? '可接受' : '不可接受'}`);
    
    return isAcceptable;
}

/**
 * 檢查長褲尺寸與長袖上衣尺寸的階級差異
 * @param {string} pantsSize - 長褲尺寸
 * @param {string} shirtSize - 長袖上衣尺寸
 * @returns {boolean} - 是否符合階級差異要求
 */
function isAcceptableLongPantsSizeDifference(pantsSize, shirtSize) {
    // 如果沒有上衣尺寸，則返回true（允許分配）
    if (!shirtSize) {
        console.log(`isAcceptableLongPantsSizeDifference: 長袖上衣尺寸為空，自動允許分配長褲尺寸 ${pantsSize}`);
        return true;
    }
    
    const shirtIndex = SIZES.indexOf(shirtSize);
    const pantsIndex = SIZES.indexOf(pantsSize);
    
    // 計算差異並檢查是否可接受
    const diff = Math.abs(shirtIndex - pantsIndex);
    const isAcceptable = diff <= 2;
    
    console.log(`isAcceptableLongPantsSizeDifference: 檢查長褲尺寸 ${pantsSize}(索引=${pantsIndex}) 與長袖上衣尺寸 ${shirtSize}(索引=${shirtIndex}) 的差異為 ${diff}，結果: ${isAcceptable ? '可接受' : '不可接受'}`);
    
    return isAcceptable;
}

/**
 * 減少庫存數量
 * @param {Object} inventory - 庫存資料
 * @param {string} size - 尺寸
 * @param {number} count - 要減少的數量
 * @param {string} [inventoryType] - 庫存類型，用於記錄最後狀態
 * @returns {boolean} - 是否成功減少庫存
 */
function decreaseInventory(inventory, size, count, inventoryType) {
    if (inventory[size]) {
        // 確保使用正確的可分配數量
        const allocatable = inventory[size].allocatable || 0;
        
        // 檢查庫存是否足夠 - 避免減少超過可用數量
        if (allocatable < count) {
            console.error(`無法減少庫存: 尺寸 ${size} 需要 ${count} 件，但只剩 ${allocatable} 件可分配`);
            return false;
        }
        
        const actualCount = Math.min(count, allocatable);
        
        // 記錄原始值以便偵錯
        console.log(`減少庫存前 ${size}: 總數=${inventory[size].total}, 可分配=${allocatable}, 已分配=${inventory[size].allocated}, 預留=${inventory[size].reserved}`);
        
        // 減少可分配數量並增加已分配數量
        inventory[size].allocatable -= actualCount;
        inventory[size].allocated += actualCount;
        
        // 輸出實際減少的庫存量與更新後的狀態
        console.log(`減少庫存 ${size}: ${actualCount} 件，剩餘可分配=${inventory[size].allocatable} 件，總分配=${inventory[size].allocated} 件，總預留=${inventory[size].reserved} 件`);
        
        // 確保不會出現負數
        inventory[size].allocatable = Math.max(0, inventory[size].allocatable);
        
        // 記錄最後分配狀態
        if (inventoryType && lastAllocationStatus[inventoryType]) {
            lastAllocationStatus[inventoryType][size] = {
                allocated: inventory[size].allocated,
                remaining: inventory[size].allocatable,
                timestamp: new Date().getTime()
            };
            console.log(`記錄 ${inventoryType} 尺寸 ${size} 的最後分配狀態: 已分配=${inventory[size].allocated}, 剩餘=${inventory[size].allocatable}`);
        }
        
        return true;
    }
    
    console.error(`無法減少庫存: 尺寸 ${size} 不存在於庫存中`);
    return false;
}

/**
 * 分配短袖上衣
 */
function allocateShortShirts(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`%c===== 開始分配 ${UNIFORM_TYPES[inventoryType]} =====`, 'background: #3498db; color: white; font-size: 14px; padding: 5px;');
        
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            _localSortedStudentData.forEach(student => {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '無庫存資料';
                student[allocatedField] = '';
                student.shirtAllocationMark = ''; // Clear mark
            });
            allocationStats[inventoryType].failed = _localSortedStudentData.length;
            resolve(false);
            return;
        }
        
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.shortSleeveShirtCount || 1;
        });
        
        let totalAllocatable = 0;
        for (const size in inventoryData[inventoryType]) {
            totalAllocatable += inventoryData[inventoryType][size]?.allocatable || 0;
        }
        
        console.log(`%c${UNIFORM_TYPES[inventoryType]}需求與庫存概況：`, 'color: #2980b9; font-weight: bold;');
        console.log(`- 總需求量：${totalDemand}件`);
        console.log(`- 總可分配數量：${totalAllocatable}件`);
        console.log(`- 學生數量：${_localSortedStudentData.length}人`);
        
        if (totalAllocatable < totalDemand) {
            console.warn(`%c警告：${UNIFORM_TYPES[inventoryType]}可分配數總和(${totalAllocatable})小於需求量(${totalDemand})，差額${totalDemand - totalAllocatable}件`, 'color: #e74c3c; font-weight: bold;');
        }
        
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        const sortedStudents = [..._localSortedStudentData].map(student => {
            const chest = student.chest || 0;
            const waist = student.waist || 0;
            const effectiveChest = Math.max(chest, waist);
            return { student, effectiveChest };
        }).sort((a, b) => a.effectiveChest - b.effectiveChest); // Sorting by effectiveChest ASC

        const stats = {
            allocated: 0, exact: 0, different: 0, failed: 0, 
            special: 0, // For general special allocations, or could be femaleDowngraded specifically
            pantsSizeAdjusted: 0, 
            preferredFailedSwitchToAlternative: 0, 
            femaleDowngraded: 0 // Specific stat for this new logic
        };
        
        for (const {student, effectiveChest} of sortedStudents) {
            if (student[allocatedField] && student.shirtAllocationMark !== undefined) { // Skip if already allocated (and mark processed)
                console.log(`學生 ${student.name}(${student.class}-${student.number}) 已有短衣分配結果：${student[allocatedField]}${student.shirtAllocationMark || ''}，跳過`);
                continue;
            }
            
            const requiredCount = student.shortSleeveShirtCount || 1;
            student.shirtAllocationMark = ''; // Initialize/reset mark for current student

            console.log(`%c處理短衣分配給學生：${student.name}(${student.gender}，班級：${student.class}，座號：${student.number})`, 'color: #2c3e50; font-weight: bold;');
            console.log(`- 胸圍：${student.chest}，腰圍：${student.waist}，褲長：${student.pantsLength}，有效胸圍：${effectiveChest}，需求：${requiredCount}`);

            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生不需要短衣，跳過`);
                stats.failed++;
                continue;
            }
            
            let preferredAdjustment, alternativeAdjustment;
            if (student.gender === '男') {
                preferredAdjustment = (effectiveChest % 2 === 0) ? 10 : 11;
                alternativeAdjustment = (effectiveChest % 2 === 0) ? 12 : 13;
            } else {
                preferredAdjustment = (effectiveChest % 2 === 0) ? 8 : 9;
                alternativeAdjustment = (effectiveChest % 2 === 0) ? 10 : 11;
            }

            const preferredTargetSizeNumber = effectiveChest + preferredAdjustment;
            const alternativeTargetSizeNumber = effectiveChest + alternativeAdjustment;

            function getTargetSizeFromNumber(sizeNumber) {
                if (sizeNumber <= 34) return 'XS/34';
                if (sizeNumber <= 36) return 'S/36';
                if (sizeNumber <= 38) return 'M/38';
                if (sizeNumber <= 40) return 'L/40';
                if (sizeNumber <= 42) return 'XL/42';
                if (sizeNumber <= 44) return '2L/44';
                if (sizeNumber <= 46) return '3L/46';
                if (sizeNumber <= 48) return '4L/48';
                if (sizeNumber <= 50) return '5L/50';
                if (sizeNumber <= 52) return '6L/52';
                if (sizeNumber <= 54) return '7L/54';
                if (sizeNumber <= 56) return '8L/56';
                return '9L/58';
            }

            const preferredTargetSize = getTargetSizeFromNumber(preferredTargetSizeNumber);
            const alternativeTargetSize = getTargetSizeFromNumber(alternativeTargetSizeNumber);

            console.log(`- 理論首選短衣尺碼: ${preferredTargetSize} (基於 ${preferredTargetSizeNumber})`);
            console.log(`- 理論備選短衣尺碼: ${alternativeTargetSize} (基於 ${alternativeTargetSizeNumber})`);

            let finalAllocatedSize = null;
            let allocationMark = '';
            let attemptLog = [];
            let isSpecialAllocation = false; // General flag for special cases (like alternative or forced smallest/largest)

            // Attempt 1: Preferred Size
            attemptLog.push(`嘗試首選尺碼 ${preferredTargetSize}`);
            let preferredPathMadePantsAdjustment = false;

            if (remainingInventory[preferredTargetSize]?.allocatable >= requiredCount) {
                let currentAttemptSize = preferredTargetSize;
                attemptLog.push(`首選尺碼 ${currentAttemptSize} 有庫存 (${remainingInventory[preferredTargetSize].allocatable})`);

                if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, currentAttemptSize)) {
                    const originalSizeBeforeAdjustment = currentAttemptSize;
                    const largerSize = getNextLargerSize(originalSizeBeforeAdjustment);
                    attemptLog.push(`褲長 (${student.pantsLength}) vs ${getSizeNumber(originalSizeBeforeAdjustment)}，建議調整到 ${largerSize}`);
                    if (remainingInventory[largerSize]?.allocatable >= requiredCount) {
                        if (largerSize !== originalSizeBeforeAdjustment) {
                            currentAttemptSize = largerSize;
                            stats.pantsSizeAdjusted++;
                            preferredPathMadePantsAdjustment = true;
                            attemptLog.push(`褲長調整成功，使用 ${currentAttemptSize} (庫存 ${remainingInventory[currentAttemptSize].allocatable})`);
                } else {
                            attemptLog.push(`褲長建議調整，但 ${originalSizeBeforeAdjustment} 已是最大可獲取尺碼或無更大尺碼。繼續使用 ${originalSizeBeforeAdjustment}。`);
                        }
                } else {
                        attemptLog.push(`褲長調整失敗：建議的尺碼 ${largerSize} 庫存不足 (${remainingInventory[largerSize]?.allocatable || 0})。`);
                    }
                }
                
                if (remainingInventory[currentAttemptSize]?.allocatable >= requiredCount) {
                    finalAllocatedSize = currentAttemptSize;
                    if (preferredPathMadePantsAdjustment) {
                        allocationMark = '↑';
                    }
                } else {
                    attemptLog.push(`最終檢查：尺碼 ${currentAttemptSize} (可能是原始首選或經調整的尺碼) 庫存不足。首選尺碼路徑失敗。`);
                    finalAllocatedSize = null;
                    preferredPathMadePantsAdjustment = false;
                }
            } else {
                attemptLog.push(`首選尺碼 ${preferredTargetSize} 庫存不足或不存在 (${remainingInventory[preferredTargetSize]?.allocatable || 0})`);
            }

            // Attempt 2: Alternative Size (if preferred failed)
            if (!finalAllocatedSize) {
                attemptLog.push(`首選短衣失敗，嘗試備選尺碼 ${alternativeTargetSize}`);
                stats.preferredFailedSwitchToAlternative++;
                isSpecialAllocation = true; // Using alternative is a "special" case
                let alternativePathMadePantsAdjustment = false;

                if (remainingInventory[alternativeTargetSize]?.allocatable >= requiredCount) {
                    let currentAttemptSize = alternativeTargetSize;
                    allocationMark = '↑'; // Base mark for alternative
                    attemptLog.push(`備選短衣尺碼 ${currentAttemptSize} 有庫存 (${remainingInventory[currentAttemptSize].allocatable})`);

                    if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, currentAttemptSize)) {
                        const originalSizeBeforeAdjustment = currentAttemptSize;
                        const largerSize = getNextLargerSize(originalSizeBeforeAdjustment);
                        attemptLog.push(`褲長 (${student.pantsLength}) vs ${getSizeNumber(originalSizeBeforeAdjustment)} (備選路徑)，建議調整到 ${largerSize}`);
                        if (remainingInventory[largerSize]?.allocatable >= requiredCount) {
                           if (largerSize !== originalSizeBeforeAdjustment) {
                                currentAttemptSize = largerSize;
                                stats.pantsSizeAdjusted++; // Count this adjustment
                                alternativePathMadePantsAdjustment = true;
                                // allocationMark is already '↑', no change needed for this sub-adjustment
                                attemptLog.push(`褲長調整成功 (備選路徑)，使用 ${currentAttemptSize} (庫存 ${remainingInventory[currentAttemptSize].allocatable})`);
                } else {
                                attemptLog.push(`褲長建議調整 (備選路徑)，但 ${originalSizeBeforeAdjustment} 已是最大可獲取尺碼。繼續使用 ${originalSizeBeforeAdjustment}。`);
                }
            } else {
                            attemptLog.push(`褲長調整失敗 (備選路徑)：建議的尺碼 ${largerSize} 庫存不足 (${remainingInventory[largerSize]?.allocatable || 0})。`);
                            // Stick with currentAttemptSize (which is alternativeTargetSize) if its stock is OK
                        }
                    }
                    
                    if (remainingInventory[currentAttemptSize]?.allocatable >= requiredCount) {
                        finalAllocatedSize = currentAttemptSize;
                        // allocationMark is already '↑' from taking alternative path.
                        // If alternativePathMadePantsAdjustment also happened, it's still '↑'.
                } else {
                         attemptLog.push(`最終檢查：備選路徑尺碼 ${currentAttemptSize} (可能經調整) 庫存不足。備選尺碼路徑失敗。`);
                         finalAllocatedSize = null;
                         allocationMark = ''; // Reset mark if alternative path also fails
                         isSpecialAllocation = false;
                    }
                } else {
                    attemptLog.push(`備選短衣尺碼 ${alternativeTargetSize} 庫存不足或不存在 (${remainingInventory[alternativeTargetSize]?.allocatable || 0})`);
                    allocationMark = ''; // Reset mark
                    isSpecialAllocation = false;
                }
            }
            
            // FEMALE SPECIAL ADJUSTMENT LOGIC FOR SHORT SHIRTS
            if (finalAllocatedSize && student.gender === '女' && getSizeNumber(finalAllocatedSize) >= 44 && student.pantsLength <= 38) {
                attemptLog.push(`觸發女生短衣特殊調整：尺碼 ${finalAllocatedSize} (>=44), 褲長 ${student.pantsLength} (<=38)`);
                const smallerSize = getPreviousSmallerSize(finalAllocatedSize);
                attemptLog.push(`嘗試縮小短衣到尺碼 ${smallerSize}`);
                if (remainingInventory[smallerSize]?.allocatable >= requiredCount) {
                    finalAllocatedSize = smallerSize;
                    allocationMark = '↓'; // Override previous mark
                    stats.femaleDowngraded++;
                    isSpecialAllocation = true; // Downgrading is a "special" case
                    attemptLog.push(`女生短衣特殊調整成功，使用 ${finalAllocatedSize} (庫存 ${remainingInventory[smallerSize].allocatable})`);
                } else {
                    attemptLog.push(`女生短衣特殊調整失敗：縮小後尺碼 ${smallerSize} 庫存不足 (${remainingInventory[smallerSize]?.allocatable || 0})。短衣分配失敗。`);
                    finalAllocatedSize = null; 
                    allocationMark = '';
                    isSpecialAllocation = false;
                }
            }

            console.log('短衣分配嘗試日誌:', attemptLog.join(' -> '));

            if (finalAllocatedSize) {
                student[allocatedField] = finalAllocatedSize;
                student.shirtAllocationMark = allocationMark; 
                student[specialField] = isSpecialAllocation; // General special flag

                // Store details for potential Excel export or debug
                student.isShirtSizeAdjustedForPantsLength = preferredPathMadePantsAdjustment || 
                    (allocationMark === '↑' && stats.preferredFailedSwitchToAlternative > 0 && student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, finalAllocatedSize)); // Re-evaluate for alternative if needed
                if (student.isShirtSizeAdjustedForPantsLength && !student.originalShirtSize) { // Avoid overwriting if set by preferred
                    student.originalShirtSize = preferredTargetSize; // Or alternativeTargetSize if that was the base for adjustment
                }


                decreaseInventory(remainingInventory, finalAllocatedSize, requiredCount, inventoryType);
                stats.allocated++;
                if (allocationMark === '↓') { // Female downgrade
                    /* stats.special already incremented by isSpecialAllocation=true */
                } else if (allocationMark === '↑') { // Alternative path or pants adjustment on preferred
                    stats.different++;
                } else if (!isSpecialAllocation) { // Exact preferred, no adjustment
                stats.exact++;
            } else {
                    // Other special cases, like forced smallest/largest if that logic was here
                    stats.special++;
            }
            
                console.log(`%c短衣分配成功：${student.name}(${student.gender}，座號：${student.number}，胸圍：${student.chest}，腰圍：${student.waist}，褲長：${student.pantsLength}) => 尺寸 ${finalAllocatedSize}${allocationMark}，需求 ${requiredCount}件`, 'color: #27ae60; font-weight: bold;');
            if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                delete student.allocationFailReason[inventoryType];
                }
            } else {
                // Allocation failed for this student for this item type
                student.allocationFailReason = student.allocationFailReason || {};
                if (!student.allocationFailReason[inventoryType]) { // Don't overwrite a more specific earlier reason
                     student.allocationFailReason[inventoryType] = `短衣分配失敗(${attemptLog[attemptLog.length-1] || '未知原因'})`;
                }
                stats.failed++;
                console.log(`%c短衣分配失敗：${student.name} (${student.allocationFailReason[inventoryType]})`, 'color: #e74c3c; font-weight: bold;');
            }
        } // End student loop

        // Second stage for unallocated students (simplified)
        const unallocatedStudents = _localSortedStudentData.filter(s => !s[allocatedField] && (s.shortSleeveShirtCount || 1) > 0 && !s.allocationFailReason?.[inventoryType]);
        if (unallocatedStudents.length > 0) {
            console.log(`%c第二階段短衣分配開始：處理 ${unallocatedStudents.length} 名未分配學生`, 'background: #8e44ad; color: white; font-size: 12px; padding: 3px;');
            let allAvailableSizesStage2 = Object.keys(remainingInventory)
                .filter(size => remainingInventory[size]?.allocatable > 0)
                .map(size => ({ size, available: remainingInventory[size].allocatable }))
            .sort((a, b) => SIZES.indexOf(a.size) - SIZES.indexOf(b.size));
        
            for (const {student, effectiveChest} of unallocatedStudents.map(s => ({student: s, effectiveChest: Math.max(s.chest || 0, s.waist || 0)}))) {
                 const requiredCount = student.shortSleeveShirtCount || 1;
                 student.shirtAllocationMark = student.shirtAllocationMark || ''; // Initialize if somehow missed

                if (allAvailableSizesStage2.length === 0) {
                    console.warn(`%c第二階段：所有短衣尺寸庫存都已用完，無法為 ${student.name} 分配`, 'color: #e74c3c;');
                    if(!student.allocationFailReason?.[inventoryType]) student.allocationFailReason[inventoryType] = '短衣分配失敗(第二階段無庫存)';
                    stats.failed++; // Ensure this student is counted as failed if not already
                    continue; 
                }
                
                const bestSizeItem = allAvailableSizesStage2.reduce((prev, curr) => (curr.available > prev.available) ? curr : prev, allAvailableSizesStage2[0]);
                let sizeToAllocateStage2 = bestSizeItem.size;
                let madePantsAdjustmentInStage2 = false;
                let allocationMarkStage2 = ''; // Usually empty for stage 2, unless pants adjustment

                console.log(`%c第二階段短衣特殊分配：${student.name}，嘗試分配庫存最多的尺寸 ${sizeToAllocateStage2} (${bestSizeItem.available}件)`);

                if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, sizeToAllocateStage2)) {
                    const originalSizeForStage2 = sizeToAllocateStage2;
                    const largerSize = getNextLargerSize(originalSizeForStage2);
                    if (remainingInventory[largerSize]?.allocatable >= requiredCount) {
                        if (largerSize !== originalSizeForStage2) {
                            sizeToAllocateStage2 = largerSize;
                            madePantsAdjustmentInStage2 = true;
                            allocationMarkStage2 = '↑';
                            stats.pantsSizeAdjusted++;
                            console.log(`%c第二階段褲長調整成功：從 ${originalSizeForStage2} 到 ${sizeToAllocateStage2}`, 'color: #27ae60;');
                        }
                        } else {
                        console.log(`%c第二階段褲長調整失敗：${largerSize} 庫存不足`, 'color: #e74c3c;');
                    }
                }
                
                if (remainingInventory[sizeToAllocateStage2]?.allocatable >= requiredCount) {
                    student[allocatedField] = sizeToAllocateStage2;
                    student.shirtAllocationMark = allocationMarkStage2;
                    student[specialField] = true; // Stage 2 is always special

                    decreaseInventory(remainingInventory, sizeToAllocateStage2, requiredCount, inventoryType);
                    stats.allocated++;
                    stats.special++; // Count as special
                    if (madePantsAdjustmentInStage2) stats.different++;
                    

                    console.log(`%c第二階段短衣分配成功：${student.name} => ${sizeToAllocateStage2}${allocationMarkStage2}`, 'color: #27ae60; font-weight: bold;');
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    allAvailableSizesStage2 = allAvailableSizesStage2.map(s => s.size === sizeToAllocateStage2 ? { ...s, available: remainingInventory[s.size].allocatable } : s).filter(s => s.available > 0);
                } else {
                    console.log(`%c第二階段短衣分配失敗：${student.name}，尺碼 ${sizeToAllocateStage2} 庫存不足`, 'color: #e74c3c;');
                    if(!student.allocationFailReason?.[inventoryType]) student.allocationFailReason[inventoryType] = '短衣分配失敗(第二階段選擇尺碼庫存不足)';
                    stats.failed++; // Ensure this student is counted as failed if not already
                }
            }
        }


        inventoryData[inventoryType] = remainingInventory;
        allocationStats[inventoryType] = {
            allocated: stats.allocated,
            exact: stats.exact,
            different: stats.different, 
            special: stats.special, // General special + female downgraded
            failed: stats.failed,
            pantsSizeAdjusted: stats.pantsSizeAdjusted,
            preferredFailedSwitchToAlternative: stats.preferredFailedSwitchToAlternative,
            femaleDowngraded: stats.femaleDowngraded
        };

        console.log(`%c${UNIFORM_TYPES[inventoryType]}分配結果統計：`, 'background: #3498db; color: white; font-size: 12px; padding: 5px;');
        console.log(`- 成功分配：${stats.allocated}人`);
        console.log(`  - 精確分配 (首選無調整)：${stats.exact}人`);
        console.log(`  - 不同尺寸分配 (含備選/褲長調整/第二階段調整)：${stats.different}人`);
        console.log(`  - 特殊分配 (含備選/降碼/第二階段)：${stats.special}人`);
        console.log(`  - 女生降碼次數：${stats.femaleDowngraded}人`);
        console.log(`- 分配失敗：${stats.failed}人`);
        console.log(`- 因褲長調整次數：${stats.pantsSizeAdjusted}`);
        console.log(`- 首選失敗轉備選次數：${stats.preferredFailedSwitchToAlternative}`);
        console.log(`%c===== ${UNIFORM_TYPES[inventoryType]}分配完成 =====`, 'background: #3498db; color: white; font-size: 14px; padding: 5px;');
        
        resolve(true);
    });
}

/**
 * 分配短褲
 */
function allocateShortPants(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]} (新規則) ====================`);
        
        _localSortedStudentData.forEach(student => {
            student.pantsAdjustmentMark = null; 
        });
        console.log(`已清除所有學生的短褲調整標記 (新規則)`);
        
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.shortSleevePantsCount || 1; // Ensure this uses the correct count field
        });
        
        let totalAllocatable = 0;
        for (const size in inventoryData[inventoryType]) {
            totalAllocatable += inventoryData[inventoryType][size]?.allocatable || 0;
        }
        
        console.log(`${UNIFORM_TYPES[inventoryType]} (新規則)：總需求量=${totalDemand}，總可分配數量=${totalAllocatable}`);
        
        if (totalAllocatable < totalDemand) {
            console.warn(`${UNIFORM_TYPES[inventoryType]} (新規則) 庫存不足！總需求量=${totalDemand}，總可分配數量=${totalAllocatable}`);
        }
        
        let hasAllocatableSizes = false;
        for (const size in inventoryData[inventoryType]) {
            if (inventoryData[inventoryType][size]?.allocatable > 0) {
                hasAllocatableSizes = true;
                break;
            }
        }
        
        if (!hasAllocatableSizes) {
            console.warn(`${UNIFORM_TYPES[inventoryType]} (新規則) 沒有可分配的尺碼！`);
            resolve(false);
            return;
        }
        
        function findMatchingInventorySize(targetSize) {
            if (inventoryData[inventoryType][targetSize] && 
                inventoryData[inventoryType][targetSize].allocatable > 0) {
                return targetSize;
            }
            console.log(`未找到尺碼 "${targetSize}" (${inventoryType}) 的庫存或庫存為零`);
            return null;
        }

        const newPantsWaistToSizeMap = [
            { 
                min: 20, max: 21, size: "XS/34", 
                adjustCondition: (student) => student.pantsLength >= 34,
                adjustMarkSuffix: "↑"
            },
            { 
                min: 22, max: 24, size: "S/36",
                adjustCondition: (student) => 
                    (student.gender === "男" && student.pantsLength >= 37) ||
                    (student.gender === "女" && student.pantsLength >= 38),
                adjustMarkSuffix: "↑"
            },
            { 
                min: 25, max: 27, size: "M/38",
                adjustCondition: (student) =>
                    (student.gender === "男" && student.pantsLength >= 38) ||
                    (student.gender === "女" && student.pantsLength >= 39),
                adjustMarkSuffix: "↑"
            },
            { 
                min: 28, max: 30, size: "L/40",
                adjustCondition: (student) =>
                    (student.gender === "男" && student.pantsLength >= 40) ||
                    (student.gender === "女" && student.pantsLength >= 41),
                adjustMarkSuffix: "↑"
            },
            { 
                min: 31, max: 33, size: "XL/42",
                adjustCondition: (student) => student.gender === "男" && student.pantsLength >= 42,
                adjustMarkSuffix: "↑"
            },
            { 
                min: 34, max: 36, size: "2L/44",
                adjustCondition: (student) => student.pantsLength >= 42,
                adjustMarkSuffix: "↑"
            },
            { 
                min: 37, max: 38, size: "3L/46",
                adjustCondition: (student) => student.pantsLength >= 42,
                adjustMarkSuffix: "↑"
            },
            { 
                min: 39, max: 40, size: "4L/48",
                adjustCondition: (student) => student.pantsLength >= 42,
                adjustMarkSuffix: "↑"
            },
            { 
                min: 41, max: 43, primarySize: "5L/50", alternateSize: "6L/52"
            },
            { 
                min: 44, max: 46, primarySize: "7L/54", alternateSize: "8L/56"
            }
        ];
        
        let successCount = 0;
        let failedCount = 0;
        
        for (const student of _localSortedStudentData) {
            student.pantsAdjustmentMark = null; // Reset for each student per allocation run

            if (student.shortSleevePantsCount && student.shortSleevePantsCount > 0) { // Corrected count field
                console.log(`\\n分配學生 [${student.id}] ${student.class || ''}-${student.number || ''} ${student.name || ''} (${inventoryType} 新規則): 腰圍=${student.waist}, 褲長=${student.pantsLength}, 性別=${student.gender}`);
                
                const waistRange = newPantsWaistToSizeMap.find(range => 
                    student.waist >= range.min && student.waist <= range.max);
                
                if (!waistRange) {
                    console.warn(`學生 [${student.id}] (${inventoryType} 新規則) 腰圍 ${student.waist} 超出分配範圍`);
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：腰圍超出範圍';
                    failedCount++;
                    continue;
                }
                
                let sizeToAllocate = null;
                let adjustmentDisplayMark = null; // This will store the "M/38↑" style string for UI

                if (waistRange.adjustCondition && waistRange.adjustCondition(student)) {
                    const baseSizeForAdjustment = waistRange.size;
                    const targetAdjustedSize = getNextLargerSize(baseSizeForAdjustment);
                    
                    console.log(`學生 [${student.id}] (${inventoryType} 新規則) 符合褲長調整條件: 原尺碼=${baseSizeForAdjustment}, 嘗試調整至=${targetAdjustedSize}`);

                    if (findMatchingInventorySize(targetAdjustedSize)) {
                        sizeToAllocate = targetAdjustedSize;
                        // Ensure formatSize is used carefully if it's for internal logic vs display
                        // For student.pantsAdjustmentMark, we want the display string like "M/38↑"
                        adjustmentDisplayMark = formatSize(targetAdjustedSize, SIZE_DISPLAY_MODES.size) + (waistRange.adjustMarkSuffix || "↑");
                        console.log(`學生 [${student.id}] (${inventoryType} 新規則) 調整後尺碼 ${targetAdjustedSize} 有庫存. 標記: ${adjustmentDisplayMark}`);
                    } else {
                        console.warn(`學生 [${student.id}] (${inventoryType} 新規則) 調整後尺碼 ${targetAdjustedSize} 無庫存 (原尺碼 ${baseSizeForAdjustment})`);
                        // Option: Try allocating baseSizeForAdjustment if adjusted is OOS? User rules say "則尺碼+1號", implying it's the target.
                        // Current logic: if adjusted OOS, it's a fail for this path.
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = `分配失敗：調整後尺碼(${targetAdjustedSize})無庫存`;
                        failedCount++;
                        continue; 
                    }
                } else if (waistRange.primarySize) {
                    console.log(`學生 [${student.id}] (${inventoryType} 新規則) 檢查主要/替代尺碼: 主要=${waistRange.primarySize}, 替代=${waistRange.alternateSize}`);
                    if (findMatchingInventorySize(waistRange.primarySize)) {
                        sizeToAllocate = waistRange.primarySize;
                        console.log(`學生 [${student.id}] (${inventoryType} 新規則) 使用主要尺碼: ${sizeToAllocate}`);
                    } else {
                        console.log(`學生 [${student.id}] (${inventoryType} 新規則) 主要尺碼 ${waistRange.primarySize} 無庫存, 嘗試替代尺碼 ${waistRange.alternateSize}`);
                        if (findMatchingInventorySize(waistRange.alternateSize)) {
                            sizeToAllocate = waistRange.alternateSize;
                            console.log(`學生 [${student.id}] (${inventoryType} 新規則) 使用替代尺碼: ${sizeToAllocate}`);
                        } else {
                            console.warn(`學生 [${student.id}] (${inventoryType} 新規則) 主要尺碼 ${waistRange.primarySize} 和替代尺碼 ${waistRange.alternateSize} 均無庫存`);
                            student.allocationFailReason = student.allocationFailReason || {};
                            student.allocationFailReason[inventoryType] = '分配失敗：主要及替代尺碼均無庫存';
                            failedCount++;
                            continue;
                        }
                    }
                } else if (waistRange.size) {
                     console.log(`學生 [${student.id}] (${inventoryType} 新規則) 檢查標準尺碼: ${waistRange.size}`);
                    if (findMatchingInventorySize(waistRange.size)) {
                        sizeToAllocate = waistRange.size;
                        console.log(`學生 [${student.id}] (${inventoryType} 新規則) 使用標準尺碼: ${sizeToAllocate}`);
                    } else {
                        console.warn(`學生 [${student.id}] (${inventoryType} 新規則) 標準尺碼 ${waistRange.size} 無庫存`);
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = `分配失敗：標準尺碼(${waistRange.size})無庫存`;
                        failedCount++;
                        continue;
                    }
                }

                if (!sizeToAllocate) {
                    if (!(student.allocationFailReason && student.allocationFailReason[inventoryType])) {
                         console.warn(`學生 [${student.id}] (${inventoryType} 新規則) 無法確定最終分配尺碼`);
                         student.allocationFailReason = student.allocationFailReason || {};
                         student.allocationFailReason[inventoryType] = '分配失敗：無法確定分配尺碼';
                         failedCount++;
                    }
                    continue; 
                }
                
                const requiredCount = student.shortSleevePantsCount || 1; // Corrected count field
                
                if (decreaseInventory(inventoryData[inventoryType], sizeToAllocate, requiredCount, inventoryType)) {
                    student[allocatedField] = sizeToAllocate; // Store the actual allocated size (e.g., "M/38")
                    if (adjustmentDisplayMark) {
                        student.pantsAdjustmentMark = adjustmentDisplayMark; // Store "M/38↑"
                    } else {
                        student.pantsAdjustmentMark = null; // Ensure it's cleared if no adjustment
                    }
                    
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    successCount++;
                    console.log(`分配成功 (${inventoryType} 新規則): 學生 [${student.id}] ${student.class || ''}-${student.number || ''} ${student.name || ''} - 實際分配尺碼 ${sizeToAllocate}${adjustmentDisplayMark ? ' (顯示標記: ' + adjustmentDisplayMark + ')' : ''}, 需求 ${requiredCount} 件`);
                } else {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = `分配失敗：尺碼(${sizeToAllocate})庫存不足`;
                    failedCount++;
                    console.warn(`分配失敗 (${inventoryType} 新規則): 學生 [${student.id}] ${student.class || ''}-${student.number || ''} ${student.name || ''} - 尺碼 ${sizeToAllocate} 庫存不足, 需求 ${requiredCount} 件`);
                }
            }
        }
        
        console.log(`${UNIFORM_TYPES[inventoryType]} (新規則) 分配完成，成功 ${successCount} 位，失敗 ${failedCount} 位 ====================`);
        // Update overall allocation stats if necessary, though this function is per type
        allocationStats[inventoryType].allocated = successCount;
        allocationStats[inventoryType].failed = failedCount; 
        // Note: 'exact', 'different', 'special' might need re-evaluation if used with this new logic.
        // For now, just updating allocated and failed.

        resolve(true); // Or resolve(successCount > 0);
    });
}

/**
 * 分配長袖上衣
 */
function allocateLongShirts(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`%c===== 開始分配 ${UNIFORM_TYPES[inventoryType]} =====`, 'background: #9b59b6; color: white; font-size: 14px; padding: 5px;');
        
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            _localSortedStudentData.forEach(student => {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '無庫存資料';
                student[allocatedField] = ''; // 確保清空
                student.longShirtAllocationMark = ''; // 新增：清空標記
            });
            allocationStats[inventoryType].failed = _localSortedStudentData.length;
            resolve(false);
            return;
        }
        
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.longSleeveShirtCount || 1; // 注意這裡是 longSleeveShirtCount
        });
        
        let totalAllocatable = 0;
        for (const size in inventoryData[inventoryType]) {
            totalAllocatable += inventoryData[inventoryType][size]?.allocatable || 0;
        }
        
        console.log(`%c${UNIFORM_TYPES[inventoryType]}需求與庫存概況：`, 'color: #8e44ad; font-weight: bold;');
        console.log(`- 總需求量：${totalDemand}件`);
        console.log(`- 總可分配數量：${totalAllocatable}件`);
        console.log(`- 學生數量：${_localSortedStudentData.length}人`);
        
        if (totalAllocatable < totalDemand) {
            console.warn(`%c警告：${UNIFORM_TYPES[inventoryType]}可分配數總和(${totalAllocatable})小於需求量(${totalDemand})，差額${totalDemand - totalAllocatable}件`, 'color: #e74c3c; font-weight: bold;');
        }
        
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        const sortedStudents = [..._localSortedStudentData].map(student => {
            const chest = student.chest || 0;
            const waist = student.waist || 0;
            const effectiveChest = Math.max(chest, waist);
            return { student, effectiveChest };
        }).sort((a, b) => a.effectiveChest - b.effectiveChest);

        console.log(`%c學生排序結果（按有效胸圍升序）：`, 'color: #8e44ad; font-weight: bold;');
        // sortedStudents.slice(0, 5).forEach((s, i) => { ... }); // 日誌可以簡化或保留

        const stats = { allocated: 0, exact: 0, different: 0, failed: 0, special: 0, pantsSizeAdjusted: 0, preferredFailedSwitchToAlternative: 0, femaleDowngraded: 0 };
        
        for (const {student, effectiveChest} of sortedStudents) {
            if (student[allocatedField]) {
                console.log(`學生 ${student.name}(${student.class}-${student.number}) 已有長衣分配結果：${student[allocatedField]}，跳過`);
                continue;
            }
            
            const requiredCount = student.longSleeveShirtCount || 1;
            student.longShirtAllocationMark = ''; 

            console.log(`%c處理長衣分配給學生：${student.name}(${student.gender}，班級：${student.class}，座號：${student.number})`, 'color: #8e44ad; font-weight: bold;');
            console.log(`- 胸圍：${student.chest}，腰圍：${student.waist}，褲長：${student.pantsLength}，有效胸圍：${effectiveChest}，需求：${requiredCount}`);

            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生不需要長衣，跳過`);
                stats.failed++;
                continue;
            }
            
            let preferredAdjustment, alternativeAdjustment;
            if (student.gender === '男') {
                preferredAdjustment = (effectiveChest % 2 === 0) ? 10 : 11;
                alternativeAdjustment = (effectiveChest % 2 === 0) ? 12 : 13;
            } else {
                preferredAdjustment = (effectiveChest % 2 === 0) ? 8 : 9;
                alternativeAdjustment = (effectiveChest % 2 === 0) ? 10 : 11;
            }

            const preferredTargetSizeNumber = effectiveChest + preferredAdjustment;
            const alternativeTargetSizeNumber = effectiveChest + alternativeAdjustment;

            function getTargetSizeFromNumber(sizeNumber) {
                if (sizeNumber <= 34) return 'XS/34';
                if (sizeNumber <= 36) return 'S/36';
                if (sizeNumber <= 38) return 'M/38';
                if (sizeNumber <= 40) return 'L/40';
                if (sizeNumber <= 42) return 'XL/42';
                if (sizeNumber <= 44) return '2L/44';
                if (sizeNumber <= 46) return '3L/46';
                if (sizeNumber <= 48) return '4L/48';
                if (sizeNumber <= 50) return '5L/50';
                if (sizeNumber <= 52) return '6L/52';
                if (sizeNumber <= 54) return '7L/54';
                if (sizeNumber <= 56) return '8L/56';
                return '9L/58';
            }

            const preferredTargetSize = getTargetSizeFromNumber(preferredTargetSizeNumber);
            const alternativeTargetSize = getTargetSizeFromNumber(alternativeTargetSizeNumber);

            console.log(`- 理論首選長衣尺碼: ${preferredTargetSize} (基於 ${preferredTargetSizeNumber})`);
            console.log(`- 理論備選長衣尺碼: ${alternativeTargetSize} (基於 ${alternativeTargetSizeNumber})`);

            let finalAllocatedSize = null;
            let allocationMark = '';
            let attemptLog = [];

            // Attempt 1: Preferred Size
            attemptLog.push(`嘗試首選長衣尺碼 ${preferredTargetSize}`);
            let preferredPathMadePantsAdjustment = false; // 新增布林變數追蹤首選路徑的褲長調整

            if (remainingInventory[preferredTargetSize]?.allocatable >= requiredCount) {
                let currentAttemptSize = preferredTargetSize;
                attemptLog.push(`首選長衣尺碼 ${currentAttemptSize} 有庫存 (${remainingInventory[preferredTargetSize].allocatable})`);

                // Pants length adjustment for preferred size (long shirts)
                if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, currentAttemptSize)) {
                    const originalSizeBeforeAdjustment = currentAttemptSize;
                    const largerSize = getNextLargerSize(originalSizeBeforeAdjustment);
                    attemptLog.push(`長衣褲長 (${student.pantsLength}) vs ${getSizeNumber(originalSizeBeforeAdjustment)}，建議調整到 ${largerSize}`);

                    if (remainingInventory[largerSize]?.allocatable >= requiredCount) {
                        if (largerSize !== originalSizeBeforeAdjustment) { // Key check: Ensure the size actually changes
                            currentAttemptSize = largerSize; // Update currentAttemptSize to the new, larger size
                            stats.pantsSizeAdjusted++;
                            preferredPathMadePantsAdjustment = true; // Set flag: successful adjustment to a DIFFERENT, larger size
                            attemptLog.push(`長衣褲長調整成功，使用 ${currentAttemptSize} (庫存 ${remainingInventory[currentAttemptSize].allocatable})`);
                } else {
                            attemptLog.push(`長衣褲長建議調整，但 ${originalSizeBeforeAdjustment} 已是最大可獲取尺碼或無更大尺碼。繼續使用 ${originalSizeBeforeAdjustment}。`);
                        }
                } else {
                        attemptLog.push(`長衣褲長調整失敗：建議的尺碼 ${largerSize} 庫存不足 (${remainingInventory[largerSize]?.allocatable || 0})。`);
                    }
                }
                // After potential adjustment, re-check stock for the (potentially updated) currentAttemptSize
                if (remainingInventory[currentAttemptSize]?.allocatable >= requiredCount) {
                    finalAllocatedSize = currentAttemptSize;
                    if (preferredPathMadePantsAdjustment) { // This is now the sole, correct condition for the mark
                        allocationMark = '↑';
                    }
                } else {
                    attemptLog.push(`最終檢查：長衣尺碼 ${currentAttemptSize} (可能是原始首選或經調整的尺碼) 庫存不足。首選尺碼路徑失敗。`);
                    finalAllocatedSize = null;
                    preferredPathMadePantsAdjustment = false;
                }
            } else {
                attemptLog.push(`首選長衣尺碼 ${preferredTargetSize} 庫存不足或不存在 (${remainingInventory[preferredTargetSize]?.allocatable || 0})`);
            }

            if (!finalAllocatedSize) {
                attemptLog.push(`首選長衣失敗，嘗試備選尺碼 ${alternativeTargetSize}`);
                stats.preferredFailedSwitchToAlternative++;
                if (remainingInventory[alternativeTargetSize]?.allocatable >= requiredCount) {
                    let currentAttemptSize = alternativeTargetSize;
                    allocationMark = '↑';
                    attemptLog.push(`備選長衣尺碼 ${currentAttemptSize} 有庫存 (${remainingInventory[currentAttemptSize].allocatable})`);
                    if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, currentAttemptSize)) {
                        const largerSize = getNextLargerSize(currentAttemptSize);
                        attemptLog.push(`長衣褲長 (${student.pantsLength}) vs ${getSizeNumber(currentAttemptSize)}，建議調整到 ${largerSize}`);
                        if (remainingInventory[largerSize]?.allocatable >= requiredCount) {
                            currentAttemptSize = largerSize;
                    stats.pantsSizeAdjusted++;
                            attemptLog.push(`長衣褲長調整成功，使用 ${currentAttemptSize} (庫存 ${remainingInventory[currentAttemptSize].allocatable})`);
                } else {
                            attemptLog.push(`長衣褲長調整失敗：${largerSize} 庫存不足 (${remainingInventory[largerSize]?.allocatable || 0})。備選尺碼路徑失敗。`);
                            allocationMark = '';
                        }
                    }
                    if (currentAttemptSize === alternativeTargetSize || (currentAttemptSize !== alternativeTargetSize && remainingInventory[currentAttemptSize]?.allocatable >= requiredCount)) {
                        finalAllocatedSize = currentAttemptSize;
                    }
                } else {
                    attemptLog.push(`備選長衣尺碼 ${alternativeTargetSize} 庫存不足或不存在 (${remainingInventory[alternativeTargetSize]?.allocatable || 0})`);
                    allocationMark = '';
                }
            }
            
            if (finalAllocatedSize && student.gender === '女' && getSizeNumber(finalAllocatedSize) >= 44 && student.pantsLength <= 38) {
                attemptLog.push(`觸發女生長衣特殊調整：尺碼 ${finalAllocatedSize} (>=44), 褲長 ${student.pantsLength} (<=38)`);
                const smallerSize = getPreviousSmallerSize(finalAllocatedSize);
                attemptLog.push(`嘗試縮小長衣到尺碼 ${smallerSize}`);
                if (remainingInventory[smallerSize]?.allocatable >= requiredCount) {
                    finalAllocatedSize = smallerSize;
                    allocationMark = '↓'; 
                    stats.femaleDowngraded++;
                    attemptLog.push(`女生長衣特殊調整成功，使用 ${finalAllocatedSize} (庫存 ${remainingInventory[finalAllocatedSize].allocatable})`);
            } else {
                    attemptLog.push(`女生長衣特殊調整失敗：縮小後尺碼 ${smallerSize} 庫存不足 (${remainingInventory[smallerSize]?.allocatable || 0})。分配失敗。`);
                    finalAllocatedSize = null; 
                    allocationMark = '';
                }
            }

            console.log('長衣分配嘗試日誌:', attemptLog.join(' -> '));
            if (finalAllocatedSize) {
                student[allocatedField] = finalAllocatedSize;
                student.longShirtAllocationMark = allocationMark; 
                decreaseInventory(remainingInventory, finalAllocatedSize, requiredCount, inventoryType);
                stats.allocated++;
                if (allocationMark === '↑') stats.different++;
                else if (allocationMark === '↓') stats.special++;
                else stats.exact++;
                console.log(`%c長衣分配成功：${student.name} => ${finalAllocatedSize}${allocationMark} (需求 ${requiredCount})`, 'color: #27ae60; font-weight: bold;');
            if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                delete student.allocationFailReason[inventoryType];
                }
            } else {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = student.allocationFailReason[inventoryType] || '分配失敗：長衣無合適尺碼或庫存不足';
                 if (attemptLog.length > 0) {
                    student.allocationFailReason[inventoryType] = `長衣分配失敗(${attemptLog[attemptLog.length-1]})`;
                }
                stats.failed++;
                console.log(`%c長衣分配失敗：${student.name} (${student.allocationFailReason[inventoryType]})`, 'color: #e74c3c; font-weight: bold;');
            }
        }

        inventoryData[inventoryType] = remainingInventory; 
        allocationStats[inventoryType] = {
            allocated: stats.allocated,
            exact: stats.exact,
            different: stats.different,
            special: stats.special,
            failed: stats.failed,
            pantsSizeAdjusted: stats.pantsSizeAdjusted,
            preferredFailedSwitchToAlternative: stats.preferredFailedSwitchToAlternative,
            femaleDowngraded: stats.femaleDowngraded
        };

        console.log(`%c${UNIFORM_TYPES[inventoryType]}分配結果統計：`, 'background: #9b59b6; color: white; font-size: 12px; padding: 5px;');
        console.log(`- 成功分配：${stats.allocated}人`);
        console.log(`  - 精確分配 (首選無調整)：${stats.exact}人`);
        console.log(`  - 不同尺寸分配 (含備選/褲長調整)：${stats.different}人`);
        console.log(`  - 特殊調整 (女生降碼)：${stats.special}人`);
        console.log(`- 分配失敗：${stats.failed}人`);
        console.log(`- 因褲長調整次數：${stats.pantsSizeAdjusted}`);
        console.log(`- 首選失敗轉備選次數：${stats.preferredFailedSwitchToAlternative}`);
        console.log(`- 女生降碼次數：${stats.femaleDowngraded}`);
        console.log(`%c===== ${UNIFORM_TYPES[inventoryType]}分配完成 =====`, 'background: #9b59b6; color: white; font-size: 14px; padding: 5px;');
        
        resolve(true);
    });
}

/**
 * 分配長袖褲子
 */
function allocateLongPants(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]} (新規則) ====================`);
        
        _localSortedStudentData.forEach(student => {
            student.longPantsAdjustmentMark = null; // 清除長褲調整標記
        });
        console.log(`已清除所有學生的長褲調整標記 (新規則)`);
        
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.longSleevePantsCount || 1; // 使用長褲需求數量
        });
        
        let totalAllocatable = 0;
        for (const size in inventoryData[inventoryType]) {
            totalAllocatable += inventoryData[inventoryType][size].allocatable || 0;
        }
        
        console.log(`${UNIFORM_TYPES[inventoryType]}：總需求量=${totalDemand}，總可分配數量=${totalAllocatable}`);
        
        // 如果庫存不足，顯示警告但繼續分配
        if (totalAllocatable < totalDemand) {
            console.warn(`${UNIFORM_TYPES[inventoryType]} 庫存不足！總需求量=${totalDemand}，總可分配數量=${totalAllocatable}`);
        }
        
        // 檢查是否有任何可分配的尺碼
        let hasAllocatableSizes = false;
        for (const size in inventoryData[inventoryType]) {
            if (inventoryData[inventoryType][size].allocatable > 0) {
                hasAllocatableSizes = true;
                break;
            }
        }
        
        if (!hasAllocatableSizes) {
            console.warn(`${UNIFORM_TYPES[inventoryType]} 沒有可分配的尺碼！`);
            resolve(false);
            return;
        }
        
        // 輔助函數：在庫存中查找匹配的尺碼 - 只使用精確匹配
        function findMatchingInventorySize(targetSize) {
            // 使用精確匹配
            if (inventoryData[inventoryType][targetSize] && 
                inventoryData[inventoryType][targetSize].allocatable > 0) {
                return targetSize;
            }
            
            console.log(`未找到尺碼 "${targetSize}" 的庫存或庫存為零`);
            return null;
        }
        
        // 腰圍範圍到尺碼的映射表 - 新的長褲映射邏輯
        const waistToSizeMap = [
            { 
                min: 20, max: 21, size: "XS/34", 
                adjustedSize: "S/36", 
                adjustCondition: (student) => student.pantsLength >= 35, 
                adjustMark: "S/36↑" 
            }, 
            { 
                min: 22, max: 24, size: "S/36",
                adjustedSizeOne: "M/38", 
                adjustConditionOne: (student) => student.pantsLength >= 37 && student.pantsLength < 39,
                adjustMarkOne: "M/38↑",
                adjustedSizeTwo: "L/40", 
                adjustConditionTwo: (student) => student.pantsLength >= 39,
                adjustMarkTwo: "L/40↑2"
            },
            { 
                min: 25, max: 27, size: "M/38", 
                adjustedSizeOne: "L/40", 
                adjustConditionOne: (student) => student.pantsLength >= 39 && student.pantsLength <= 40,
                adjustMarkOne: "L/40↑",
                adjustedSizeTwo: "XL/42", 
                adjustConditionTwo: (student) => student.pantsLength >= 41,
                adjustMarkTwo: "XL/42↑2"
            },
            { 
                min: 28, max: 30, size: "L/40", 
                adjustedSizeOne: "XL/42", 
                adjustConditionOne: (student) => (student.gender === "男" && student.pantsLength >= 41 && student.pantsLength <= 42) || 
                                               (student.gender === "女" && student.pantsLength >= 41),
                adjustMarkOne: "XL/42↑",
                adjustedSizeTwo: "2L/44", 
                adjustConditionTwo: (student) => student.gender === "男" && student.pantsLength >= 43,
                adjustMarkTwo: "2L/44↑2"
            },
            { 
                min: 31, max: 33, size: "XL/42",
                adjustedSizeOne: "2L/44", 
                adjustConditionOne: (student) => student.gender === "男" && student.pantsLength >= 42 && student.pantsLength <= 43,
                adjustMarkOne: "2L/44↑",
                adjustedSizeTwo: "3L/46", 
                adjustConditionTwo: (student) => student.gender === "男" && student.pantsLength >= 44,
                adjustMarkTwo: "3L/46↑2",
                downAdjustedSize: "L/40", 
                downAdjustCondition: (student) => student.gender === "女" && student.pantsLength <= 40,
                downAdjustMark: "L/40↓"
            },
            { 
                min: 34, max: 36, size: "2L/44",
                downAdjustedSize: "XL/42", 
                downAdjustCondition: (student) => student.gender === "女" && student.pantsLength <= 40,
                downAdjustMark: "XL/42↓"
            },
            { 
                min: 37, max: 38, size: "3L/46",
                downAdjustedSize: "2L/44", 
                downAdjustCondition: (student) => student.gender === "女" && student.pantsLength <= 40,
                downAdjustMark: "2L/44↓"
            },
            { 
                min: 39, max: 40, size: "4L/48",
                downAdjustedSize: "3L/46", 
                downAdjustCondition: (student) => student.gender === "女" && student.pantsLength <= 40,
                downAdjustMark: "3L/46↓"
            },
            { 
                min: 41, max: 43, primarySize: "5L/50", alternateSize: "6L/52",
                downAdjustedSize: "4L/48", 
                downAdjustCondition: (student) => student.gender === "女" && student.pantsLength <= 40,
                downAdjustMark: "4L/48↓"
            },
            { 
                min: 44, max: 46, primarySize: "7L/54", alternateSize: "8L/56",
                downAdjustedSize: "6L/52", 
                downAdjustCondition: (student) => student.gender === "女" && student.pantsLength <= 40,
                downAdjustMark: "6L/52↓"
            }
        ];
        
        // 成功分配的學生計數
        let successCount = 0;
        // 失敗的學生計數
        let failedCount = 0;
        
        // 單階段分配過程 - 直接分配
        for (const student of _localSortedStudentData) {
            // 檢查學生是否需要分配長褲
            if (student.longSleevePantsCount && student.longSleevePantsCount > 0) {
                console.log(`\n分配學生 [${student.id}] ${student.class}-${student.number} ${student.name}: 腰圍=${student.waist}, 褲長=${student.pantsLength}, 性別=${student.gender}`);
                
                // 根據腰圍找到對應的尺碼範圍
                const waistRange = waistToSizeMap.find(range => 
                    student.waist >= range.min && student.waist <= range.max);
                
                if (!waistRange) {
                    console.warn(`學生 [${student.id}] 腰圍 ${student.waist} 超出分配範圍，無法分配長褲`);
                    // 設置失敗原因
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：腰圍超出範圍';
                    failedCount++;
                    continue;
                }
                
                // 判斷是否需要根據性別和褲長進行調整
                let sizeToAllocate;
                let adjustmentMark = null;
                
                // 首先檢查是否需要向下調整（女生且褲長≤40）
                if (waistRange.downAdjustCondition && waistRange.downAdjustCondition(student)) {
                    console.log(`學生 [${student.id}] 符合向下調整條件: 原尺碼=${waistRange.size || waistRange.primarySize}, 調整後尺碼=${waistRange.downAdjustedSize}`);
                    
                    // 檢查調整後的尺碼是否有庫存
                    const downAdjustedSizeAvailable = findMatchingInventorySize(waistRange.downAdjustedSize);
                    
                    if (downAdjustedSizeAvailable) {
                        sizeToAllocate = downAdjustedSizeAvailable;
                        adjustmentMark = waistRange.downAdjustMark;
                    } else {
                        // 調整後尺碼無庫存，標記為分配失敗
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '分配失敗：調整尺碼無庫存';
                        failedCount++;
                        console.warn(`學生 [${student.id}] 調整後尺碼 ${waistRange.downAdjustedSize} 無庫存，分配失敗`);
                        continue;
                    }
                }
                // 檢查是否需要向上調整2號
                else if (waistRange.adjustConditionTwo && waistRange.adjustConditionTwo(student)) {
                    console.log(`學生 [${student.id}] 符合向上調整2號條件: 原尺碼=${waistRange.size}, 調整後尺碼=${waistRange.adjustedSizeTwo}`);
                    
                    // 檢查調整後的尺碼是否有庫存
                    const adjustedSizeTwoAvailable = findMatchingInventorySize(waistRange.adjustedSizeTwo);
                    
                    if (adjustedSizeTwoAvailable) {
                        sizeToAllocate = adjustedSizeTwoAvailable;
                        adjustmentMark = waistRange.adjustMarkTwo;
                    } else {
                        // 調整後尺碼無庫存，標記為分配失敗
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '分配失敗：調整尺碼無庫存';
                        failedCount++;
                        console.warn(`學生 [${student.id}] 調整後尺碼 ${waistRange.adjustedSizeTwo} 無庫存，分配失敗`);
                        continue;
                    }
                }
                // 檢查是否需要向上調整1號
                else if (waistRange.adjustConditionOne && waistRange.adjustConditionOne(student)) {
                    console.log(`學生 [${student.id}] 符合向上調整1號條件: 原尺碼=${waistRange.size}, 調整後尺碼=${waistRange.adjustedSizeOne}`);
                    
                    // 檢查調整後的尺碼是否有庫存
                    const adjustedSizeOneAvailable = findMatchingInventorySize(waistRange.adjustedSizeOne);
                    
                    if (adjustedSizeOneAvailable) {
                        sizeToAllocate = adjustedSizeOneAvailable;
                        adjustmentMark = waistRange.adjustMarkOne;
                    } else {
                        // 調整後尺碼無庫存，標記為分配失敗
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '分配失敗：調整尺碼無庫存';
                        failedCount++;
                        console.warn(`學生 [${student.id}] 調整後尺碼 ${waistRange.adjustedSizeOne} 無庫存，分配失敗`);
                        continue;
                    }
                }
                // 需要進行常規調整
                else if (waistRange.adjustCondition && waistRange.adjustCondition(student)) {
                    console.log(`學生 [${student.id}] 符合特殊調整條件: 原尺碼=${waistRange.size}, 調整後尺碼=${waistRange.adjustedSize}`);
                    
                    // 檢查調整後的尺碼是否有庫存
                    const adjustedSizeAvailable = findMatchingInventorySize(waistRange.adjustedSize);
                    
                    if (adjustedSizeAvailable) {
                        sizeToAllocate = adjustedSizeAvailable;
                        adjustmentMark = waistRange.adjustMark;
                    } else {
                        // 調整後尺碼無庫存，標記為分配失敗
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '分配失敗：調整尺碼無庫存';
                        failedCount++;
                        console.warn(`學生 [${student.id}] 調整後尺碼 ${waistRange.adjustedSize} 無庫存，分配失敗`);
                        continue;
                    }
                }
                // 如果是需要先嘗試主要尺碼再嘗試替代尺碼的情況
                else if (waistRange.primarySize) {
                    const primarySizeAvailable = findMatchingInventorySize(waistRange.primarySize);
                    
                    if (primarySizeAvailable) {
                        console.log(`學生 [${student.id}] 使用主要尺碼: ${waistRange.primarySize}`);
                        sizeToAllocate = primarySizeAvailable;
                    } else {
                        const alternateSizeAvailable = findMatchingInventorySize(waistRange.alternateSize);
                        if (alternateSizeAvailable) {
                            console.log(`學生 [${student.id}] 主要尺碼 ${waistRange.primarySize} 無庫存，使用替代尺碼: ${waistRange.alternateSize}`);
                            sizeToAllocate = alternateSizeAvailable;
                        } else {
                            console.warn(`學生 [${student.id}] 主要尺碼 ${waistRange.primarySize} 和替代尺碼 ${waistRange.alternateSize} 均無庫存`);
                            // 設置失敗原因
                            student.allocationFailReason = student.allocationFailReason || {};
                            student.allocationFailReason[inventoryType] = '分配失敗：無可用尺碼';
                            failedCount++;
                            continue;
                        }
                    }
                } 
                // 一般情況
                else {
                    const sizeAvailable = findMatchingInventorySize(waistRange.size);
                    if (sizeAvailable) {
                        console.log(`學生 [${student.id}] 使用標準尺碼: ${waistRange.size}`);
                        sizeToAllocate = sizeAvailable;
                    } else {
                        console.warn(`學生 [${student.id}] 尺碼 ${waistRange.size} 無庫存`);
                        // 設置失敗原因
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '分配失敗：尺碼無庫存';
                        failedCount++;
                        continue;
                    }
                }
                
                // 確保找到了可用尺碼
                if (!sizeToAllocate) {
                    console.warn(`學生 [${student.id}] 無法分配長褲：找不到合適的尺碼或庫存不足`);
                    // 設置失敗原因
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：無合適尺碼';
                    failedCount++;
                    continue;
                }
                
                // 檢查庫存是否足夠並直接分配
                const requiredCount = student.longSleevePantsCount || 1;
                
                // 直接嘗試減少庫存並分配
                if (decreaseInventory(inventoryData[inventoryType], sizeToAllocate, requiredCount, inventoryType)) {
                    // 分配成功，更新學生資料
                    student[allocatedField] = sizeToAllocate;
                    if (adjustmentMark) {
                        student.longPantsAdjustmentMark = adjustmentMark;
                    }
                    
                    // 清除失敗原因（如果之前有）
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    
                    successCount++;
                    console.log(`分配成功: 學生 [${student.id}] ${student.class}-${student.number} ${student.name} - 尺碼 ${sizeToAllocate}${adjustmentMark ? ' (標記: ' + adjustmentMark + ')' : ''}, 需求 ${requiredCount} 件`);
                } else {
                    // 分配失敗 - 庫存不足
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：庫存不足';
                    failedCount++;
                    console.warn(`分配失敗: 學生 [${student.id}] ${student.class}-${student.number} ${student.name} - 尺碼 ${sizeToAllocate} 庫存不足, 需求 ${requiredCount} 件`);
                }
            }
        }
        
        // 更新分配統計
        allocationStats[inventoryType] = {
            allocated: successCount,
            exact: successCount, // 在這個簡化的模型中，所有成功分配都視為精確匹配
            different: 0,
            failed: failedCount,
            special: 0,
            pantsSizeAdjusted: 0 // 不再使用這個標記
        };
        
        console.log(`${UNIFORM_TYPES[inventoryType]} 分配完成，共分配給 ${successCount} 位學生，${failedCount} 位學生分配失敗 ====================`);
        resolve(true);
    });
}

/**
 * 檢查學生分配到的尺碼是否仍比褲長小2或更多
 * 同時檢查上衣和褲子是否符合褲長要求
 */
function checkPantsLengthDeficiency() {
    console.log('開始檢查尺碼與褲長不足情況');
    
    // 計數器
    let shortShirtDeficiencyCount = 0;
    let shortPantsDeficiencyCount = 0;
    let longShirtDeficiencyCount = 0;
    let longPantsDeficiencyCount = 0;
    
    // 檢查每個學生
    _localSortedStudentData.forEach(student => {
        // 清除之前可能存在的標記
        student.shortShirtLengthDeficiency = false;
        student.shortPantsLengthDeficiency = false;
        student.longShirtLengthDeficiency = false;
        student.longPantsLengthDeficiency = false;
        
        // 檢查短衣尺碼與褲長差距
        if (student.allocatedShirtSize && student.pantsLength) {
            const shortShirtSizeNumber = getSizeNumber(student.allocatedShirtSize);
            if ((student.pantsLength - shortShirtSizeNumber) >= 2) {
                student.shortShirtLengthDeficiency = true;
                shortShirtDeficiencyCount++;
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 短衣尺碼 ${student.allocatedShirtSize}(${shortShirtSizeNumber}) 比褲長 ${student.pantsLength} 小 ${student.pantsLength - shortShirtSizeNumber}`);
            }
        }
        
        // 檢查短褲尺碼與褲長差距
        if (student.allocatedPantsSize && student.pantsLength) {
            const shortPantsSizeNumber = getSizeNumber(student.allocatedPantsSize);
            if ((student.pantsLength - shortPantsSizeNumber) >= 2) {
                // student.shortPantsLengthDeficiency = true;
                shortPantsDeficiencyCount++;
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 短褲尺碼 ${student.allocatedPantsSize}(${shortPantsSizeNumber}) 比褲長 ${student.pantsLength} 小 ${student.pantsLength - shortPantsSizeNumber}`);
            }
        }
        
        // 檢查長衣尺碼與褲長差距
        if (student.allocatedLongShirtSize && student.pantsLength) {
            const longShirtSizeNumber = getSizeNumber(student.allocatedLongShirtSize);
            if ((student.pantsLength - longShirtSizeNumber) >= 2) {
                student.longShirtLengthDeficiency = true;
                longShirtDeficiencyCount++;
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 長衣尺碼 ${student.allocatedLongShirtSize}(${longShirtSizeNumber}) 比褲長 ${student.pantsLength} 小 ${student.pantsLength - longShirtSizeNumber}`);
            }
        }
        
        // 檢查長褲尺碼與褲長差距
        if (student.allocatedLongPantsSize && student.pantsLength) {
            const longPantsSizeNumber = getSizeNumber(student.allocatedLongPantsSize);
            if ((student.pantsLength - longPantsSizeNumber) >= 3) {
                student.longPantsLengthDeficiency = true;
                longPantsDeficiencyCount++;
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 長褲尺碼 ${student.allocatedLongPantsSize}(${longPantsSizeNumber}) 比褲長 ${student.pantsLength} 小 ${student.pantsLength - longPantsSizeNumber}`);
            }
        }
    });
    
    console.log(`尺碼與褲長不足檢查完成：短衣不足 ${shortShirtDeficiencyCount} 人，短褲不足 ${shortPantsDeficiencyCount} 人，長衣不足 ${longShirtDeficiencyCount} 人，長褲不足 ${longPantsDeficiencyCount} 人`);
}

/**
 * 更新分配結果頁面
 */
export function updateAllocationResults() {
    console.log('開始更新分配結果頁面');
    
    // 檢查褲長不足情況
    checkPantsLengthDeficiency();
    
    // 更新分配統計
    updateAllocationStats();
    
    // 更新學生詳細分配結果（移到尺寸分配結果之前）
    updateStudentDetailedResults();
    
    // 更新庫存分配結果
    updateInventoryAllocationResults();
    
    console.log('分配結果頁面更新完成');
}

/**
 * 更新學生詳細分配結果
 */
function updateStudentDetailedResults() {
    // 尋找結果頁面容器
    const resultTab = document.getElementById('result');
    if (!resultTab) return;

    // 尋找尺寸分配結果表格作為參考點
    const shortSleeveShirtTable = document.getElementById('shortSleeveShirtResultTable');
    if (!shortSleeveShirtTable) return;

    // 尋找或創建詳細結果表格
    let detailTable = document.getElementById('studentDetailTable');
    if (!detailTable) {
        // 創建表格區域
        const detailSection = document.createElement('div');
        detailSection.className = 'mt-4 mb-4';
        detailSection.innerHTML = `
            <style>
                #studentDetailTable .count-column {
                    width: 60px;
                    text-align: center;
                }
                .failure-reason {
                    color: #dc3545;
                    font-style: italic;
                    font-size: 0.9em;
                }
                .adjustment-reason {
                    color: #17a2b8;
                    font-style: italic;
                    font-size: 0.9em;
                }
                .allocation-mark {
                    font-weight: bold;
                    color: #007bff; /* 藍色用於標記 */
                }
            </style>
            <div class="d-flex justify-content-between align-items-center mb-2">
                <h4>學生分配詳細結果</h4>
                <button id="exportAllocationResultsBtn" class="btn btn-success">
                    <i class="bi bi-file-excel me-1"></i>匯出Excel
                </button>
            </div>
            <div class="table-responsive">
                <table id="studentDetailTable" class="table table-striped table-bordered">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>班級</th>
                            <th>號碼</th>
                            <th>姓名</th>
                            <th>性別</th>
                            <th>胸圍</th>
                            <th>腰圍</th>
                            <th>褲長</th>
                            <th>短衣尺寸</th>
                            <th class="count-column">件數</th>
                            <th>短褲尺寸</th>
                            <th class="count-column">件數</th>
                            <th>長衣尺寸</th>
                            <th class="count-column">件數</th>
                            <th>長褲尺寸</th>
                            <th class="count-column">件數</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        `;
        
        const cardBody = resultTab.querySelector('.card-body');
        if (cardBody) {
            cardBody.insertBefore(detailSection, cardBody.firstChild);
        } else {
            resultTab.insertBefore(detailSection, resultTab.firstChild);
        }
        
        detailTable = document.getElementById('studentDetailTable');
        
        const exportBtn = document.getElementById('exportAllocationResultsBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', exportAllocationResultsToExcel);
        }
    }

    const tbody = detailTable.querySelector('tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const sortedStudents = [...studentData].sort((a, b) => {
        const classA = parseInt(a.class) || 0; // 使用 a.class 而非 a.className
        const classB = parseInt(b.class) || 0; // 使用 b.class 而非 b.className
        if (classA !== classB) return classA - classB;
        
        const numberA = parseInt(a.number) || 0;
        const numberB = parseInt(b.number) || 0;
        return numberA - numberB;
    });

    sortedStudents.forEach((student, index) => {
        const row = document.createElement('tr');
        
        // DEBUGGING CODE BLOCK START
        if (student.name === '郭柏瑜' && 
            (String(student.class) === '101') && 
            (String(student.number) === '8' || String(student.number).padStart(2, '0') === '08')) {
            console.log('[DEBUG] updateStudentDetailedResults for 郭柏瑜 (宽松条件):');
            console.log('  student object:', JSON.parse(JSON.stringify(student))); // Log a snapshot
            console.log('  student.shirtAllocationMark:', student.shirtAllocationMark);
            console.log('  student.longShirtAllocationMark:', student.longShirtAllocationMark);
        }
        // DEBUGGING CODE BLOCK END

        const shortShirtFailReason = student.allocationFailReason?.shortSleeveShirt || '';
        const shortPantsFailReason = student.allocationFailReason?.shortSleevePants || '';
        const longShirtFailReason = student.allocationFailReason?.longSleeveShirt || '';
        const longPantsFailReason = student.allocationFailReason?.longSleevePants || '';
        
        // 獲取新的分配標記
        let shirtMark = student.shirtAllocationMark || '';
        const longShirtMark = student.longShirtAllocationMark || '';

        // REMEDIAL ACTION for missing shirtMark when pants adjustment occurred
        if (student.allocatedShirtSize && !shirtMark && student.isShirtSizeAdjustedForPantsLength) {
            // This case implies that isShirtSizeAdjustedForPantsLength was set (likely in allocateShortShirts)
            // but the shirtAllocationMark itself was lost before rendering.
            // We assume if isShirtSizeAdjustedForPantsLength is true and no mark exists, it should have been '↑'.
            // This is a targeted fix for cases like Guo Boyu where the mark disappears.
            // It won't apply if shirtMark was deliberately set to '↓' (female downgrade) or was already '↑'.
            console.warn(`[REMEDY] Student ${student.name} (${student.class}-${String(student.number).padStart(2,'0')}) has isShirtSizeAdjustedForPantsLength=true but an empty/undefined shirtAllocationMark. Applying '↑' for display.`);
            shirtMark = '↑';
        }
        
        // 先處理尺碼中的星號問題
        // Determine display marks for pants, prioritizing '*' if present in allocated string
        let displayShortPantsMark = '';
        let cleanShortPantsSize = student.allocatedPantsSize;
        if (student.allocatedPantsSize) {
            if (student.allocatedPantsSize.endsWith('*')) {
                displayShortPantsMark = '*';
                // 移除原始尺碼中的星號，以避免重複顯示
                cleanShortPantsSize = student.allocatedPantsSize.slice(0, -1);
            } else if (student.pantsAdjustmentMark) {
                displayShortPantsMark = student.pantsAdjustmentMark;
            }
        }

        let displayLongPantsMark = '';
        let cleanLongPantsSize = student.allocatedLongPantsSize;
        if (student.allocatedLongPantsSize) {
            if (student.allocatedLongPantsSize.endsWith('*')) {
                displayLongPantsMark = '*';
                // 移除原始尺碼中的星號，以避免重複顯示
                cleanLongPantsSize = student.allocatedLongPantsSize.slice(0, -1);
            } else if (student.longPantsAdjustmentMark) {
                displayLongPantsMark = student.longPantsAdjustmentMark;
            }
        }
        
        // 在取得清理後的尺碼後，再進行格式化
        const formattedShirtSize = student.allocatedShirtSize ? formatSize(student.allocatedShirtSize) : '-';
        // 使用已清除星號的短褲尺碼來格式化
        const formattedPantsSize = student.allocatedPantsSize ? formatSize(cleanShortPantsSize) : '-';
        const formattedLongShirtSize = student.allocatedLongShirtSize ? formatSize(student.allocatedLongShirtSize) : '-';
        // 使用已清除星號的長褲尺碼來格式化
        const formattedLongPantsSize = student.allocatedLongPantsSize ? formatSize(cleanLongPantsSize) : '-';
        
        const isDebugMode = currentSizeDisplayMode === SIZE_DISPLAY_MODES.debug;
        const simplifiedFailureMessage = '分配失敗';
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>${student.class || ''}</td>
            <td>${String(student.number).padStart(2, '0') || ''}</td>
            <td>${student.name || ''}</td>
            <td>${student.gender || ''}</td>
            <td>${student.chest || ''}</td>
            <td>${student.waist || ''}</td>
            <td>${student.pantsLength || ''}</td>
            <td>
                ${formattedShirtSize}${shirtMark ? `<span class="allocation-mark">${shirtMark}</span>` : ''}
                ${shortShirtFailReason ? `<div class="failure-reason">${isDebugMode ? shortShirtFailReason : simplifiedFailureMessage}</div>` : ''}
                ${isDebugMode && student.isShirtSizeAdjustedForPantsLength ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
                ${isDebugMode && student.shortShirtLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥2</div>` : ''}
            </td>
            <td class="count-column">${student.shortSleeveShirtCount != null ? student.shortSleeveShirtCount : '-'}</td>
            <td>
                ${formattedPantsSize}${displayShortPantsMark ? `<span class="allocation-mark">${displayShortPantsMark}</span>` : ''} 
                ${shortPantsFailReason ? `<div class="failure-reason">${isDebugMode ? shortPantsFailReason : simplifiedFailureMessage}</div>` : ''}
                ${isDebugMode && student.isPantsLengthAdjusted ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
                ${isDebugMode && student.shortPantsLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥2</div>` : ''}
            </td>
            <td class="count-column">${student.shortSleevePantsCount != null ? student.shortSleevePantsCount : '-'}</td>
            <td>
                ${formattedLongShirtSize}${longShirtMark ? `<span class="allocation-mark">${longShirtMark}</span>` : ''}
                ${longShirtFailReason ? `<div class="failure-reason">${isDebugMode ? longShirtFailReason : simplifiedFailureMessage}</div>` : ''}
                ${isDebugMode && student.isLongShirtSizeAdjustedForPantsLength ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
                ${isDebugMode && student.longShirtLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥2</div>` : ''}
            </td>
            <td class="count-column">${student.longSleeveShirtCount != null ? student.longSleeveShirtCount : '-'}</td>
            <td>
                ${formattedLongPantsSize}${displayLongPantsMark ? `<span class="allocation-mark">${displayLongPantsMark}</span>` : ''} 
                ${longPantsFailReason ? `<div class="failure-reason">${isDebugMode ? longPantsFailReason : simplifiedFailureMessage}</div>` : ''}
                ${isDebugMode && student.isLongPantsSizeAdjustedForPantsLength ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
                ${isDebugMode && student.longPantsLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥3</div>` : ''}
            </td>
            <td class="count-column">${student.longSleevePantsCount != null ? student.longSleevePantsCount : '-'}</td>
        `;
        tbody.appendChild(row);
    });

    console.log('學生詳細分配結果表格已更新');
}

/**
 * 更新分配統計
 */
export function updateAllocationStats() {
    // 分配結果統計表格已被移除，此函數已禁用
    console.log('分配結果統計表格(#resultTable)已被移除，不再更新統計資料');
    return;
    
    // 以下代碼保留但不再執行
    // 尋找結果表格
    let resultTable = document.getElementById('resultTable');
    
    // 如果表格不存在（可能被註釋掉了），直接返回
    if (!resultTable) {
        console.log('未找到分配結果統計表格（#resultTable），可能已被註釋掉');
        return;
    }
    
    // 更新統計表格
    for (const type in allocationStats) {
        if (!allocationStats.hasOwnProperty(type)) continue;
        
        const stats = allocationStats[type];
        const displayName = UNIFORM_TYPES[type];
        
        let tbody = resultTable.querySelector('tbody');
        if (!tbody) continue;
        
        // 尋找現有行或創建新行
        let row = tbody.querySelector(`tr[data-type="${type}"]`);
        
        if (!row) {
            row = document.createElement('tr');
            row.setAttribute('data-type', type);
            tbody.appendChild(row);
        }
        
        // 計算分配率
        const totalDemand = demandData[type] ? demandData[type].totalDemand : 0;
        const allocatedPercent = totalDemand > 0 ? ((stats.allocated / totalDemand) * 100).toFixed(1) : '0.0';
        const exactPercent = totalDemand > 0 ? ((stats.exact / totalDemand) * 100).toFixed(1) : '0.0';
        const differentPercent = totalDemand > 0 ? ((stats.different / totalDemand) * 100).toFixed(1) : '0.0';
        const pantsSizeAdjustedPercent = totalDemand > 0 ? ((stats.pantsSizeAdjusted / totalDemand) * 100).toFixed(1) : '0.0';
        const failedPercent = totalDemand > 0 ? ((stats.failed / totalDemand) * 100).toFixed(1) : '0.0';
        
        // 更新行內容
        row.innerHTML = `
            <td>${displayName}</td>
            <td>${totalDemand}</td>
            <td>${stats.allocated} (${allocatedPercent}%)</td>
            <td>${stats.exact} (${exactPercent}%)</td>
            <td>${stats.different} (${differentPercent}%)</td>
            <td>${stats.pantsSizeAdjusted} (${pantsSizeAdjustedPercent}%)</td>
            <td>${stats.failed} (${failedPercent}%)</td>
        `;
    }
}

/**
 * 更新庫存分配結果
 */
export function updateInventoryAllocationResults() {
    console.log('執行updateInventoryAllocationResults');
    // 尺寸分配結果表格
    for (const type in inventoryData) {
        if (!inventoryData.hasOwnProperty(type)) continue;
        
        // 確定表格ID
        const tableId = `${type}ResultTable`;
        const table = document.getElementById(tableId);
        
        if (!table) {
            console.warn(`未找到表格 #${tableId}`);
            continue;
        }
        
        const tbody = table.querySelector('tbody');
        if (!tbody) {
            console.warn(`表格 #${tableId} 沒有tbody元素`);
            continue;
        }
        
        // 清空表格
        tbody.innerHTML = '';
        
        console.log(`處理 ${type} 庫存數據:`, inventoryData[type]);
        console.log(`最後分配狀態:`, lastAllocationStatus[type]);
        
        // 檢查是否為短袖上衣
        const isShortSleeveShirt = (type === 'shortSleeveShirt');
        
        if (isShortSleeveShirt) {
            console.log('%c===== 短袖上衣分配結果剩餘數邏輯詳細說明 =====', 'background: #3498db; color: white; font-size: 14px; padding: 5px;');
        }
        
        // 遍歷每個尺寸添加行
        for (const size of SIZES) {
            if (!inventoryData[type][size]) continue;
            
            const invData = inventoryData[type][size];
            const total = invData.total || 0;
            const reserved = invData.reserved || 0;
            const originalAllocatable = total - reserved; // 原始可分配數量
            
            // 優先使用記錄的最後分配狀態，如果存在
            let allocated = invData.allocated || 0; // 預設使用當前值
            let remaining = invData.allocatable || 0; // 預設使用當前值
            
            const lastStatus = lastAllocationStatus[type][size];
            if (lastStatus) {
                if (isShortSleeveShirt) {
                    console.log(`尺寸 ${size} - 使用記錄的最後分配狀態:`);
                    console.log(`- 記錄的已分配數: ${lastStatus.allocated}件`);
                    console.log(`- 記錄的剩餘數: ${lastStatus.remaining}件`);
                    console.log(`- 記錄時間: ${new Date(lastStatus.timestamp).toLocaleTimeString()}`);
                }
                
                allocated = lastStatus.allocated;
                remaining = lastStatus.remaining;
            } else if (isShortSleeveShirt) {
                console.log(`尺寸 ${size} - 未找到記錄的最後分配狀態，使用當前數據`);
            }
            
            // 添加更詳細的日誌記錄
            if (isShortSleeveShirt) {
                console.log(`%c尺寸 ${size} 詳細數據:`, 'color: #2980b9; font-weight: bold;');
                console.log(`- 總庫存(total): ${total}件`);
                console.log(`- 預留數(reserved): ${reserved}件`);
                console.log(`- 原始可分配數(originalAllocatable = total - reserved): ${originalAllocatable}件`);
                console.log(`- 系統記錄的已分配數(allocated): ${invData.allocated}件`);
                console.log(`- 系統記錄的剩餘數(allocatable): ${invData.allocatable}件`);
                console.log(`- 使用的已分配數(來自${lastStatus ? '記錄' : '系統'}): ${allocated}件`);
                console.log(`- 使用的剩餘數(來自${lastStatus ? '記錄' : '系統'}): ${remaining}件`);
            } else {
                console.log(`尺寸 ${size} 詳細數據:`);
                console.log(`- 總庫存(total): ${total}件`);
                console.log(`- 原始可分配數(originalAllocatable): ${originalAllocatable}件`);
                console.log(`- 已分配數(allocated): ${allocated}件`);
                console.log(`- 分配剩餘數(remaining/allocatable): ${remaining}件`);
                console.log(`- 預留數(reserved): ${reserved}件`);
            }
            
            // 驗證剩餘數計算是否正確
            const calculatedRemaining = Math.max(0, originalAllocatable - allocated);
            if (remaining !== calculatedRemaining && !lastStatus) {
                if (isShortSleeveShirt) {
                    console.warn(`%c警告: 尺寸 ${size} 的分配剩餘數不一致!`, 'color: #e74c3c; font-weight: bold;');
                    console.warn(`- 顯示值(系統記錄): ${remaining}件`);
                    console.warn(`- 計算值(原始可分配數 - 已分配數): ${calculatedRemaining}件`);
                    console.warn(`- 計算公式: ${originalAllocatable} - ${allocated} = ${calculatedRemaining}`);
                } else {
                    console.warn(`警告: 尺寸 ${size} 的分配剩餘數不一致! 顯示值=${remaining}, 計算值=${calculatedRemaining}`);
                }
                
                // 如果沒有記錄的最後狀態，則使用計算值
                if (!lastStatus) {
                    if (isShortSleeveShirt) {
                        console.log(`沒有找到記錄的最後狀態，使用計算值替換:`);
                        console.log(`- 更新前: ${remaining}件`);
                        console.log(`- 更新後: ${calculatedRemaining}件`);
                    } else {
                        console.log(`沒有找到記錄的最後狀態，使用計算值替換`);
                    }
                    
                    remaining = calculatedRemaining;
                    // 更新庫存數據，確保後續操作使用正確的值
                    invData.allocatable = calculatedRemaining;
                    
                    if (isShortSleeveShirt) {
                        console.log(`%c已更新: 尺寸 ${size} 的分配剩餘數現在為 ${remaining}件`, 'color: #27ae60; font-weight: bold;');
                    } else {
                        console.log(`已更新: 尺寸 ${size} 的分配剩餘數現在為 ${remaining}件`);
                    }
                } else {
                    if (isShortSleeveShirt) {
                        console.log(`%c使用記錄的最後狀態而非計算值:`, 'color: #f39c12; font-weight: bold;');
                        console.log(`- 記錄的已分配數: ${allocated}件`);
                        console.log(`- 記錄的剩餘數: ${remaining}件`);
                    } else {
                        console.log(`使用記錄的最後狀態: 已分配=${allocated}, 剩餘=${remaining}`);
                    }
                }
            } else if (isShortSleeveShirt) {
                console.log(`%c尺寸 ${size} 的分配剩餘數一致:`, 'color: #27ae60;');
                console.log(`- 顯示值: ${remaining}件`);
                console.log(`- 計算值: ${calculatedRemaining}件`);
                console.log(`- 計算公式: ${originalAllocatable} - ${allocated} = ${calculatedRemaining}`);
            }
            
            // 創建行（更改顯示順序：總庫存 可分配數 已分配 分配剩餘數 預留數）
            const row = document.createElement('tr');
            row.innerHTML = `
                <td data-original-size="${size}">${formatSize(size)}</td>
                <td>${total}</td>
                <td>${originalAllocatable}</td>
                <td><strong>${allocated}</strong></td>
                <td><strong>${remaining}</strong></td>
                <td>${reserved}</td>
            `;
            
            tbody.appendChild(row);
            
            if (isShortSleeveShirt) {
                console.log(`%c尺寸 ${size} 的最終顯示結果:`, 'color: #8e44ad; font-weight: bold;');
                console.log(`- 總庫存: ${total}件`);
                console.log(`- 可分配數: ${originalAllocatable}件`);
                console.log(`- 已分配: ${allocated}件`);
                console.log(`- 分配剩餘數: ${remaining}件`);
                console.log(`- 預留數: ${reserved}件`);
                console.log('-----------------------------------');
            }
        }
        
        if (isShortSleeveShirt) {
            console.log('%c===== 短袖上衣分配結果剩餘數邏輯說明結束 =====', 'background: #3498db; color: white; font-size: 14px; padding: 5px;');
        }
    }
    
    // 儲存更新後的庫存數據到本地存儲
    saveToLocalStorage('inventoryData', inventoryData);
    
    // 確保更新後的表格是可見的
    console.log('庫存分配結果表格已更新');
}

/**
 * 生成分配統計資料
 */
export function generateAllocationStats() {
    const stats = {
        totalStudents: studentData.length,
        allocationRates: {}
    };

    // 計算每種制服類型的分配比率
    for (const type in UNIFORM_TYPES) {
        if (!UNIFORM_TYPES.hasOwnProperty(type)) continue;
        
        const typeStats = allocationStats[type] || { allocated: 0, exact: 0, different: 0, failed: 0 };
        const totalDemand = demandData[type] ? demandData[type].totalDemand : 0;
        const totalInventory = calculateTotalInventory(inventoryData[type] || {});
        
        stats.allocationRates[type] = {
            type: UNIFORM_TYPES[type],
            demand: totalDemand,
            inventory: totalInventory,
            ratio: totalDemand > 0 && totalInventory > 0 ? (totalDemand / totalInventory) : 0,
            allocated: typeStats.allocated,
            exact: typeStats.exact,
            different: typeStats.different,
            failed: typeStats.failed,
            allocatedPercent: totalDemand > 0 ? (typeStats.allocated / totalDemand) : 0,
            exactPercent: totalDemand > 0 ? (typeStats.exact / totalDemand) : 0,
            differentPercent: totalDemand > 0 ? (typeStats.different / totalDemand) : 0,
            failedPercent: totalDemand > 0 ? (typeStats.failed / totalDemand) : 0
        };
    }
    
    return stats;
}

/**
 * 載入分配結果
 */
export function loadAllocationResults() {
    console.log('載入分配結果');
    
    // 先重置統計數據，確保正確初始化
    resetAllocationStats();
    
    // 載入分配統計
    const savedStats = loadFromLocalStorage('allocationStats', null);
    if (savedStats) {
        // 確保所有類型都有統計數據
        for (const type in allocationStats) {
            if (savedStats[type]) {
                allocationStats[type] = savedStats[type];
            } else {
                console.warn(`載入的分配統計數據中缺少 ${type} 類型的數據，將使用預設值`);
            }
        }
    }
    
    // 載入最後分配狀態
    const savedLastStatus = loadFromLocalStorage('lastAllocationStatus', null);
    if (savedLastStatus) {
        // 確保所有類型都有狀態數據
        for (const type in lastAllocationStatus) {
            if (savedLastStatus[type]) {
                lastAllocationStatus[type] = savedLastStatus[type];
            } else {
                console.warn(`載入的最後分配狀態中缺少 ${type} 類型的數據，將使用預設值`);
            }
        }
    }
    
    // 更新分配結果頁面
    updateAllocationResults();
    
    console.log('分配結果已載入並更新頁面');
}

/**
 * 獲取可用的尺寸列表並排序
 * @param {Object} inventoryData - 庫存數據
 * @param {string} inventoryType - 庫存類型
 * @returns {Array} - 排序後的可用尺寸
 */
function getAvailableSizes(inventoryData, inventoryType) {
    const availableSizes = [];
    let totalAvailable = 0;
    
    console.log(`獲取 ${UNIFORM_TYPES[inventoryType]} 可用尺寸列表:`);
    
    for (const size in inventoryData[inventoryType]) {
        if (!inventoryData[inventoryType].hasOwnProperty(size)) continue;
        
        const invData = inventoryData[inventoryType][size];
        const available = invData.allocatable || 0;
        const total = invData.total || 0;
        const reserved = invData.reserved || 0;
        
        // 記錄詳細的尺寸資訊
        console.log(`  尺寸 ${size}: 總庫存=${total}, 預留=${reserved}, 可分配=${available}`);
        
        if (available > 0) {
            availableSizes.push({
                size: size,
                available: available
            });
            totalAvailable += available;
        }
    }
    
    // 按照尺寸索引排序
    availableSizes.sort((a, b) => {
        return SIZES.indexOf(a.size) - SIZES.indexOf(b.size);
    });
    
    console.log(`${UNIFORM_TYPES[inventoryType]} 總可用尺寸數量: ${availableSizes.length}, 總可分配數量: ${totalAvailable}`);
    
    return availableSizes;
}

/**
 * 查找尺寸的索引
 * @param {string} size - 尺寸代碼
 * @returns {number} - 尺寸索引
 */
function getSizeIndex(size) {
    return SIZES.indexOf(size);
}

/**
 * 找到最合適的尺寸
 * @param {Array} suitableSizes - 合適的尺寸列表
 * @param {Object} inventory - 庫存數據
 * @returns {string} - 最合適的尺寸
 */
function findBestSize(suitableSizes, inventory) {
    let bestSize = null;
    let maxAvailable = 0;
    
    console.log(`尋找最佳尺寸，考慮 ${suitableSizes.length} 個候選尺寸:`);
    
    // 尋找庫存最多的尺寸
    for (const size of suitableSizes) {
        const available = inventory[size]?.allocatable || 0;
        console.log(`  尺寸 ${size}: 可分配數量=${available}`);
        
        if (available > maxAvailable) {
            maxAvailable = available;
            bestSize = size;
            console.log(`  -> 更新最佳尺寸為 ${size}，可分配數量=${available}`);
        }
    }
    
    console.log(`選定最佳尺寸: ${bestSize || '無可用尺寸'}${bestSize ? `，可分配數量=${inventory[bestSize]?.allocatable}` : ''}`);
    
    return bestSize;
}

/**
 * 重置分配統計數據
 */
export function resetAllocationStats() {
    // 重置分配統計
    allocationStats = {
        shortSleeveShirt: { allocated: 0, exact: 0, different: 0, failed: 0, special: 0 },
        shortSleevePants: { allocated: 0, exact: 0, different: 0, failed: 0, special: 0 },
        longSleeveShirt: { allocated: 0, exact: 0, different: 0, failed: 0, special: 0 },
        longSleevePants: { allocated: 0, exact: 0, different: 0, failed: 0, special: 0 }
    };
}

/**
 * 保存數據到本地存儲
 */
export function saveData() {
    // 保存學生數據
    saveToLocalStorage('studentData', studentData);
    
    // 保存庫存數據
    saveToLocalStorage('inventoryData', inventoryData);
    
    // 保存分配統計
    saveToLocalStorage('allocationStats', allocationStats);
}

/**
 * 嘗試分配特定尺寸
 * @param {Object} student - 學生資料
 * @param {string} size - 要分配的尺寸
 * @param {number} requiredCount - 需求件數
 * @param {Object} inventory - 庫存資料
 * @param {string} allocatedField - 分配結果欄位
 * @param {string} specialField - 特殊分配標記欄位
 * @param {string} inventoryType - 庫存類型
 * @returns {Object} - 分配結果和原因
 */
function tryAllocateSize(student, size, requiredCount, inventory, allocatedField, specialField, inventoryType) {
    // 檢查庫存是否足夠
    if (!inventory[size]) {
        return { success: false, reason: `尺寸 ${size} 沒有庫存` };
    }
    
    if (inventory[size].allocatable < requiredCount) {
        return { success: false, reason: `尺寸 ${size} 庫存不足，需要 ${requiredCount} 件，但只剩 ${inventory[size].allocatable} 件` };
    }

    // 分配並扣減庫存
    student[allocatedField] = size;
    student[specialField] = false;
    decreaseInventory(inventory, size, requiredCount, inventoryType);
    return { success: true }; 
}

/**
 * 將分配結果匯出到Excel
 */
function exportAllocationResultsToExcel() {
    try {
        // 分配結果統計表 (#resultTable) 已從界面中移除，不會匯出到Excel中
        
        // 記錄當前使用的尺寸顯示模式
        console.log(`匯出Excel檔案 - 使用尺寸顯示模式: ${currentSizeDisplayMode}`);
        
        // 創建一個新的工作簿
        const workbook = XLSX.utils.book_new();

        // 學生分配詳細結果工作表
        const studentWorksheet = createStudentDetailWorksheet();
        XLSX.utils.book_append_sheet(workbook, studentWorksheet, '學生分配詳細結果');

        // 各制服類型分配結果工作表
        const typeWorksheets = [
            { id: 'shortSleeveShirtResultTable', name: '短衣分配結果' },
            { id: 'shortSleevePantsResultTable', name: '短褲分配結果' },
            { id: 'longSleeveShirtResultTable', name: '長衣分配結果' },
            { id: 'longSleevePantsResultTable', name: '長褲分配結果' }
        ];

        typeWorksheets.forEach(({ id, name }) => {
            const worksheet = createUniformTypeWorksheet(id);
            if (worksheet) {
                XLSX.utils.book_append_sheet(workbook, worksheet, name);
            }
        });

        // 生成文件名稱
        const now = new Date();
        const timestamp = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
        
        // 將尺寸顯示模式加入檔名
        let displayModeText = '';
        switch (currentSizeDisplayMode) {
            case SIZE_DISPLAY_MODES.size:
                displayModeText = '尺寸';
                break;
            case SIZE_DISPLAY_MODES.number:
                displayModeText = '尺碼';
                break;
            case SIZE_DISPLAY_MODES.debug:
                displayModeText = 'Debug';
                break;
            case SIZE_DISPLAY_MODES.both:
            default:
                displayModeText = '完整';
                break;
        }
        
        const filename = `制服分配結果_${displayModeText}_${timestamp}.xlsx`;

        // 下載Excel文件
        XLSX.writeFile(workbook, filename);
        showAlert('分配結果已成功匯出為Excel檔案', 'success');
    } catch (error) {
        console.error('匯出分配結果時發生錯誤:', error);
        showAlert(`匯出失敗: ${error.message}`, 'error');
    }
}

/**
 * 創建學生分配詳細結果工作表
 * @returns {Object} XLSX工作表對象
 */
function createStudentDetailWorksheet() {
    // 按照班級和座號排序學生數據
    const sortedStudents = [...studentData].sort((a, b) => {
        const classA = parseInt(a.class) || 0;
        const classB = parseInt(b.class) || 0;
        if (classA !== classB) return classA - classB;
        
        const numberA = parseInt(a.number) || 0;
        const numberB = parseInt(b.number) || 0;
        return numberA - numberB;
    });

    // 準備工作表數據
    const data = [];
    
    // 添加標題行
    data.push([
        '序號', '班級', '號碼', '姓名', '性別', '胸圍', '腰圍', '褲長',
        '短衣尺寸', '短衣件數', '短褲尺寸', '短褲件數',
        '長衣尺寸', '長衣件數', '長褲尺寸', '長褲件數'
    ]);

    // 判斷是否為Debug模式
    const isDebugMode = currentSizeDisplayMode === SIZE_DISPLAY_MODES.debug;
    // 簡化的分配失敗信息
    const simplifiedFailureMessage = '分配失敗';

    // 添加學生數據
    sortedStudents.forEach((student, index) => {
        // 處理各種制服的分配情況 - 使用formatSize函數格式化尺寸
        let shortShirtSize = student.allocatedShirtSize ? formatSize(student.allocatedShirtSize) : '-';
        // 如果有褲長調整，添加標記 (只在Debug模式下顯示詳細原因)
        if (student.isShirtSizeAdjustedForPantsLength && shortShirtSize !== '-') {
            shortShirtSize += '↑' + (isDebugMode ? '(褲長調整)' : '');
        }
        // 如果短衣褲長仍不足，添加標記 (只在Debug模式下顯示)
        if (isDebugMode && student.shortShirtLengthDeficiency && shortShirtSize !== '-') {
            shortShirtSize += '!(褲長仍不足≥2)';
        }
        
        // 修改短褲尺寸處理邏輯
        let shortPantsSize;
        if (student.allocatedPantsSize) {
            shortPantsSize = formatSize(student.allocatedPantsSize); // 格式化後的基本尺寸, e.g., "34"
            const originalAllocatedPantsValue = student.allocatedPantsSize; // 原始分配值, e.g., "XS/34*"

            // 如果原始分配值以 '*' 結尾 (補救標記)
            if (originalAllocatedPantsValue.endsWith('*')) {
                // 確保 '*' 被加到格式化後的尺寸上 (如果 formatSize 移除了它)
                if (shortPantsSize !== '-' && !shortPantsSize.endsWith('*')) {
                    shortPantsSize += '*';
                }
            }
            // 否則，如果不是補救分配，檢查是否有 pantsAdjustmentMark (如 '↑')
            else if (student.pantsAdjustmentMark) {
                shortPantsSize += student.pantsAdjustmentMark;
            }
        } else {
            shortPantsSize = '-';
        }

        // 如果短褲褲長仍不足，添加標記 (只在Debug模式下顯示)
        if (isDebugMode && student.shortPantsLengthDeficiency && shortPantsSize !== '-') {
            // 僅在 student.allocatedPantsSize 存在 (即有分配嘗試) 時附加
            if (student.allocatedPantsSize) {
            shortPantsSize += '!(褲長仍不足≥2)';
            }
        }
        
        let longShirtSize = student.allocatedLongShirtSize ? formatSize(student.allocatedLongShirtSize) : '-';
        // 如果長衣有褲長調整，添加標記 (只在Debug模式下顯示詳細原因)
        if (student.isLongShirtSizeAdjustedForPantsLength && longShirtSize !== '-') {
            longShirtSize += '↑' + (isDebugMode ? '(褲長調整)' : '');
        }
        // 如果長衣褲長仍不足，添加標記 (只在Debug模式下顯示)
        if (isDebugMode && student.longShirtLengthDeficiency && longShirtSize !== '-') {
            longShirtSize += '!(褲長仍不足≥2)';
        }
        
        // 修改長褲尺寸處理邏輯
        let longPantsSize;
        if (student.allocatedLongPantsSize) {
            longPantsSize = formatSize(student.allocatedLongPantsSize); // 格式化後的基本尺寸
            const originalAllocatedLongPantsValue = student.allocatedLongPantsSize; // 原始分配值

            // 如果原始分配值以 '*' 結尾 (補救標記)
            if (originalAllocatedLongPantsValue.endsWith('*')) {
                // 確保 '*' 被加到格式化後的尺寸上
                if (longPantsSize !== '-' && !longPantsSize.endsWith('*')) {
                    longPantsSize += '*';
                }
            }
            // 否則，如果不是補救分配，檢查是否有 longPantsAdjustmentMark
            else if (student.longPantsAdjustmentMark) {
                longPantsSize += student.longPantsAdjustmentMark;
            }
        } else {
            longPantsSize = '-';
        }

        // 如果長褲褲長仍不足，添加標記 (只在Debug模式下顯示)
        if (isDebugMode && student.longPantsLengthDeficiency && longPantsSize !== '-') {
            // 僅在 student.allocatedLongPantsSize 存在 (即有分配嘗試) 時附加
            if (student.allocatedLongPantsSize) {
            longPantsSize += '!(褲長仍不足≥3)';
            }
        }
        
        // 如果有分配失敗原因，根據顯示模式決定顯示內容
        if (student.allocationFailReason) {
            if (student.allocationFailReason.shortSleeveShirt && !student.allocatedShirtSize) {
                shortShirtSize = isDebugMode ? student.allocationFailReason.shortSleeveShirt : simplifiedFailureMessage;
            }
            if (student.allocationFailReason.shortSleevePants && !student.allocatedPantsSize) {
                shortPantsSize = isDebugMode ? student.allocationFailReason.shortSleevePants : simplifiedFailureMessage;
            }
            if (student.allocationFailReason.longSleeveShirt && !student.allocatedLongShirtSize) {
                longShirtSize = isDebugMode ? student.allocationFailReason.longSleeveShirt : simplifiedFailureMessage;
            }
            if (student.allocationFailReason.longSleevePants && !student.allocatedLongPantsSize) {
                longPantsSize = isDebugMode ? student.allocationFailReason.longSleevePants : simplifiedFailureMessage;
            }
        }

        // 添加學生行數據
        data.push([
            index + 1,
            student.class || '',
            String(student.number).padStart(2, '0') || '',
            student.name || '',
            student.gender || '',
            student.chest || '',
            student.waist || '',
            student.pantsLength || '',
            shortShirtSize,
            student.allocatedShirtSize ? (student.shortSleeveShirtCount || 1) : '-',
            shortPantsSize,
            student.allocatedPantsSize ? (student.shortSleevePantsCount || 1) : '-',
            longShirtSize,
            student.allocatedLongShirtSize ? (student.longSleeveShirtCount || 1) : '-',
            longPantsSize,
            student.allocatedLongPantsSize ? (student.longSleevePantsCount || 1) : '-'
        ]);
    });

    // 創建工作表
    return XLSX.utils.aoa_to_sheet(data);
}

/**
 * 創建制服類型分配結果工作表
 * @param {string} tableId - 表格ID
 * @returns {Object|null} XLSX工作表對象，如果表格不存在則返回null
 */
function createUniformTypeWorksheet(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return null;

    // 獲取表格標題
    const headers = [];
    const headerRow = table.querySelector('thead tr');
    if (headerRow) {
        const headerCells = headerRow.querySelectorAll('th');
        headerCells.forEach(cell => {
            headers.push(cell.textContent.trim());
        });
    }

    // 獲取表格數據
    const data = [headers]; // 第一行是標題
    const rows = table.querySelectorAll('tbody tr');
    rows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('td');
        cells.forEach((cell, index) => {
            // 檢查是否是尺寸欄位（第一列）
            if (index === 0) {
                // 獲取原始尺寸
                const originalSize = cell.getAttribute('data-original-size');
                // 如果有原始尺寸，使用formatSize函數格式化
                if (originalSize) {
                    rowData.push(formatSize(originalSize));
                } else {
                    rowData.push(cell.textContent.trim());
                }
            } else {
                rowData.push(cell.textContent.trim());
            }
        });
        data.push(rowData);
    });

    // 創建工作表
    return XLSX.utils.aoa_to_sheet(data);
}

/**
 * 從尺碼標記中獲取數字部分
 * @param {string} size - 尺寸代碼，如 'XS/34'
 * @returns {number} - 尺碼數字部分，如 34
 */
function getSizeNumber(size) {
    if (!size) return 0;
    
    // 嘗試分割尺寸代碼
    const parts = size.split('/');
    if (parts.length === 2) {
        return parseInt(parts[1], 10);
    }
    
    // 嘗試從尺碼中提取數字部分
    const match = size.match(/(\d+)/);
    if (match) {
        return parseInt(match[0], 10);
    }
    
    return 0;
}

/**
 * 檢查學生褲長是否需要調整上衣尺寸
 * @param {number} pantsLength - 學生褲長
 * @param {string} shirtSize - 原始上衣尺寸
 * @returns {boolean} - 是否需要增加尺寸
 */
function shouldAdjustShirtSizeForLongPants(pantsLength, shirtSize) {
    const sizeNumber = getSizeNumber(shirtSize);
    // 檢查褲長與尺碼差異是否大於等於3
    return (pantsLength - sizeNumber >= 3);
}

/**
 * 獲取下一個更大的尺寸
 * @param {string} size - 當前尺寸
 * @returns {string} - 下一個更大的尺寸，如果已經是最大尺寸則返回原尺寸
 */
function getNextLargerSize(size) {
    const sizeIndex = SIZES.indexOf(size);
    if (sizeIndex < 0) return size; // 尺寸不在列表中
    
    // 如果已經是最大尺寸，則返回原尺寸
    if (sizeIndex >= SIZES.length - 1) return size;
    
    // 返回下一個更大的尺寸
    return SIZES[sizeIndex + 1];
}

/**
 * 獲取前一個更小的尺寸
 * @param {string} size - 當前尺寸
 * @returns {string} - 前一個更小的尺寸，如果已經是最小尺寸則返回原尺寸
 */
function getPreviousSmallerSize(size) {
    const sizeIndex = SIZES.indexOf(size);
    if (sizeIndex <= 0) return size; // 尺寸不在列表中
    
    // 如果已經是最小尺寸，則返回原尺寸
    if (sizeIndex === 0) return size;
    
    // 返回前一個更小的尺寸
    return SIZES[sizeIndex - 1];
}

/**
 * 短褲和長褲專用排序函數 - 按照新邏輯排序學生
 * @param {Array} students - 學生資料
 * @returns {Array} - 排序後的學生列表
 */
function sortStudentsForPants(students) {
    return [...students].sort((a, b) => {
        // 主要按腰圍從小到大排序
        if (a.waist !== b.waist) {
            return a.waist - b.waist;
        }
        // 腰圍相同時，按胸圍從小到大排序
        return a.chest - b.chest;
    });
}

/**
 * 根據尺碼獲取對應的長度數值
 * @param {string} size - 尺碼標識，如 "XS/34", "S/36" 等
 * @returns {number} - 尺碼對應的長度數值，如 34, 36 等
 */
function getLengthValueFromSize(size) {
    // 從尺碼字符串中提取數字部分
    const match = size.match(/\/(\d+)/);
    if (match && match[1]) {
        return parseInt(match[1], 10);
    }
    // 如果無法提取，返回 0 作為默認值
    console.warn(`無法從尺碼 ${size} 提取長度數值`);
    return 0;
}

/**
 * 全新的褲子分配邏輯 - 取代原有的 allocatePantsUnified 函數
 * @param {Array} students - 學生資料列表
 * @param {string} inventoryType - 庫存類型標識 (shortSleevePants 或 longSleevePants)
 * @param {string} allocatedField - 學生對象中存儲分配結果的屬性名
 * @param {string} adjustmentMarkField - 學生對象中存儲調整標記的屬性名
 * @param {string} studentPantsCountField - 學生對象中存儲需求數量的屬性名
 * @param {Object} pantsInventoryData - 特定褲子類型的庫存數據
 */
function allocatePantsNewLogic(students, inventoryType, allocatedField, adjustmentMarkField, studentPantsCountField, pantsInventoryData) {
    return new Promise((resolve) => {
        console.log(`%c===== 開始分配 ${UNIFORM_TYPES[inventoryType]} (新邏輯) =====`, 'background: #f39c12; color: white; font-size: 14px; padding: 5px;');

        // 檢查庫存是否存在
        if (!pantsInventoryData) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            students.forEach(student => {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '無庫存資料';
                student[allocatedField] = '';
                student[adjustmentMarkField] = null;
            });
            
            // 更新統計
            let failedCount = 0;
            students.forEach(student => {
                if ((student[studentPantsCountField] || 0) > 0) {
                    failedCount++;
                }
            });
            
            allocationStats[inventoryType] = {
                allocated: 0,
                failed: failedCount,
                exact: 0,
                different: 0,
                special: 0,
                pantsSizeAdjusted: 0 
            };
            
            resolve(false);
            return;
        }

        // 計算總需求和庫存
        let totalDemand = 0;
        students.forEach(student => {
            totalDemand += student[studentPantsCountField] || 1;
        });

        let totalAllocatable = 0;
        for (const size in pantsInventoryData) {
            totalAllocatable += pantsInventoryData[size]?.allocatable || 0;
        }

        console.log(`%c${UNIFORM_TYPES[inventoryType]}需求與庫存概況：`, 'color: #d35400; font-weight: bold;');
        console.log(`- 總需求量：${totalDemand}件`);
        console.log(`- 總可分配數量：${totalAllocatable}件`);

        if (totalAllocatable < totalDemand) {
            console.warn(`%c警告：${UNIFORM_TYPES[inventoryType]}可分配數總和(${totalAllocatable})小於需求量(${totalDemand})，差額${totalDemand - totalAllocatable}件`, 'color: #e74c3c; font-weight: bold;');
        }

        // 初始化統計
        const stats = { 
            allocated: 0, 
            failed: 0, 
            pantsSizeAdjusted: 0, 
            fallbackUsed: 0 
        };

        // 按新邏輯排序學生
        const sortedStudentsForPants = sortStudentsForPants(students);
        
        // 獲取所有可用尺碼，按照尺碼大小排序
        const availableSizes = [];
        let totalAvailable = 0;
        
        console.log(`獲取 ${UNIFORM_TYPES[inventoryType]} 可用尺寸列表:`);
        
        for (const size in pantsInventoryData) {
            if (!pantsInventoryData.hasOwnProperty(size)) continue;
            
            const invData = pantsInventoryData[size];
            const available = invData.allocatable || 0;
            const total = invData.total || 0;
            const reserved = invData.reserved || 0;
            
            // 記錄詳細的尺寸資訊
            console.log(`  尺寸 ${size}: 總庫存=${total}, 預留=${reserved}, 可分配=${available}`);
            
            if (available > 0) {
                availableSizes.push({
                    size: size,
                    available: available
                });
                totalAvailable += available;
            }
        }
        
        // 按照尺寸索引排序
        availableSizes.sort((a, b) => {
            return SIZES.indexOf(a.size) - SIZES.indexOf(b.size);
        });
        
        console.log(`${UNIFORM_TYPES[inventoryType]} 總可用尺寸數量: ${availableSizes.length}, 總可分配數量: ${totalAvailable}`);
        console.log(`可用尺碼列表: ${availableSizes.map(s => s.size).join(', ')}`);

        // 對每個學生進行分配
        for (const student of sortedStudentsForPants) {
            // 清除之前的分配結果和標記
            student[allocatedField] = '';
            student[adjustmentMarkField] = null;
            if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                delete student.allocationFailReason[inventoryType];
            }

            const requiredCount = student[studentPantsCountField] || 0;
            if (requiredCount <= 0) {
                // 不需要分配
                continue;
            }

            console.log(`%c處理 ${UNIFORM_TYPES[inventoryType]} 分配給學生：${student.name} (班級：${student.class}，座號：${student.number})`, 'color: #2c3e50; font-weight: bold;');
            console.log(`- 腰圍：${student.waist}，褲長：${student.pantsLength}，性別：${student.gender}，需求：${requiredCount}件`);

            // 尋找適合的尺碼 - 從可用尺碼中找到第一個腰圍足夠大的
            let targetSize = null;
            let targetSizeIndex = -1;

            for (let i = 0; i < availableSizes.length; i++) {
                const size = availableSizes[i].size;
                const sizeLength = getLengthValueFromSize(size);
                
                // 簡單規則：尺碼數字 >= 腰圍*1.2 即視為合適
                // 這裡可以根據實際需求調整判定標準
                if (sizeLength >= student.waist * 1.2) {
                    targetSize = size;
                    targetSizeIndex = i;
                    break;
                }
            }

            if (!targetSize) {
                // 沒有找到合適尺碼
                console.warn(`學生 ${student.name} 腰圍 ${student.waist} 找不到合適尺碼，嘗試分配最大尺碼`);
                if (availableSizes.length > 0) {
                    targetSize = availableSizes[availableSizes.length - 1].size;
                    targetSizeIndex = availableSizes.length - 1;
                } else {
                    // 無可用尺碼
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：無可用尺碼';
                    stats.failed++;
                    continue;
                }
            }

            // 檢查庫存是否足夠
            if (pantsInventoryData[targetSize]?.allocatable < requiredCount) {
                // 尋找其他可用尺碼
                let foundAlternative = false;
                // 先嘗試大一號尺碼
                for (let i = targetSizeIndex + 1; i < availableSizes.length; i++) {
                    const altSize = availableSizes[i].size;
                    if (pantsInventoryData[altSize]?.allocatable >= requiredCount) {
                        targetSize = altSize;
                        foundAlternative = true;
                        break;
                    }
                }
                
                // 如果沒找到大一號尺碼，嘗試小一號尺碼
                if (!foundAlternative) {
                    for (let i = targetSizeIndex - 1; i >= 0; i--) {
                        const altSize = availableSizes[i].size;
                        if (pantsInventoryData[altSize]?.allocatable >= requiredCount) {
                            targetSize = altSize;
                            foundAlternative = true;
                            break;
                        }
                    }
                }
                
                if (!foundAlternative) {
                    // 沒有足夠庫存的尺碼
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：庫存不足';
                    stats.failed++;
                    continue;
                }
            }

            console.log(`學生 ${student.name} 初步分配尺碼: ${targetSize}`);
            
            // 褲長監聽器 - 檢查褲長是否需要調整尺碼
            const targetSizeLength = getLengthValueFromSize(targetSize);
            let finalSize = targetSize;
            let adjustmentMark = null;
            let fallbackMark = null;
            
            // 當學生褲長比分配到的尺碼數值 >= 2 時，提升一個尺碼
            if (student.pantsLength - targetSizeLength >= 2) {
                console.log(`學生 ${student.name} 褲長(${student.pantsLength})比分配尺碼長度(${targetSizeLength})大於等於2，嘗試提升尺碼`);
                
                // 從SIZES數組中獲取當前尺碼的索引位置
                const currentSizeIndex = SIZES.indexOf(targetSize);
                
                // 確保是有效的索引，並且不是最後一個尺碼
                if (currentSizeIndex !== -1 && currentSizeIndex < SIZES.length - 1) {
                    // 獲取下一個尺碼
                    const nextSize = SIZES[currentSizeIndex + 1];
                    const nextSizeLength = getLengthValueFromSize(nextSize);
                    
                    // 檢查大一號尺碼庫存
                    if (pantsInventoryData[nextSize]?.allocatable >= requiredCount) {
                        finalSize = nextSize;
                        adjustmentMark = '↑';
                        console.log(`成功提升尺碼: ${targetSize}(${targetSizeLength}) -> ${finalSize}(${nextSizeLength})${adjustmentMark}`);
                        stats.pantsSizeAdjusted++;
                    } else {
                        // 大一號尺碼庫存不足，回退到原尺碼並標記
                        console.log(`大一號尺碼 ${nextSize}(${nextSizeLength}) 庫存不足，回退到原尺碼 ${targetSize}(${targetSizeLength}) 並標記 *`);
                        fallbackMark = '*';
                        stats.fallbackUsed++;
                    }
                } else {
                    // 沒有更大的尺碼可用
                    console.log(`沒有更大的尺碼可用，維持原尺碼 ${targetSize}(${targetSizeLength}) 並標記 *`);
                    fallbackMark = '*';
                    stats.fallbackUsed++;
                }
            }
            
            // 分配最終尺碼並減少庫存
            if (decreaseInventory(pantsInventoryData, finalSize, requiredCount, inventoryType)) {
                student[allocatedField] = finalSize;
                if (adjustmentMark) {
                    student[adjustmentMarkField] = adjustmentMark;
                } else if (fallbackMark) {
                    student[allocatedField] = targetSize + fallbackMark;
                }
                
                stats.allocated++;
                // 更新日誌格式，確保顯示正確的分配結果
                const displaySize = student[allocatedField];
                const displayMark = student[adjustmentMarkField] || '';
                console.log(`%c${UNIFORM_TYPES[inventoryType]} 分配成功：${student.name} => ${displaySize}${displayMark} (需求 ${requiredCount}件)`, 'color: #27ae60; font-weight: bold;');
            } else {
                // 扣減庫存失敗
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = `分配失敗：扣減庫存時發生錯誤`;
                stats.failed++;
                console.error(`%c庫存扣減異常：${UNIFORM_TYPES[inventoryType]} 分配給 ${student.name} 尺碼 ${finalSize}，但 decreaseInventory 返回 false`, 'color: red; font-weight: bold;');
            }
        }

        // 更新庫存數據
        inventoryData[inventoryType] = pantsInventoryData;
        
        // 更新總體統計
        allocationStats[inventoryType] = {
            allocated: stats.allocated,
            failed: stats.failed,
            pantsSizeAdjusted: stats.pantsSizeAdjusted,
            fallbackUsed: stats.fallbackUsed,
            exact: 0,
            different: 0,
            special: 0 
        };

        console.log(`%c${UNIFORM_TYPES[inventoryType]}分配結果統計：`, 'background: #f39c12; color: white; font-size: 12px; padding: 5px;');
        console.log(`- 成功分配：${stats.allocated}人`);
        console.log(`- 分配失敗：${stats.failed}人`);
        console.log(`- 因褲長調整次數：${stats.pantsSizeAdjusted}`);
        console.log(`- 因調整失敗而標記 * 次數：${stats.fallbackUsed}`);
        console.log(`%c===== ${UNIFORM_TYPES[inventoryType]}分配完成 =====`, 'background: #f39c12; color: white; font-size: 14px; padding: 5px;');
        
        resolve(true);
    });
}