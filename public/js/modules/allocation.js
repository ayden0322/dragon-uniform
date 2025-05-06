// 分配相關功能模組
import { saveToLocalStorage, loadFromLocalStorage, showAlert, downloadExcel } from './utils.js';
import { SIZES, UNIFORM_TYPES, formatSize, currentSizeDisplayMode, SIZE_DISPLAY_MODES, getFemaleChestAdjustment, getCurrentSchoolConfig } from './config.js';
import { inventoryData, calculateTotalInventory, updateInventoryUI, manualAdjustments, initInventoryFeatures, saveManualAdjustments, saveManualAdjustmentsSilent } from './inventory.js';
import { studentData, sortedStudentData, demandData, updateStudentAllocationUI, updateAdjustmentPage } from './students.js';
import { updateAllocationRatios, formatSizeWithAdjustment } from './ui.js';

// 記錄最後分配狀態
const lastAllocationStatus = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

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
    // 確保使用正確的庫存類型
    console.log('開始分配短褲');
    return allocateShortPants('shortSleevePants', 'allocatedPantsSize', 'isSpecialPantsAllocation');
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
    // 確保使用正確的庫存類型
    console.log('開始分配長褲');
    return allocateLongPants('longSleevePants', 'allocatedLongPantsSize', 'isSpecialLongPantsAllocation');
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
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            
            // 將所有學生標記為分配失敗（無庫存）
            _localSortedStudentData.forEach(student => {
                if (!student[allocatedField]) {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '無庫存資料';
                }
            });
            
            allocationStats[inventoryType].failed = _localSortedStudentData.length;
            resolve(false);
            return;
        }
        
        // 計算總需求量
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.shortSleeveShirtCount || 1;
        });
        
        // 計算總可分配數量
        let totalAllocatable = 0;
        for (const size in inventoryData[inventoryType]) {
            totalAllocatable += inventoryData[inventoryType][size].allocatable || 0;
        }
        
        console.log(`%c${UNIFORM_TYPES[inventoryType]}需求與庫存概況：`, 'color: #2980b9; font-weight: bold;');
        console.log(`- 總需求量：${totalDemand}件`);
        console.log(`- 總可分配數量：${totalAllocatable}件`);
        console.log(`- 學生數量：${_localSortedStudentData.length}人`);
        
        // 如果總可分配數小於總需求量，發出警告
        if (totalAllocatable < totalDemand) {
            console.warn(`%c警告：${UNIFORM_TYPES[inventoryType]}可分配數總和(${totalAllocatable})小於需求量(${totalDemand})，差額${totalDemand - totalAllocatable}件`, 'color: #e74c3c; font-weight: bold;');
        }
        
        // 獲取可用尺寸並排序（小到大）
        let availableSizes = getAvailableSizes(inventoryData, inventoryType);
        
        // 過濾掉可分配數為0的尺寸
        availableSizes = availableSizes.filter(size => size.available > 0);
        
        console.log(`%c可用尺寸列表：`, 'color: #2980b9; font-weight: bold;');
        availableSizes.forEach(s => {
            console.log(`- ${s.size}: ${s.available}件可分配`);
        });
        
        if (availableSizes.length === 0) {
            console.warn(`%c沒有可用的 ${UNIFORM_TYPES[inventoryType]} 庫存，無法進行分配`, 'color: #e74c3c; font-weight: bold;');
            
            // 將所有學生標記為分配失敗（尺寸無庫存）
            _localSortedStudentData.forEach(student => {
                if (!student[allocatedField]) {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '所有尺寸無庫存';
                }
            });
            
            allocationStats[inventoryType].failed = _localSortedStudentData.length;
            resolve(false);
            return;
        }
        
        // 複製庫存資料用於分配
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        // 複製學生數組並根據有效胸圍排序
        const sortedStudents = [..._localSortedStudentData].map(student => {
            // 計算有效胸圍（胸圍和腰圍的較大值）
            const chest = student.chest || 0;
            const waist = student.waist || 0;
            const effectiveChest = Math.max(chest, waist);
            return { student, effectiveChest };
        }).sort((a, b) => a.effectiveChest - b.effectiveChest);

        console.log(`%c學生排序結果（按有效胸圍排序）：`, 'color: #2980b9; font-weight: bold;');
        console.log(`學生總數：${sortedStudents.length}人`);
        console.log(`前5名學生：`);
        sortedStudents.slice(0, 5).forEach((s, i) => {
            console.log(`${i+1}. ${s.student.name}(${s.student.gender}，班級：${s.student.class}，座號：${s.student.number})：` +
                     `胸圍=${s.student.chest}，腰圍=${s.student.waist}，有效胸圍=${s.effectiveChest}`);
        });
        
        // 用於記錄分配結果的統計資料
        const stats = {
            allocated: 0,
            exact: 0,
            different: 0,
            failed: 0,
            special: 0,
            pantsSizeAdjusted: 0 // 新增：因褲長調整尺寸的學生數
        };
        
        // 記錄未分配的學生
        const unallocatedStudents = [];
        
        // 第一階段：按照排序分配
        console.log(`%c第一階段分配開始：按胸圍順序分配`, 'background: #27ae60; color: white; font-size: 12px; padding: 3px;');
        
        // 添加檢查已存在的失敗原因
        const studentsWithPriorFailures = sortedStudents
            .filter(({student}) => student.allocationFailReason && Object.keys(student.allocationFailReason).length > 0);
        
        if (studentsWithPriorFailures.length > 0) {
            console.log(`%c注意：檢測到 ${studentsWithPriorFailures.length} 名學生在之前的步驟已有失敗記錄：`, 'background: #f39c12; color: white; font-size: 12px; padding: 3px;');
            studentsWithPriorFailures.forEach(({student}) => {
                const reasons = Object.entries(student.allocationFailReason)
                    .map(([type, reason]) => `${UNIFORM_TYPES[type]}: ${reason}`)
                    .join(', ');
                console.log(`學生 ${student.name}(${student.class}-${student.number}) 之前失敗原因: ${reasons}`);
            });
        }
        
        for (const {student, effectiveChest} of sortedStudents) {
            // 跳過已分配的學生
            if (student[allocatedField]) {
                console.log(`學生 ${student.name}(${student.class}-${student.number}) 已有分配結果：${student[allocatedField]}，跳過`);
                continue;
            }
            
            // 檢查學生是否需要此類型的制服
            const requiredCount = student.shortSleeveShirtCount || 1;
            
            // 詳細記錄學生資訊
            console.log(`%c處理學生：${student.name}(${student.gender}，班級：${student.class}，座號：${student.number})`, 'color: #8e44ad; font-weight: bold;');
            console.log(`- 胸圍：${student.chest}`);
            console.log(`- 腰圍：${student.waist}`);
            console.log(`- 褲長：${student.pantsLength}`);
            console.log(`- 有效胸圍：${effectiveChest}`);
            console.log(`- 需求件數：${requiredCount}`);
            
            // 如果學生不需要此類型制服（件數為0），標記為不需要
            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生不需要此制服（需求件數為0），跳過分配`);
                continue;
            }
            
            // 新增：計算理論上的尺碼數值（根據有效胸圍和性別）
            let adjustment = 0;
            if (student.gender === '男') {
                adjustment = (effectiveChest % 2 === 0) ? 10 : 11;
            } else {
                adjustment = (effectiveChest % 2 === 0) ? 8 : 9;
            }
            const targetSizeNumber = effectiveChest + adjustment;
            
            console.log(`%c計算理論尺碼：`, 'color: #2c3e50;');
            console.log(`- 有效胸圍：${effectiveChest}`);
            console.log(`- 性別：${student.gender}`);
            console.log(`- 性別調整值：${adjustment}`);
            console.log(`- 計算結果：${effectiveChest} + ${adjustment} = ${targetSizeNumber}`);
            
            // 找出理論上的尺碼代碼
            let targetSize = '';
            if (targetSizeNumber <= 34) targetSize = 'XS/34';
            else if (targetSizeNumber <= 36) targetSize = 'S/36';
            else if (targetSizeNumber <= 38) targetSize = 'M/38';
            else if (targetSizeNumber <= 40) targetSize = 'L/40';
            else if (targetSizeNumber <= 42) targetSize = 'XL/42';
            else if (targetSizeNumber <= 44) targetSize = '2L/44';
            else if (targetSizeNumber <= 46) targetSize = '3L/46';
            else if (targetSizeNumber <= 48) targetSize = '4L/48';
            else if (targetSizeNumber <= 50) targetSize = '5L/50';
            else if (targetSizeNumber <= 52) targetSize = '6L/52';
            else if (targetSizeNumber <= 54) targetSize = '7L/54';
            else if (targetSizeNumber <= 56) targetSize = '8L/56';
            else targetSize = '9L/58';
            
            console.log(`- 對應尺碼：${targetSize}`);
            
            let allocated = false;
            
            // 處理邊界情況：計算結果低於最小尺碼或超過最大尺碼
            const availableSizesArray = availableSizes.map(s => s.size);
            const minAvailableSize = availableSizesArray[0];
            const maxAvailableSize = availableSizesArray[availableSizesArray.length - 1];
            const minSizeNumber = getSizeNumber(minAvailableSize);
            const maxSizeNumber = getSizeNumber(maxAvailableSize);
            
            console.log(`%c檢查尺碼邊界情況：`, 'color: #2c3e50;');
            console.log(`- 最小可用尺碼：${minAvailableSize}(${minSizeNumber})`);
            console.log(`- 最大可用尺碼：${maxAvailableSize}(${maxSizeNumber})`);
            console.log(`- 計算的目標尺碼：${targetSize}(${getSizeNumber(targetSize)})`);
            
            // 檢查計算的尺碼是否在可用範圍內
            let finalSize = targetSize;
            let isSpecialAllocation = false;
            
            // 如果計算結果低於最小可用尺碼
            if (targetSizeNumber < minSizeNumber) {
                console.log(`%c計算結果 ${targetSize}(${targetSizeNumber}) 小於最小可用尺碼 ${minAvailableSize}(${minSizeNumber})`, 'color: #e74c3c;');
                
                // 如果庫存有最小尺碼，則分配最小尺碼
                if (remainingInventory[minAvailableSize]?.allocatable >= requiredCount) {
                    finalSize = minAvailableSize;
                    isSpecialAllocation = true;
                    console.log(`%c使用最小可用尺碼 ${minAvailableSize} 作為替代`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：計算尺碼過小且無可用替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 的計算尺碼過小且無可用替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    unallocatedStudents.push({ student, requiredCount });
                    continue;
                }
            }
            // 如果計算結果超過最大可用尺碼
            else if (targetSizeNumber > maxSizeNumber) {
                console.log(`%c計算結果 ${targetSize}(${targetSizeNumber}) 大於最大可用尺碼 ${maxAvailableSize}(${maxSizeNumber})`, 'color: #e74c3c;');
                
                // 如果庫存有最大尺碼，則分配最大尺碼
                if (remainingInventory[maxAvailableSize]?.allocatable >= requiredCount) {
                    finalSize = maxAvailableSize;
                    isSpecialAllocation = true;
                    console.log(`%c使用最大可用尺碼 ${maxAvailableSize} 作為替代`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：計算尺碼過大且無可用替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 的計算尺碼過大且無可用替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    unallocatedStudents.push({ student, requiredCount });
                    continue;
                }
            }
            // 檢查目標尺碼是否存在於庫存中
            else if (!remainingInventory[finalSize]) {
                console.log(`%c計算的目標尺碼 ${finalSize} 在庫存中不存在`, 'color: #e74c3c;');
                
                // 尋找最接近的尺碼
                const sizesWithDistance = availableSizesArray.map(size => {
                    const distance = Math.abs(getSizeNumber(size) - targetSizeNumber);
                    return { size, distance };
                }).sort((a, b) => a.distance - b.distance);
                
                if (sizesWithDistance.length > 0) {
                    finalSize = sizesWithDistance[0].size;
                    isSpecialAllocation = true;
                    console.log(`%c使用最接近的尺碼 ${finalSize} 作為替代，距離=${sizesWithDistance[0].distance}`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：無法找到替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 無法找到替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    unallocatedStudents.push({ student, requiredCount });
                    continue;
                }
            }
            // 檢查目標尺碼庫存是否足夠
            else if (remainingInventory[finalSize].allocatable < requiredCount) {
                // 如果目標尺碼庫存不足，嘗試尋找最接近的尺碼
                console.log(`%c目標尺碼 ${finalSize} 庫存不足，需要 ${requiredCount}件，但只剩 ${remainingInventory[finalSize]?.allocatable || 0}件`, 'color: #e74c3c;');
                
                // 找出所有還有庫存的尺碼
                const sizesWithStock = availableSizesArray.filter(size => 
                    remainingInventory[size] && remainingInventory[size].allocatable >= requiredCount
                );
                
                if (sizesWithStock.length > 0) {
                    // 按照與目標尺碼的差距排序
                    const targetSizeIndex = SIZES.indexOf(finalSize);
                    sizesWithStock.sort((a, b) => {
                        const diffA = Math.abs(SIZES.indexOf(a) - targetSizeIndex);
                        const diffB = Math.abs(SIZES.indexOf(b) - targetSizeIndex);
                        return diffA - diffB;
                    });
                    
                    // 使用最接近的尺碼
                    const originalSize = finalSize;
                    finalSize = sizesWithStock[0];
                    isSpecialAllocation = true;
                    console.log(`%c使用最接近的可用尺碼 ${finalSize} 替代原目標尺碼 ${originalSize}`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：目標尺碼庫存不足且無可用替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 的目標尺碼庫存不足且無可用替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    unallocatedStudents.push({ student, requiredCount });
                    continue;
                }
            } else {
                console.log(`%c目標尺碼 ${finalSize} 存在且庫存充足（需要 ${requiredCount}件，庫存 ${remainingInventory[finalSize]?.allocatable || 0}件）`, 'color: #27ae60;');
            }
            
            // 檢查是否因褲長需要調整尺寸
            let originalSize = finalSize;
            let adjustedSize = finalSize;
            let isAdjusted = false;
                    
            if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, finalSize)) {
                console.log(`%c檢測到褲長調整需求：`, 'color: #e67e22;');
                console.log(`- 學生褲長：${student.pantsLength}`);
                console.log(`- 當前尺寸：${finalSize}`);
                console.log(`- 尺寸數字：${getSizeNumber(finalSize)}`);
                console.log(`- 差值：${student.pantsLength - getSizeNumber(finalSize)}`);
                
                const largerSize = getNextLargerSize(finalSize);
                console.log(`- 建議調整到：${largerSize}`);
                        
                // 檢查更大的尺寸是否有足夠庫存
                if (largerSize !== finalSize && remainingInventory[largerSize]?.allocatable >= requiredCount) {
                    adjustedSize = largerSize;
                    isAdjusted = true;
                    stats.pantsSizeAdjusted++;
                    
                    console.log(`%c褲長調整成功：從 ${finalSize} 調整為 ${largerSize}`, 'color: #27ae60;');
                } else {
                    console.log(`%c褲長調整失敗：${largerSize} 尺寸庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[largerSize]?.allocatable || 0} 件`, 'color: #e74c3c;');
                }
            } else if (student.pantsLength) {
                console.log(`學生褲長(${student.pantsLength})與尺碼(${getSizeNumber(finalSize)})差值<3，無需調整`);
            }
                    
            // 分配調整後的尺寸
            student[allocatedField] = adjustedSize;
            student[specialField] = isSpecialAllocation;
                    
            // 標記褲長調整
            student.isShirtSizeAdjustedForPantsLength = isAdjusted;
            if (isAdjusted) {
                student.originalShirtSize = originalSize;
                student.adjustmentReason = `褲長(${student.pantsLength})比尺碼(${getSizeNumber(originalSize)})大3以上`;
            }
            
            // 減少庫存
            console.log(`%c分配尺寸 ${adjustedSize}：`, 'color: #27ae60;');
            console.log(`- 分配前庫存：${remainingInventory[adjustedSize].allocatable}件`);
                    
            decreaseInventory(remainingInventory, adjustedSize, requiredCount, inventoryType);
            
            console.log(`- 分配後庫存：${remainingInventory[adjustedSize].allocatable}件`);
            console.log(`- 已分配數量：${remainingInventory[adjustedSize].allocated}件`);
            
            stats.allocated++;
            
            if (isSpecialAllocation) {
                stats.special++;
                console.log(`- 統計：特殊分配`);
            } else if (!isAdjusted) {
                stats.exact++;
                console.log(`- 統計：完全符合尺寸`);
            } else {
                stats.different++;
                console.log(`- 統計：不同尺寸`);
            }
            
            allocated = true;
                    
            // 清除失敗原因（如果之前有）
            if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                delete student.allocationFailReason[inventoryType];
                console.log(`- 清除先前的失敗原因`);
            }
                    
            const sizeDisplay = isAdjusted ? `${adjustedSize}↑` : adjustedSize;
            const specialMark = isSpecialAllocation ? '(特殊分配)' : '';
            console.log(`%c短衣分配成功${specialMark}：${student.name}(${student.gender}，座號：${student.number}，胸圍：${student.chest}，腰圍：${student.waist}，褲長：${student.pantsLength}) => 尺寸 ${sizeDisplay}，需求 ${requiredCount}件`, 'color: #27ae60; font-weight: bold;');
            
            // 更新可用尺寸列表
            availableSizes = availableSizes.filter(s => remainingInventory[s.size]?.allocatable > 0);
            if (availableSizes.length === 0) {
                console.warn(`%c所有尺寸庫存都已用完，結束第一階段分配`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
                break;
            }
        }
        
        // 第二階段：處理未分配的學生（嘗試特殊分配）
        console.log(`%c第二階段分配開始：處理 ${unallocatedStudents.length} 名未分配學生`, 'background: #8e44ad; color: white; font-size: 12px; padding: 3px;');
        
        // 重新獲取所有尺寸，包括庫存為0的尺寸（可能某些學生需求數量較小，可以分配）
        let allSizes = Object.keys(remainingInventory)
            .filter(size => remainingInventory[size] && remainingInventory[size].allocatable > 0)
            .map(size => ({
                size,
                available: remainingInventory[size].allocatable
            }))
            .sort((a, b) => SIZES.indexOf(a.size) - SIZES.indexOf(b.size));
        
        console.log(`%c第二階段可用尺寸：`, 'color: #2980b9;');
        allSizes.forEach(s => {
            console.log(`- ${s.size}: ${s.available}件可分配`);
        });
        
        if (allSizes.length > 0) {
            for (const {student, requiredCount} of unallocatedStudents) {
                console.log(`%c特殊分配：處理未分配學生 ${student.name}(${student.gender}，班級：${student.class}，座號：${student.number})`, 'color: #8e44ad;');
                console.log(`- 需求件數：${requiredCount}件`);
                
                // 尋找庫存最多的尺寸
                const bestSize = allSizes.reduce((prev, curr) => 
                    (curr.available > prev.available) ? curr : prev, allSizes[0]);
                
                console.log(`- 庫存最多的尺寸：${bestSize.size}，可分配數量：${bestSize.available}件`);
                
                if (bestSize && bestSize.available >= requiredCount) {
                    // 儲存原始尺寸
                    let originalSize = bestSize.size;
                    let adjustedSize = bestSize.size;
                    let isAdjusted = false;
                    
                    // 檢查是否因褲長需要調整尺寸
                    if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, bestSize.size)) {
                        console.log(`%c檢測到褲長調整需求：`, 'color: #e67e22;');
                        console.log(`- 學生褲長：${student.pantsLength}`);
                        console.log(`- 當前尺寸：${bestSize.size}`);
                        console.log(`- 尺寸數字：${getSizeNumber(bestSize.size)}`);
                        console.log(`- 差值：${student.pantsLength - getSizeNumber(bestSize.size)}`);
                        
                        const largerSize = getNextLargerSize(bestSize.size);
                        console.log(`- 建議調整到：${largerSize}`);
                        
                        // 檢查更大的尺寸是否有足夠庫存
                        if (largerSize !== bestSize.size && remainingInventory[largerSize]?.allocatable >= requiredCount) {
                            adjustedSize = largerSize;
                            isAdjusted = true;
                            stats.pantsSizeAdjusted++;
                            
                            console.log(`%c褲長調整成功：從 ${bestSize.size} 調整為 ${largerSize}`, 'color: #27ae60;');
                        } else {
                            console.log(`%c褲長調整失敗：${largerSize} 尺寸庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[largerSize]?.allocatable || 0} 件`, 'color: #e74c3c;');
                        }
                    }
                    
                    // 分配調整後的尺寸
                    student[allocatedField] = adjustedSize;
                    student[specialField] = true; // 標記為特殊分配
                    
                    // 標記褲長調整
                    student.isShirtSizeAdjustedForPantsLength = isAdjusted;
                    if (isAdjusted) {
                        student.originalShirtSize = originalSize;
                        student.adjustmentReason = `褲長(${student.pantsLength})比尺碼(${getSizeNumber(originalSize)})大3以上`;
                    }
                    
                    // 減少庫存
                    console.log(`%c特殊分配尺寸 ${adjustedSize}：`, 'color: #27ae60;');
                    console.log(`- 分配前庫存：${remainingInventory[adjustedSize].allocatable}件`);
                    
                    decreaseInventory(remainingInventory, adjustedSize, requiredCount, inventoryType);
                    
                    console.log(`- 分配後庫存：${remainingInventory[adjustedSize].allocatable}件`);
                    console.log(`- 已分配數量：${remainingInventory[adjustedSize].allocated}件`);
                    
                    stats.allocated++;
                    stats.special++;
                    
                    // 清除失敗原因（如果之前有）
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                        console.log(`- 清除先前的失敗原因`);
                    }
                    
                    const sizeDisplay = isAdjusted ? `${adjustedSize}↑` : adjustedSize;
                    console.log(`%c短衣特殊分配成功：${student.name}(${student.gender}，座號：${student.number}，胸圍：${student.chest}，褲長：${student.pantsLength}) => 尺寸 ${sizeDisplay}，需求 ${requiredCount}件`, 'color: #27ae60; font-weight: bold;');
                    
                    // 更新尺寸可用數量
                    allSizes = allSizes.map(s => s.size === adjustedSize ? 
                        { ...s, available: remainingInventory[s.size].allocatable } : s)
                        .filter(s => s.available > 0);
                    
                    if (allSizes.length === 0) {
                        console.warn(`%c所有尺寸庫存都已用完，結束第二階段分配`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
                        break;
                    }
                } else {
                    stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：庫存不足';
                    console.log(`%c學生 ${student.name}(${student.class}-${student.number}) 特殊分配失敗：庫存不足`, 'color: #e74c3c; font-weight: bold;');
                }
            }
        } else {
            // 如果沒有任何庫存可用於第二階段，將所有未分配學生標記為失敗
            console.warn(`%c沒有任何尺寸庫存可用於第二階段分配`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
            
            unallocatedStudents.forEach(({student}) => {
                stats.failed++;
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '分配失敗：庫存不足';
                console.log(`%c學生 ${student.name}(${student.class}-${student.number}) 特殊分配失敗：沒有任何可用庫存`, 'color: #e74c3c; font-weight: bold;');
            });
        }
        
        // 確保統計數據包含所有學生
        let totalProcessed = stats.allocated + stats.failed;
        let totalStudentsNeedingAllocation = sortedStudents.filter(
            ({student}) => (student.shortSleeveShirtCount || 1) > 0
        ).length;
        
        // 如果有學生未被處理，將其標記為失敗
        if (totalProcessed < totalStudentsNeedingAllocation) {
            console.warn(`%c檢測到 ${totalStudentsNeedingAllocation - totalProcessed} 名學生未被處理，標記為分配失敗`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
            
            sortedStudents.forEach(({student}) => {
                if (!student[allocatedField] && !student.allocationFailReason?.[inventoryType] && 
                    (student.shortSleeveShirtCount || 1) > 0) {
                    stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：庫存不足';
                    console.log(`%c學生 ${student.name}(${student.class}-${student.number}) 標記為分配失敗：未被處理`, 'color: #e74c3c;');
                }
            });
        }
        
        // 更新庫存資料
        inventoryData[inventoryType] = remainingInventory;
        
        // 更新分配統計
        allocationStats[inventoryType] = stats;
        
        // 總結分配結果
        console.log(`%c${UNIFORM_TYPES[inventoryType]}分配結果統計：`, 'background: #3498db; color: white; font-size: 12px; padding: 5px;');
        console.log(`- 總學生數：${sortedStudents.length}人`);
        console.log(`- 成功分配：${stats.allocated}人 (${((stats.allocated/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 完全符合：${stats.exact}人 (${((stats.exact/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 分配不同尺寸：${stats.different}人 (${((stats.different/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 特殊分配：${stats.special}人 (${((stats.special/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 因褲長調整：${stats.pantsSizeAdjusted}人 (${((stats.pantsSizeAdjusted/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 分配失敗：${stats.failed}人 (${((stats.failed/sortedStudents.length)*100).toFixed(1)}%)`);
        
        console.log(`%c===== ${UNIFORM_TYPES[inventoryType]}分配完成 =====`, 'background: #3498db; color: white; font-size: 14px; padding: 5px;');
        
        resolve(true);
    });
}

/**
 * 分配短褲
 */
function allocateShortPants(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]} ====================`);
        
        // 清除所有學生的短褲調整標記，確保新的分配不受舊數據影響
        _localSortedStudentData.forEach(student => {
            student.pantsAdjustmentMark = null; // 清除短褲調整標記
        });
        console.log(`已清除所有學生的短褲調整標記，確保按照新規則分配`);
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        // 計算總需求量
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.shortSleevePantsCount || 1;
        });
        
        // 計算總可分配數量
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
        
        // 腰圍範圍到尺碼的映射表 - 確保使用與庫存相同的尺碼格式
        const waistToSizeMap = [
            { min: 20, max: 21, size: "XS/34" }, 
            { min: 22, max: 24, size: "S/36" },
            { min: 25, max: 27, size: "M/38", adjustedSize: "L/40", adjustCondition: (student) => student.gender === "男" && student.pantsLength == 40, adjustMark: "L/40↑" },
            { min: 28, max: 30, size: "L/40", adjustedSize: "XL/42", adjustCondition: (student) => student.gender === "男" && student.pantsLength == 41, adjustMark: "XL/42↑" },
            { min: 31, max: 33, size: "XL/42" },
            { min: 34, max: 36, size: "2L/44" },
            { min: 37, max: 38, size: "3L/46" },
            { min: 39, max: 40, size: "4L/48" },
            { min: 41, max: 43, primarySize: "5L/50", alternateSize: "6L/52" },
            { min: 44, max: 46, primarySize: "7L/54", alternateSize: "8L/56" }
        ];
        
        // 成功分配的學生計數
        let successCount = 0;
        
        // 單階段分配過程 - 直接分配而不是先收集
        for (const student of _localSortedStudentData) {
            // 檢查學生是否需要分配短褲
            if (student.shortSleevePantsCount && student.shortSleevePantsCount > 0) {
                console.log(`\n分配學生 [${student.id}] ${student.className}-${student.number} ${student.name}: 腰圍=${student.waist}, 褲長=${student.pantsLength}, 性別=${student.gender}`);
                
                // 根據腰圍找到對應的尺碼範圍
                const waistRange = waistToSizeMap.find(range => 
                    student.waist >= range.min && student.waist <= range.max);
                
                if (!waistRange) {
                    console.warn(`學生 [${student.id}] 腰圍 ${student.waist} 超出分配範圍，無法分配短褲`);
                    // 添加以下代碼設置失敗原因
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：腰圍超出範圍';
                    continue;
                }
                
                // 判斷是否需要根據性別和褲長進行調整
                let sizeToAllocate;
                let adjustmentMark = null;
                
                if (waistRange.adjustCondition && waistRange.adjustCondition(student)) {
                    console.log(`學生 [${student.id}] 符合特殊調整條件: 原尺碼=${waistRange.size}, 調整後尺碼=${waistRange.adjustedSize}`);
                    sizeToAllocate = waistRange.adjustedSize;
                    adjustmentMark = waistRange.adjustMark;
                } else if (waistRange.primarySize) {
                    // 如果是需要先嘗試主要尺碼再嘗試替代尺碼的情況
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
                            continue;
                        }
                    }
                } else {
                    // 一般情況
                    const sizeAvailable = findMatchingInventorySize(waistRange.size);
                    if (sizeAvailable) {
                        console.log(`學生 [${student.id}] 使用標準尺碼: ${waistRange.size}`);
                        sizeToAllocate = sizeAvailable;
                    } else {
                        console.warn(`學生 [${student.id}] 尺碼 ${waistRange.size} 無庫存`);
                        // 設置失敗原因
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '分配失敗：尺碼無庫存';
                        continue;
                    }
                }
                
                // 確保找到了可用尺碼
                if (!sizeToAllocate) {
                    console.warn(`學生 [${student.id}] 無法分配短褲：找不到合適的尺碼或庫存不足`);
                    // 設置失敗原因
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：無合適尺碼';
                    continue;
                }
                
                // 檢查庫存是否足夠並直接分配
                const requiredCount = student.shortSleevePantsCount || 1;
                
                // 直接嘗試減少庫存並分配 - 單階段分配
                if (decreaseInventory(inventoryData[inventoryType], sizeToAllocate, requiredCount, inventoryType)) {
                    // 分配成功，更新學生資料
                    student[allocatedField] = sizeToAllocate;
                    if (adjustmentMark) {
                        student.pantsAdjustmentMark = adjustmentMark;
                    }
                    // 確保不使用一般的褲長調整標記
                    student.isPantsLengthAdjusted = false;
                    
                    // 清除失敗原因（如果之前有）
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    
                    successCount++;
                    console.log(`分配成功: 學生 [${student.id}] ${student.className}-${student.number} ${student.name} - 尺碼 ${sizeToAllocate}${adjustmentMark ? ' (標記: ' + adjustmentMark + ')' : ''}, 需求 ${requiredCount} 件`);
                } else {
                    // 分配失敗 - 庫存不足
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：庫存不足';
                    console.warn(`分配失敗: 學生 [${student.id}] ${student.className}-${student.number} ${student.name} - 尺碼 ${sizeToAllocate} 庫存不足, 需求 ${requiredCount} 件`);
                }
            }
        }
        
        console.log(`${UNIFORM_TYPES[inventoryType]} 分配完成，共分配給 ${successCount} 位學生 ====================`);
        resolve(true);
    });
}

/**
 * 分配長袖上衣
 */
function allocateLongShirts(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`%c===== 開始分配 ${UNIFORM_TYPES[inventoryType]} =====`, 'background: #9b59b6; color: white; font-size: 14px; padding: 5px;');
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            
            // 將所有學生標記為分配失敗（無庫存）
            _localSortedStudentData.forEach(student => {
                if (!student[allocatedField]) {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '無庫存資料';
                }
            });
            
            allocationStats[inventoryType].failed = _localSortedStudentData.length;
            resolve(false);
            return;
        }
        
        // 計算總需求量
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.longSleeveShirtCount || 1;
        });
        
        // 計算總可分配數量
        let totalAllocatable = 0;
        for (const size in inventoryData[inventoryType]) {
            totalAllocatable += inventoryData[inventoryType][size].allocatable || 0;
        }
        
        console.log(`%c${UNIFORM_TYPES[inventoryType]}需求與庫存概況：`, 'color: #8e44ad; font-weight: bold;');
        console.log(`- 總需求量：${totalDemand}件`);
        console.log(`- 總可分配數量：${totalAllocatable}件`);
        console.log(`- 學生數量：${_localSortedStudentData.length}人`);
        
        // 如果總可分配數小於總需求量，發出警告
        if (totalAllocatable < totalDemand) {
            console.warn(`%c警告：${UNIFORM_TYPES[inventoryType]}可分配數總和(${totalAllocatable})小於需求量(${totalDemand})，差額${totalDemand - totalAllocatable}件`, 'color: #e74c3c; font-weight: bold;');
        }
        
        // 獲取可用尺寸並排序（小到大）
        let availableSizes = getAvailableSizes(inventoryData, inventoryType);
        
        // 過濾掉可分配數為0的尺寸
        availableSizes = availableSizes.filter(size => size.available > 0);
        
        console.log(`%c可用尺寸列表：`, 'color: #8e44ad; font-weight: bold;');
        availableSizes.forEach(s => {
            console.log(`- ${s.size}: ${s.available}件可分配`);
        });
        
        if (availableSizes.length === 0) {
            console.warn(`%c沒有可用的 ${UNIFORM_TYPES[inventoryType]} 庫存，無法進行分配`, 'color: #e74c3c; font-weight: bold;');
            
            // 將所有學生標記為分配失敗（尺寸無庫存）
            _localSortedStudentData.forEach(student => {
                if (!student[allocatedField]) {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '所有尺寸無庫存';
                }
            });
            
            allocationStats[inventoryType].failed = _localSortedStudentData.length;
            resolve(false);
            return;
        }
        
        // 複製庫存資料用於分配
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        // 複製學生數組並根據有效胸圍排序（未加上性別調整值）
        const sortedStudents = [..._localSortedStudentData].map(student => {
            // 取胸圍和腰圍的較大值作為有效胸圍（未加性別調整值）
            const chest = student.chest || 0;
            const waist = student.waist || 0;
            const effectiveChest = Math.max(chest, waist);
            return { student, effectiveChest };
        }).sort((a, b) => a.effectiveChest - b.effectiveChest);

        console.log(`%c學生排序結果（按有效胸圍排序）：`, 'color: #8e44ad; font-weight: bold;');
        console.log(`學生總數：${sortedStudents.length}人`);
        console.log(`前5名學生：`);
        sortedStudents.slice(0, 5).forEach((s, i) => {
            console.log(`${i+1}. ${s.student.name}(${s.student.gender}，班級：${s.student.class}，座號：${s.student.number})：` +
                      `胸圍=${s.student.chest}，腰圍=${s.student.waist}，有效胸圍=${s.effectiveChest}`);
        });
        
        // 用於記錄分配結果的統計資料
        const stats = {
            allocated: 0,
            exact: 0,
            different: 0,
            failed: 0,
            special: 0,
            pantsSizeAdjusted: 0 // 新增：因褲長調整尺寸的學生數
        };
        
        // 分配按照排序進行
        console.log(`%c開始分配：按有效胸圍順序分配`, 'background: #8e44ad; color: white; font-size: 12px; padding: 3px;');
        
        // 檢查短袖上衣和短褲分配失敗的學生
        const priorFailedStudents = sortedStudents
            .filter(({student}) => 
                student.allocationFailReason && 
                (student.allocationFailReason.shortSleeveShirt || student.allocationFailReason.shortSleevePants));
        
        if (priorFailedStudents.length > 0) {
            console.log(`%c長衣分配開始前，檢測到 ${priorFailedStudents.length} 名學生先前分配有失敗記錄：`, 'background: #f39c12; color: white; font-size: 12px; padding: 3px;');
            
            // 顯示不同先前階段失敗的學生數量
            const shortShirtFailed = priorFailedStudents.filter(({student}) => 
                student.allocationFailReason.shortSleeveShirt).length;
            const shortPantsFailed = priorFailedStudents.filter(({student}) => 
                student.allocationFailReason.shortSleevePants).length;
            const bothFailed = priorFailedStudents.filter(({student}) => 
                student.allocationFailReason.shortSleeveShirt && student.allocationFailReason.shortSleevePants).length;
            
            console.log(`  - 短袖上衣分配失敗: ${shortShirtFailed} 名學生`);
            console.log(`  - 短褲分配失敗: ${shortPantsFailed} 名學生`);
            console.log(`  - 短袖上衣和短褲都分配失敗: ${bothFailed} 名學生`);
        }
        
        for (const {student, effectiveChest} of sortedStudents) {
            // 跳過已分配的學生
            if (student[allocatedField]) {
                console.log(`學生 ${student.name}(${student.class}-${student.number}) 已有分配結果：${student[allocatedField]}，跳過`);
                continue;
            }
            
            // 檢查先前分配失敗情況
            const hasPriorFailure = student.allocationFailReason && 
                (student.allocationFailReason.shortSleeveShirt || student.allocationFailReason.shortSleevePants);
            
            if (hasPriorFailure) {
                const priorReasons = [];
                if (student.allocationFailReason.shortSleeveShirt) {
                    priorReasons.push(`短袖上衣: ${student.allocationFailReason.shortSleeveShirt}`);
                }
                if (student.allocationFailReason.shortSleevePants) {
                    priorReasons.push(`短褲: ${student.allocationFailReason.shortSleevePants}`);
                }
                console.log(`%c關注：學生 ${student.name}(${student.class}-${student.number}) 的先前分配有失敗記錄，但仍嘗試分配長衣：%c`, 'background: #e67e22; color: white; font-size: 12px; padding: 3px;', '');
                priorReasons.forEach(reason => console.log(`  - ${reason}`));
            }
            
            const requiredCount = student.longSleeveShirtCount || 1;
            
            // 詳細記錄學生資訊
            console.log(`%c處理學生：${student.name}(${student.gender}，班級：${student.class}，座號：${student.number})`, 'color: #8e44ad; font-weight: bold;');
            console.log(`- 胸圍：${student.chest}`);
            console.log(`- 腰圍：${student.waist}`);
            console.log(`- 褲長：${student.pantsLength}`);
            console.log(`- 有效胸圍：${effectiveChest}`);
            console.log(`- 需求件數：${requiredCount}`);
            
            // 如果學生不需要此類型制服（件數為0），標記為不需要
            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生不需要此制服（需求件數為0），跳過分配`);
                continue;
            }
            
            // 計算理論上的尺碼數值（根據有效胸圍和性別）
            let adjustment = 0;
            if (student.gender === '男') {
                adjustment = (effectiveChest % 2 === 0) ? 10 : 11;
            } else {
                adjustment = (effectiveChest % 2 === 0) ? 8 : 9;
            }
            const targetSizeNumber = effectiveChest + adjustment;
            
            console.log(`%c計算理論尺碼：`, 'color: #2c3e50;');
            console.log(`- 有效胸圍：${effectiveChest}`);
            console.log(`- 性別：${student.gender}`);
            console.log(`- 性別調整值：${adjustment}`);
            console.log(`- 計算結果：${effectiveChest} + ${adjustment} = ${targetSizeNumber}`);
            
            // 找出理論上的尺碼代碼
            let targetSize = '';
            if (targetSizeNumber <= 34) targetSize = 'XS/34';
            else if (targetSizeNumber <= 36) targetSize = 'S/36';
            else if (targetSizeNumber <= 38) targetSize = 'M/38';
            else if (targetSizeNumber <= 40) targetSize = 'L/40';
            else if (targetSizeNumber <= 42) targetSize = 'XL/42';
            else if (targetSizeNumber <= 44) targetSize = '2L/44';
            else if (targetSizeNumber <= 46) targetSize = '3L/46';
            else if (targetSizeNumber <= 48) targetSize = '4L/48';
            else if (targetSizeNumber <= 50) targetSize = '5L/50';
            else if (targetSizeNumber <= 52) targetSize = '6L/52';
            else if (targetSizeNumber <= 54) targetSize = '7L/54';
            else if (targetSizeNumber <= 56) targetSize = '8L/56';
            else targetSize = '9L/58';
            
            console.log(`- 對應尺碼：${targetSize}`);
            
            let allocated = false;
            
            // 處理邊界情況：計算結果低於最小尺碼或超過最大尺碼
            const availableSizesArray = availableSizes.map(s => s.size);
            const minAvailableSize = availableSizesArray[0];
            const maxAvailableSize = availableSizesArray[availableSizesArray.length - 1];
            const minSizeNumber = getSizeNumber(minAvailableSize);
            const maxSizeNumber = getSizeNumber(maxAvailableSize);
            
            console.log(`%c檢查尺碼邊界情況：`, 'color: #2c3e50;');
            console.log(`- 最小可用尺碼：${minAvailableSize}(${minSizeNumber})`);
            console.log(`- 最大可用尺碼：${maxAvailableSize}(${maxSizeNumber})`);
            console.log(`- 計算的目標尺碼：${targetSize}(${getSizeNumber(targetSize)})`);
            
            // 檢查計算的尺碼是否在可用範圍內
            let finalSize = targetSize;
            let isSpecialAllocation = false;
            
            // 如果計算結果低於最小可用尺碼
            if (targetSizeNumber < minSizeNumber) {
                console.log(`%c計算結果 ${targetSize}(${targetSizeNumber}) 小於最小可用尺碼 ${minAvailableSize}(${minSizeNumber})`, 'color: #e74c3c;');
                
                // 如果庫存有最小尺碼，則分配最小尺碼
                if (remainingInventory[minAvailableSize]?.allocatable >= requiredCount) {
                    finalSize = minAvailableSize;
                    isSpecialAllocation = true;
                    console.log(`%c使用最小可用尺碼 ${minAvailableSize} 作為替代`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：計算尺碼過小且無可用替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 的計算尺碼過小且無可用替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    continue;
                }
            }
            // 如果計算結果超過最大可用尺碼
            else if (targetSizeNumber > maxSizeNumber) {
                console.log(`%c計算結果 ${targetSize}(${targetSizeNumber}) 大於最大可用尺碼 ${maxAvailableSize}(${maxSizeNumber})`, 'color: #e74c3c;');
                
                // 如果庫存有最大尺碼，則分配最大尺碼
                if (remainingInventory[maxAvailableSize]?.allocatable >= requiredCount) {
                    finalSize = maxAvailableSize;
                    isSpecialAllocation = true;
                    console.log(`%c使用最大可用尺碼 ${maxAvailableSize} 作為替代`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：計算尺碼過大且無可用替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 的計算尺碼過大且無可用替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    continue;
                }
            }
            // 檢查目標尺碼是否存在於庫存中
            else if (!remainingInventory[finalSize]) {
                console.log(`%c計算的目標尺碼 ${finalSize} 在庫存中不存在`, 'color: #e74c3c;');
                
                // 尋找最接近的尺碼
                const sizesWithDistance = availableSizesArray.map(size => {
                    const distance = Math.abs(getSizeNumber(size) - targetSizeNumber);
                    return { size, distance };
                }).sort((a, b) => a.distance - b.distance);
                
                if (sizesWithDistance.length > 0) {
                    finalSize = sizesWithDistance[0].size;
                    isSpecialAllocation = true;
                    console.log(`%c使用最接近的尺碼 ${finalSize} 作為替代，距離=${sizesWithDistance[0].distance}`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：無法找到替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 無法找到替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    continue;
                }
            }
            // 檢查目標尺碼庫存是否足夠
            else if (remainingInventory[finalSize].allocatable < requiredCount) {
                // 如果目標尺碼庫存不足，嘗試尋找最接近的尺碼
                console.log(`%c目標尺碼 ${finalSize} 庫存不足，需要 ${requiredCount}件，但只剩 ${remainingInventory[finalSize]?.allocatable || 0}件`, 'color: #e74c3c;');
                
                // 找出所有還有庫存的尺碼
                const sizesWithStock = availableSizesArray.filter(size => 
                    remainingInventory[size] && remainingInventory[size].allocatable >= requiredCount
                );
                
                if (sizesWithStock.length > 0) {
                    // 按照與目標尺碼的差距排序
                    const targetSizeIndex = SIZES.indexOf(finalSize);
                    sizesWithStock.sort((a, b) => {
                        const diffA = Math.abs(SIZES.indexOf(a) - targetSizeIndex);
                        const diffB = Math.abs(SIZES.indexOf(b) - targetSizeIndex);
                        return diffA - diffB;
                    });
                    
                    // 使用最接近的尺碼
                    const originalSize = finalSize;
                    finalSize = sizesWithStock[0];
                    isSpecialAllocation = true;
                    console.log(`%c使用最接近的可用尺碼 ${finalSize} 替代原目標尺碼 ${originalSize}`, 'color: #27ae60;');
                } else {
                    // 標記為分配失敗
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：目標尺碼庫存不足且無可用替代尺碼';
                    stats.failed++;
                    console.log(`%c分配失敗：學生 ${student.name} 的目標尺碼庫存不足且無可用替代尺碼`, 'color: #e74c3c; font-weight: bold;');
                    continue;
                }
            } else {
                console.log(`%c目標尺碼 ${finalSize} 存在且庫存充足（需要 ${requiredCount}件，庫存 ${remainingInventory[finalSize]?.allocatable || 0}件）`, 'color: #27ae60;');
            }
            
            // 檢查是否因褲長需要調整尺寸
            let originalSize = finalSize;
            let adjustedSize = finalSize;
            let isAdjusted = false;
            
            if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, finalSize)) {
                console.log(`%c檢測到褲長調整需求：`, 'color: #e67e22;');
                console.log(`- 學生褲長：${student.pantsLength}`);
                console.log(`- 當前尺寸：${finalSize}`);
                console.log(`- 尺寸數字：${getSizeNumber(finalSize)}`);
                console.log(`- 差值：${student.pantsLength - getSizeNumber(finalSize)}`);
                
                const largerSize = getNextLargerSize(finalSize);
                console.log(`- 建議調整到：${largerSize}`);
                
                // 檢查更大的尺寸是否有足夠庫存
                if (largerSize !== finalSize && remainingInventory[largerSize]?.allocatable >= requiredCount) {
                    adjustedSize = largerSize;
                    isAdjusted = true;
                    stats.pantsSizeAdjusted++;
                    
                    console.log(`%c褲長調整成功：從 ${finalSize} 調整為 ${largerSize}`, 'color: #27ae60;');
                } else {
                    console.log(`%c褲長調整失敗：${largerSize} 尺寸庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[largerSize]?.allocatable || 0} 件`, 'color: #e74c3c;');
                }
            } else if (student.pantsLength) {
                console.log(`學生褲長(${student.pantsLength})與尺碼(${getSizeNumber(finalSize)})差值<3，無需調整`);
            }
            
            // 分配調整後的尺寸
            student[allocatedField] = adjustedSize;
            student[specialField] = isSpecialAllocation;
            
            // 標記褲長調整
            student.isLongShirtSizeAdjustedForPantsLength = isAdjusted;
            if (isAdjusted) {
                student.originalLongShirtSize = originalSize;
                student.longShirtAdjustmentReason = `褲長(${student.pantsLength})比尺碼(${getSizeNumber(originalSize)})大3以上`;
            }
            
            // 減少庫存
            console.log(`%c分配尺寸 ${adjustedSize}：`, 'color: #27ae60;');
            console.log(`- 分配前庫存：${remainingInventory[adjustedSize].allocatable}件`);
            
            decreaseInventory(remainingInventory, adjustedSize, requiredCount, inventoryType);
            
            console.log(`- 分配後庫存：${remainingInventory[adjustedSize].allocatable}件`);
            console.log(`- 已分配數量：${remainingInventory[adjustedSize].allocated}件`);
            
            stats.allocated++;
            
            if (isSpecialAllocation) {
                stats.special++;
                console.log(`- 統計：特殊分配`);
            } else if (!isAdjusted) {
                stats.exact++;
                console.log(`- 統計：完全符合尺寸`);
            } else {
                stats.different++;
                console.log(`- 統計：不同尺寸`);
            }
            
            allocated = true;
            
            // 清除失敗原因（如果之前有）
            if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                delete student.allocationFailReason[inventoryType];
                console.log(`- 清除先前的失敗原因`);
            }
            
            const sizeDisplay = isAdjusted ? `${adjustedSize}↑` : adjustedSize;
            const specialMark = isSpecialAllocation ? '(特殊分配)' : '';
            console.log(`%c長衣分配成功${specialMark}：${student.name}(${student.gender}，座號：${student.number}，胸圍：${student.chest}，腰圍：${student.waist}，褲長：${student.pantsLength}) => 尺寸 ${sizeDisplay}，需求 ${requiredCount}件`, 'color: #27ae60; font-weight: bold;');
            
            if (hasPriorFailure) {
                console.log(`%c特別關注：儘管先前分配有失敗記錄，學生 ${student.name}(${student.class}-${student.number}) 仍然分配到長衣尺寸 ${adjustedSize}%c`, 'background: #2ecc71; color: white; font-size: 12px; padding: 3px;', '');
            }
            
            // 更新可用尺寸列表
            availableSizes = availableSizes.filter(s => remainingInventory[s.size]?.allocatable > 0);
            if (availableSizes.length === 0) {
                console.warn(`%c所有尺寸庫存都已用完，結束分配`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
                break;
            }
        }
        
        // 更新分配統計資料
        allocationStats[inventoryType] = {
            allocated: stats.allocated,
            exact: stats.exact,
            different: stats.different,
            special: stats.special,
            failed: _localSortedStudentData.length - stats.allocated,
            pantsSizeAdjusted: stats.pantsSizeAdjusted
        };
        
        // 在分配完成後添加詳細的統計日誌
        console.log(`%c${UNIFORM_TYPES[inventoryType]}分配結果統計：`, 'background: #9b59b6; color: white; font-size: 12px; padding: 5px;');
        console.log(`- 總學生數：${sortedStudents.length}人`);
        console.log(`- 成功分配：${stats.allocated}人 (${((stats.allocated/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 完全符合：${stats.exact}人 (${((stats.exact/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 分配不同尺寸：${stats.different}人 (${((stats.different/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 特殊分配：${stats.special}人 (${((stats.special/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 因褲長調整：${stats.pantsSizeAdjusted}人 (${((stats.pantsSizeAdjusted/sortedStudents.length)*100).toFixed(1)}%)`);
        console.log(`- 分配失敗：${stats.failed}人 (${((stats.failed/sortedStudents.length)*100).toFixed(1)}%)`);
        
        console.log(`%c===== ${UNIFORM_TYPES[inventoryType]}分配完成 =====`, 'background: #9b59b6; color: white; font-size: 14px; padding: 5px;');
        
        resolve(stats.allocated > 0);
    });
}

/**
 * 根據尺碼代碼獲取實際褲長（厘米）
 * 用於長褲分配時的褲長比較
 */
function getPantsLengthInCm(size) {
    // 根據提供的尺碼對照表
    const sizeMap = {
        'XS/34': 34,
        'S/36': 36,
        'M/38': 38,
        'L/40': 40,
        'XL/42': 42,
        '2L/44': 44,
        '3L/46': 46,
        '4L/48': 48,
        '6L/52': 52
    };
    
    // 如果在對照表中找到尺碼，返回對應的厘米值
    if (sizeMap[size]) {
        return sizeMap[size];
    }
    
    // 如果是數字或數字字符串，直接返回
    if (!isNaN(size)) {
        return parseInt(size, 10);
    }
    
    // 嘗試從尺碼中提取數字部分
    const match = size.match(/(\d+)/);
    if (match) {
        return parseInt(match[0], 10);
    }
    
    // 無法確定尺碼，返回0（無效值）
    console.warn(`無法確定尺碼 ${size} 的長褲長度`);
    return 0;
}

/**
 * 檢查長褲尺寸與長袖上衣尺寸的階級差異
 * @param {string} pantsSize - 長褲尺寸
 * @param {string} shirtSize - 長袖上衣尺寸
 * @returns {boolean} - 是否符合階級差異要求
 */
function isPantsLengthAcceptable(studentPantsLength, pantsSize) {
    // 移除適合性判斷，改為始終返回true以允許所有尺寸分配
    return true;
}

/**
 * 檢查學生褲長是否需要增加尺碼
 * @param {number} studentPantsLength - 學生的褲長(cm)
 * @param {string} currentSize - 當前分配的尺碼
 * @returns {string} - 可能調整後的尺碼
 */
function adjustSizeForPantsLength(studentPantsLength, currentSize) {
    const sizeLengthInCm = getPantsLengthInCm(currentSize);
    
    // 如果學生褲長比尺碼長度大於或等於3厘米，自動增加一個尺碼
    if (studentPantsLength - sizeLengthInCm >= 3) {
        console.log(`檢測到學生褲長 ${studentPantsLength}cm 比尺碼 ${currentSize}(${sizeLengthInCm}cm) 大於或等於3厘米，自動增加一個尺碼`);
        return getNextLargerSize(currentSize);
    }
    
    return currentSize;
}

/**
 * 分配長袖褲子
 */
function allocateLongPants(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]}`);
        
        // 清除所有學生的長褲調整標記，確保新的分配不受舊數據影響
        _localSortedStudentData.forEach(student => {
            student.longPantsAdjustmentMark = null; // 清除長褲調整標記
        });
        console.log(`已清除所有學生的長褲調整標記，確保按照新規則分配`);
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        // 計算總需求量
        let totalDemand = 0;
        _localSortedStudentData.forEach(student => {
            totalDemand += student.longSleevePantsCount || 1;
        });
        
        // 計算總可分配數量
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
        
        // 將新表格插入到結果標籤的內容開始處
        const cardBody = resultTab.querySelector('.card-body');
        if (cardBody) {
            cardBody.insertBefore(detailSection, cardBody.firstChild);
        } else {
            // 如果找不到 card-body，則插入到結果標籤的開頭
            resultTab.insertBefore(detailSection, resultTab.firstChild);
        }
        
        detailTable = document.getElementById('studentDetailTable');
        
        // 添加匯出按鈕的事件監聽器
        const exportBtn = document.getElementById('exportAllocationResultsBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', exportAllocationResultsToExcel);
        }
    }

    // 獲取表格主體
    const tbody = detailTable.querySelector('tbody');
    if (!tbody) return;

    // 清空現有內容
    tbody.innerHTML = '';

    // 按照班級和座號排序學生數據
    const sortedStudents = [...studentData].sort((a, b) => {
        const classA = parseInt(a.className) || 0;
        const classB = parseInt(b.className) || 0;
        if (classA !== classB) return classA - classB;
        
        const numberA = parseInt(a.number) || 0;
        const numberB = parseInt(b.number) || 0;
        return numberA - numberB;
    });

    // 添加學生數據
    sortedStudents.forEach((student, index) => {
        const row = document.createElement('tr');
        
        // 檢查是否存在分配失敗的原因
        const shortShirtFailReason = student.allocationFailReason?.shortSleeveShirt || '';
        const shortPantsFailReason = student.allocationFailReason?.shortSleevePants || '';
        const longShirtFailReason = student.allocationFailReason?.longSleeveShirt || '';
        const longPantsFailReason = student.allocationFailReason?.longSleevePants || '';
        
        // 使用formatSize函數格式化尺寸
        const formattedShirtSize = student.allocatedShirtSize ? formatSize(student.allocatedShirtSize) : '-';
        const formattedPantsSize = student.allocatedPantsSize ? formatSize(student.allocatedPantsSize) : '-';
        const formattedLongShirtSize = student.allocatedLongShirtSize ? formatSize(student.allocatedLongShirtSize) : '-';
        const formattedLongPantsSize = student.allocatedLongPantsSize ? formatSize(student.allocatedLongPantsSize) : '-';
        
        // 設置行內容
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
                ${formattedShirtSize}${student.isShirtSizeAdjustedForPantsLength ? '<sup>↑</sup>' : ''}
                ${shortShirtFailReason ? `<div class="failure-reason">${shortShirtFailReason}</div>` : ''}
                ${student.isShirtSizeAdjustedForPantsLength ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
                ${student.shortShirtLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥2</div>` : ''}
            </td>
            <td class="count-column">${student.allocatedShirtSize ? (student.shortSleeveShirtCount || 1) : '-'}</td>
            <td>
                ${student.pantsAdjustmentMark || formattedPantsSize}
                ${shortPantsFailReason ? `<div class="failure-reason">${shortPantsFailReason}</div>` : ''}
                ${student.shortPantsLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥2</div>` : ''}
            </td>
            <td class="count-column">${student.allocatedPantsSize ? (student.shortSleevePantsCount || 1) : '-'}</td>
            <td>
                ${formattedLongShirtSize}${student.isLongShirtSizeAdjustedForPantsLength ? '<sup>↑</sup>' : ''}
                ${longShirtFailReason ? `<div class="failure-reason">${longShirtFailReason}</div>` : ''}
                ${student.isLongShirtSizeAdjustedForPantsLength ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
                ${student.longShirtLengthDeficiency ? `<div class="adjustment-reason" style="color: #e67e22; font-size: 0.85em;">褲長仍不足≥2</div>` : ''}
            </td>
            <td class="count-column">${student.allocatedLongShirtSize ? (student.longSleeveShirtCount || 1) : '-'}</td>
            <td>
                ${student.longPantsAdjustmentMark || formattedLongPantsSize}
                ${longPantsFailReason ? `<div class="failure-reason">${longPantsFailReason}</div>` : ''}
            </td>
            <td class="count-column">${student.allocatedLongPantsSize ? (student.longSleevePantsCount || 1) : '-'}</td>
        `;

        tbody.appendChild(row);
    });
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

    // 添加學生數據
    sortedStudents.forEach((student, index) => {
        // 處理各種制服的分配情況 - 使用formatSize函數格式化尺寸
        let shortShirtSize = student.allocatedShirtSize ? formatSize(student.allocatedShirtSize) : '-';
        // 如果有褲長調整，添加標記
        if (student.isShirtSizeAdjustedForPantsLength && shortShirtSize !== '-') {
            shortShirtSize += '↑(褲長調整)';
        }
        // 如果短衣褲長仍不足，添加標記
        if (student.shortShirtLengthDeficiency && shortShirtSize !== '-') {
            shortShirtSize += '!(褲長仍不足≥2)';
        }
        
        let shortPantsSize = student.allocatedPantsSize ? formatSize(student.allocatedPantsSize) : '-';
        // 如果短褲有調整標記，優先使用標記
        if (student.pantsAdjustmentMark && shortPantsSize !== '-') {
            shortPantsSize = student.pantsAdjustmentMark;
        }
        // 移除對於短褲的一般褲長調整標記，只使用特定標記
        // 如果短褲褲長仍不足，添加標記
        if (student.shortPantsLengthDeficiency && shortPantsSize !== '-') {
            shortPantsSize += '!(褲長仍不足≥2)';
        }
        
        let longShirtSize = student.allocatedLongShirtSize ? formatSize(student.allocatedLongShirtSize) : '-';
        // 如果長衣有褲長調整，添加標記
        if (student.isLongShirtSizeAdjustedForPantsLength && longShirtSize !== '-') {
            longShirtSize += '↑(褲長調整)';
        }
        // 如果長衣褲長仍不足，添加標記
        if (student.longShirtLengthDeficiency && longShirtSize !== '-') {
            longShirtSize += '!(褲長仍不足≥2)';
        }
        
        let longPantsSize = student.allocatedLongPantsSize ? formatSize(student.allocatedLongPantsSize) : '-';
        // 如果長褲有調整標記，優先使用調整標記
        if (student.longPantsAdjustmentMark && longPantsSize !== '-') {
            longPantsSize = student.longPantsAdjustmentMark;
        }
        
        // 如果有分配失敗原因，直接顯示在對應欄位
        if (student.allocationFailReason) {
            if (student.allocationFailReason.shortSleeveShirt && !student.allocatedShirtSize) {
                shortShirtSize = student.allocationFailReason.shortSleeveShirt;
            }
            if (student.allocationFailReason.shortSleevePants && !student.allocatedPantsSize) {
                shortPantsSize = student.allocationFailReason.shortSleevePants;
            }
            if (student.allocationFailReason.longSleeveShirt && !student.allocatedLongShirtSize) {
                longShirtSize = student.allocationFailReason.longSleeveShirt;
            }
            if (student.allocationFailReason.longSleevePants && !student.allocatedLongPantsSize) {
                longPantsSize = student.allocationFailReason.longSleevePants;
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