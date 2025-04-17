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

/**
 * 開始制服分配
 */
export async function startAllocation() {
    console.log('開始制服分配程序');
    
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
    
    // 重置所有分配結果
    resetAllocation();
    
    // 檢查是否有學生資料
    if (!sortedStudentData || sortedStudentData.length === 0) {
        console.warn('沒有學生資料可供分配');
        throw new Error('沒有學生資料可供分配');
    }
    
    // 檢查是否有庫存資料
    if (!inventoryData) {
        console.warn('沒有庫存資料可供分配');
        throw new Error('沒有庫存資料可供分配');
    }
    
    console.log(`開始為 ${sortedStudentData.length} 名學生分配制服`);
    
    let allocationSuccess = true;
    
    try {
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
        
        if (allocationSuccess) {
            console.log('制服分配完成');
        } else {
            console.warn('制服分配部分完成，有些類型的分配過程發生錯誤');
        }
    } catch (error) {
        console.error('分配過程發生錯誤:', error);
        throw error;
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
        
        for (const size in inventoryData[type]) {
            if (!inventoryData[type].hasOwnProperty(size)) continue;
            
            const total = inventoryData[type][size].total || 0;
            const reserved = inventoryData[type][size].reserved || 0;
            
            // 重置已分配數量為0
            inventoryData[type][size].allocated = 0;
            // 重置可分配數量為總數減去預留數
            inventoryData[type][size].allocatable = total - reserved;
            
            console.log(`重置庫存 ${type}-${size}: 總數=${total}, 預留=${reserved}, 可分配=${inventoryData[type][size].allocatable}`);
        }
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
    // 女生胸圍調整，使用配置的調整值
    if (student.gender === '女') {
        effectiveChest += getFemaleChestAdjustment(); // 使用學校配置的調整值
    }
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
 */
function decreaseInventory(inventory, size, count, inventoryType) {
    if (inventory[size]) {
        const actualCount = Math.min(count, inventory[size].allocatable);
        inventory[size].allocatable -= actualCount;
        inventory[size].allocated += actualCount;
        
        // 輸出實際減少的庫存量
        console.log(`減少庫存 ${size}: ${actualCount} 件，剩餘 ${inventory[size].allocatable} 件，總分配 ${inventory[size].allocated} 件`);
        
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
    }
}

/**
 * 分配短袖上衣
 */
function allocateShortShirts(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]}`);
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            
            // 將所有學生標記為分配失敗（無庫存）
            sortedStudentData.forEach(student => {
                if (!student[allocatedField]) {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '無庫存資料';
                }
            });
            
            allocationStats[inventoryType].failed = sortedStudentData.length;
            resolve(false);
            return;
        }
        
        // 獲取可用尺寸並排序（小到大）
        let availableSizes = getAvailableSizes(inventoryData, inventoryType);
        
        // 過濾掉可分配數為0的尺寸
        availableSizes = availableSizes.filter(size => size.available > 0);
        
        console.log(`可用尺寸：`, availableSizes.map(s => `${s.size}(${s.available})`).join(', '));
        
        if (availableSizes.length === 0) {
            console.warn(`沒有可用的 ${UNIFORM_TYPES[inventoryType]} 庫存`);
            
            // 將所有學生標記為分配失敗（尺寸無庫存）
            sortedStudentData.forEach(student => {
                if (!student[allocatedField]) {
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '所有尺寸無庫存';
                }
            });
            
            allocationStats[inventoryType].failed = sortedStudentData.length;
            resolve(false);
            return;
        }
        
        // 複製庫存資料用於分配
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        // 複製學生數組並根據有效胸圍排序
        const sortedStudents = [...sortedStudentData].map(student => ({
            student,
            effectiveChest: calculateEffectiveChest(student)
        })).sort((a, b) => a.effectiveChest - b.effectiveChest);

        console.log(`學生排序（前5名）：`, 
            sortedStudents.slice(0, 5).map(s => 
                `${s.student.name}(${s.student.className}-${s.student.number}): 有效胸圍=${s.effectiveChest}`
            )
        );
        
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
        console.log('第一階段分配開始：按胸圍順序分配');
        
        // 添加檢查已存在的失敗原因
        const studentsWithPriorFailures = sortedStudents
            .filter(({student}) => student.allocationFailReason && Object.keys(student.allocationFailReason).length > 0);
        
        if (studentsWithPriorFailures.length > 0) {
            console.log(`%c注意：檢測到 ${studentsWithPriorFailures.length} 名學生在之前的步驟已有失敗記錄：`, 'background: #f39c12; color: white; font-size: 12px; padding: 3px;');
            studentsWithPriorFailures.forEach(({student}) => {
                const reasons = Object.entries(student.allocationFailReason)
                    .map(([type, reason]) => `${UNIFORM_TYPES[type]}: ${reason}`)
                    .join(', ');
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 之前失敗原因: ${reasons}`);
            });
        }
        
        for (const {student} of sortedStudents) {
            // 跳過已分配的學生
            if (student[allocatedField]) continue;
            
            // 檢查學生是否需要此類型的制服
            const requiredCount = student.shortSleeveShirtCount || 1;
            
            // 如果學生不需要此類型制服（件數為0），標記為不需要
            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                continue;
            }
            
            let allocated = false;
            
            // 尋找可用的尺寸
            for (let i = 0; i < availableSizes.length; i++) {
                const size = availableSizes[i].size;
                
                // 檢查庫存是否足夠
                if (remainingInventory[size]?.allocatable >= requiredCount) {
                    // 儲存原始尺寸
                    let originalSize = size;
                    let adjustedSize = size;
                    let isAdjusted = false;
                    
                    // 檢查是否因褲長需要調整尺寸
                    if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, size)) {
                        const largerSize = getNextLargerSize(size);
                        
                        // 檢查更大的尺寸是否有足夠庫存
                        if (largerSize !== size && remainingInventory[largerSize]?.allocatable >= requiredCount) {
                            adjustedSize = largerSize;
                            isAdjusted = true;
                            stats.pantsSizeAdjusted++;
                            
                            console.log(`因褲長(${student.pantsLength})與尺碼(${getSizeNumber(size)})差值>=3，學生 ${student.name} 的短衣尺寸從 ${size} 調整為 ${largerSize}`);
                        } else {
                            console.log(`雖然褲長(${student.pantsLength})與尺碼(${getSizeNumber(size)})差值>=3，但 ${largerSize} 尺寸庫存不足，學生 ${student.name} 維持原尺寸 ${size}`);
                        }
                    }
                    
                    // 分配調整後的尺寸
                    student[allocatedField] = adjustedSize;
                    student[specialField] = false;
                    
                    // 標記褲長調整
                    student.isShirtSizeAdjustedForPantsLength = isAdjusted;
                    if (isAdjusted) {
                        student.originalShirtSize = originalSize;
                        student.adjustmentReason = `褲長(${student.pantsLength})比尺碼(${getSizeNumber(originalSize)})大3以上`;
                    }
                    
                    decreaseInventory(remainingInventory, adjustedSize, requiredCount, inventoryType);
                    stats.allocated++;
                    
                    if (i === 0 && !isAdjusted) {
                        stats.exact++;
                    } else {
                        stats.different++;
                    }
                    allocated = true;
                    
                    // 清除失敗原因（如果之前有）
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    
                    const sizeDisplay = isAdjusted ? `${adjustedSize}↑` : adjustedSize;
                    console.log(`短衣：${student.name}(${student.gender}，座號：${student.number}，胸圍：${student.chest}，褲長：${student.pantsLength}) ，分配尺寸 ${sizeDisplay}, 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[adjustedSize]?.allocatable}`);
                    break;
                } else {
                    console.log(`學生 ${student.name}(${student.className}-${student.number}) 尺寸 ${size} 庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[size]?.allocatable} 件`);
                }
            }
            
            if (!allocated) {
                // 記錄未分配學生
                unallocatedStudents.push({
                    student,
                    requiredCount
                });
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '所有尺寸庫存不足';
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 暫時無法分配，所有尺寸庫存不足`);
            }
            
            // 更新可用尺寸列表
            availableSizes = availableSizes.filter(s => remainingInventory[s.size]?.allocatable > 0);
            if (availableSizes.length === 0) {
                console.warn('所有尺寸庫存都已用完，結束第一階段分配');
                
                // 將剩餘所有未分配的學生標記為失敗（庫存已用完）
                sortedStudents.forEach(({student}) => {
                    if (!student[allocatedField] && !student.allocationFailReason?.[inventoryType]) {
                        student.allocationFailReason = student.allocationFailReason || {};
                        student.allocationFailReason[inventoryType] = '所有尺寸庫存已用完';
                        stats.failed++;
                    }
                });
                break;
            }
        }
        
        // 第二階段：處理未分配的學生（嘗試特殊分配）
        console.log(`第二階段分配開始：處理 ${unallocatedStudents.length} 名未分配學生`);
        
        // 重新獲取所有尺寸，包括庫存為0的尺寸（可能某些學生需求數量較小，可以分配）
        let allSizes = Object.keys(remainingInventory)
            .filter(size => remainingInventory[size] && remainingInventory[size].allocatable > 0)
            .map(size => ({
                size,
                available: remainingInventory[size].allocatable
            }))
            .sort((a, b) => SIZES.indexOf(a.size) - SIZES.indexOf(b.size));
        
        if (allSizes.length > 0) {
            for (const {student, requiredCount} of unallocatedStudents) {
                // 尋找庫存最多的尺寸
                const bestSize = allSizes.reduce((prev, curr) => 
                    (curr.available > prev.available) ? curr : prev, allSizes[0]);
                
                if (bestSize && bestSize.available >= requiredCount) {
                    // 儲存原始尺寸
                    let originalSize = bestSize.size;
                    let adjustedSize = bestSize.size;
                    let isAdjusted = false;
                    
                    // 檢查是否因褲長需要調整尺寸
                    if (student.pantsLength && shouldAdjustShirtSizeForLongPants(student.pantsLength, bestSize.size)) {
                        const largerSize = getNextLargerSize(bestSize.size);
                        
                        // 檢查更大的尺寸是否有足夠庫存
                        if (largerSize !== bestSize.size && remainingInventory[largerSize]?.allocatable >= requiredCount) {
                            adjustedSize = largerSize;
                            isAdjusted = true;
                            stats.pantsSizeAdjusted++;
                            
                            console.log(`第二階段：因褲長(${student.pantsLength})與尺碼(${getSizeNumber(bestSize.size)})差值>=3，學生 ${student.name} 的短衣尺寸從 ${bestSize.size} 調整為 ${largerSize}`);
                        } else {
                            console.log(`第二階段：雖然褲長(${student.pantsLength})與尺碼(${getSizeNumber(bestSize.size)})差值>=3，但 ${largerSize} 尺寸庫存不足，學生 ${student.name} 維持原尺寸 ${bestSize.size}`);
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
                    
                    decreaseInventory(remainingInventory, adjustedSize, requiredCount, inventoryType);
                    stats.allocated++;
                    stats.special++;
                    
                    // 清除失敗原因
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    
                    const sizeDisplay = isAdjusted ? `${adjustedSize}↑` : adjustedSize;
                    console.log(`短衣（特殊分配）：${student.name}(${student.gender}，座號：${student.number}，胸圍：${student.chest}，褲長：${student.pantsLength}) ，分配尺寸 ${sizeDisplay}, 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[adjustedSize]?.allocatable}`);
                    
                    // 更新尺寸可用數量
                    allSizes = allSizes.map(s => s.size === adjustedSize ? 
                        { ...s, available: remainingInventory[s.size].allocatable } : s)
                        .filter(s => s.available > 0);
                    
                    if (allSizes.length === 0) {
                        console.warn('所有尺寸庫存都已用完，結束第二階段分配');
                        break;
                    }
            } else {
                stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：沒有足夠的庫存';
                    console.log(`學生 ${student.name}(${student.className}-${student.number}) 分配失敗：沒有足夠的庫存`);
                }
            }
        } else {
            // 如果沒有任何庫存可用於第二階段，將所有未分配學生標記為失敗
            unallocatedStudents.forEach(({student}) => {
                if (!student[allocatedField]) {
                    stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '第二階段分配失敗：沒有可用庫存';
                }
            });
        }
        
        // 確保統計數據包含所有學生
        let totalProcessed = stats.allocated + stats.failed;
        let totalStudentsNeedingAllocation = sortedStudents.filter(
            ({student}) => (student.shortSleeveShirtCount || 1) > 0
        ).length;
        
        // 如果有學生未被處理，將其標記為失敗
        if (totalProcessed < totalStudentsNeedingAllocation) {
            sortedStudents.forEach(({student}) => {
                if (!student[allocatedField] && !student.allocationFailReason?.[inventoryType] && 
                    (student.shortSleeveShirtCount || 1) > 0) {
                    stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配過程中被跳過';
                }
            });
        }
        
        // 更新庫存資料
        inventoryData[inventoryType] = remainingInventory;
        
        // 更新分配統計
        allocationStats[inventoryType] = stats;
        
        console.log(`分配結果：成功=${stats.allocated}, 完全符合=${stats.exact}, 不同尺寸=${stats.different}, 特殊分配=${stats.special}, 因褲長調整=${stats.pantsSizeAdjusted}, 失敗=${stats.failed}`);
        
        resolve(true);
    });
}

/**
 * 分配短袖褲子
 */
function allocateShortPants(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]}`);
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        // 獲取可用尺寸並排序（小到大）
        let availableSizes = getAvailableSizes(inventoryData, inventoryType);
        
        // 過濾掉可分配數為0的尺寸
        availableSizes = availableSizes.filter(size => size.available > 0);
        
        console.log(`可用尺寸：`, availableSizes.map(s => `${s.size}(${s.available})`).join(', '));
        
        if (availableSizes.length === 0) {
            console.warn(`沒有可用的 ${UNIFORM_TYPES[inventoryType]} 庫存`);
            resolve(false);
            return;
        }
        
        // 複製庫存資料用於分配
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        // 複製學生數組並根據腰圍排序
        const sortedStudents = [...sortedStudentData].map(student => {
            return {
                student,
                waist: student.waist || 0
            };
        }).sort((a, b) => a.waist - b.waist);

        console.log(`學生排序（前5名）：`, 
            sortedStudents.slice(0, 5).map(s => 
                `${s.student.name}(${s.student.className}-${s.student.number}): 腰圍=${s.waist}`
            )
        );
        
        // 用於記錄分配結果的統計資料
        const stats = {
            allocated: 0,
            exact: 0,
            different: 0,
            failed: 0,
            special: 0,
            pantsSizeAdjusted: 0 // 新增：因褲長調整尺寸的學生數
        };
        
        // 第一階段：按照腰圍排序分配
        console.log('分配開始：按腰圍順序分配');
        
        // 在分配開始前檢查短袖上衣分配失敗的學生
        const shortShirtFailedStudents = sortedStudents
            .filter(({student}) => 
                student.allocationFailReason && 
                student.allocationFailReason.shortSleeveShirt);
        
        if (shortShirtFailedStudents.length > 0) {
            console.log(`%c短褲分配開始前，檢測到 ${shortShirtFailedStudents.length} 名學生短袖上衣分配失敗：`, 'background: #f39c12; color: white; font-size: 12px; padding: 3px;');
            
            // 統計不同失敗原因的數量
            const failureReasons = {};
            shortShirtFailedStudents.forEach(({student}) => {
                const reason = student.allocationFailReason.shortSleeveShirt;
                failureReasons[reason] = (failureReasons[reason] || 0) + 1;
            });
            
            // 顯示每種失敗原因的數量
            for (const [reason, count] of Object.entries(failureReasons)) {
                console.log(`  - 失敗原因 "${reason}": ${count} 名學生`);
            }
            
            // 顯示部分學生的詳細信息
            console.log('失敗學生樣本(最多顯示5名):');
            shortShirtFailedStudents.slice(0, 5).forEach(({student}) => {
                console.log(`  - 學生 ${student.name}(${student.className}-${student.number}), 胸圍=${student.chest}, 腰圍=${student.waist}, 失敗原因: ${student.allocationFailReason.shortSleeveShirt}`);
            });
        }
        
        // 分配流程：按腰圍和胸圍順序分配
        for (const {student, waist} of sortedStudents) {
            // 跳過已分配的學生
            if (student.allocatedPantsSize) continue;
            
            // 檢查學生是否需要此類型的制服
            const requiredCount = student.shortSleevePantsCount || 1;
            
            // 如果學生不需要此類型制服（件數為0），標記為不需要
            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 不需要短褲（件數為0）`);
                continue;
            }
            
            // 檢查短袖上衣分配失敗情況
            const hasShirtFailure = student.allocationFailReason && student.allocationFailReason.shortSleeveShirt;
            if (hasShirtFailure) {
                console.log(`%c關注：學生 ${student.name}(${student.className}-${student.number}) 的短袖上衣分配失敗(${student.allocationFailReason.shortSleeveShirt})，但仍嘗試分配短褲%c`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;', '');
            }
            
            let allocated = false;
            
            // 尋找可用的尺寸
            for (let i = 0; i < availableSizes.length; i++) {
                const size = availableSizes[i].size;
                
                // 檢查剩餘庫存
                const remainingStock = remainingInventory[size]?.allocatable || 0;
                
                // 如果當前尺寸庫存不足，繼續查找
                if (remainingStock <= 0) {
                    if (hasShirtFailure) {
                        console.log(`學生 ${student.name}(${student.className}-${student.number}) 尺寸 ${size} 無庫存（注意：此學生短袖上衣分配失敗）`);
                    }
                    continue;
                }
                
                // 移除對上衣尺寸的階級差異檢查
                // 檢查庫存是否足夠
                let finalSize = size;
                
                // 檢查學生褲長與短褲尺碼的差距，如果差距>=3，則分配大一個尺寸
                if (student.pantsLength && (student.pantsLength - getSizeNumber(size) >= 3)) {
                    // 添加詳細日誌
                    console.log(`檢查學生 ${student.name} 的褲長調整: 褲長=${student.pantsLength}, 尺碼=${size}, 尺碼數字=${getSizeNumber(size)}, 差距=${student.pantsLength - getSizeNumber(size)}`);
                    
                    // 嘗試獲取大一號的尺寸
                    const largerSize = getNextLargerSize(size);
                    console.log(`  - 獲取大一號尺寸: ${size} -> ${largerSize}`);
                    
                    // 檢查大一號尺寸是否有足夠庫存
                    if (largerSize !== size && remainingInventory[largerSize]?.allocatable >= requiredCount) {
                        console.log(`  - 褲長調整成功: ${size} -> ${largerSize} (褲長=${student.pantsLength}, 差距=${student.pantsLength - getSizeNumber(size)})`);
                        finalSize = largerSize;
                        stats.pantsSizeAdjusted++;
                        student.isPantsLengthAdjusted = true; // 標記學生因褲長而調整了尺寸
                        i = availableSizes.findIndex(s => s.size === largerSize); // 更新索引
                    } else {
                        console.log(`  - 褲長調整失敗: 更大尺寸 ${largerSize} 庫存不足, 需要 ${requiredCount} 件, 剩餘 ${remainingInventory[largerSize]?.allocatable || 0} 件`);
                        console.log(`學生 ${student.name}(${student.className}-${student.number}) 褲長(${student.pantsLength})與尺寸${size}(${getSizeNumber(size)})差距過大，但更大尺寸${largerSize}無足夠庫存`);
                    }
                } else if (student.pantsLength) {
                    console.log(`學生 ${student.name} 無需褲長調整: 褲長=${student.pantsLength}, 尺碼=${size}, 尺碼數字=${getSizeNumber(size)}, 差距=${student.pantsLength - getSizeNumber(size)}`);
                }
                
                if (remainingInventory[finalSize]?.allocatable >= requiredCount) {
                    // 分配此尺寸
                    student.allocatedPantsSize = finalSize;
                    student.isSpecialPantsAllocation = false;
                    decreaseInventory(remainingInventory, finalSize, requiredCount, inventoryType);
                    stats.allocated++;
                    if (i === 0) {
                        stats.exact++;
                    } else {
                        stats.different++;
                    }
                    allocated = true;
                    
                    // 清除失敗原因（如果之前有）
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    
                    console.log(`短褲：${student.name}(${student.gender}，座號：${student.number}，腰圍：${waist}) ，分配尺寸 ${finalSize}, 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[finalSize]?.allocatable}${student.isPantsLengthAdjusted ? ', 因褲長調整' : ''}`);
                    break;
                } else {
                    console.log(`學生 ${student.name}(${student.className}-${student.number}) 尺寸 ${finalSize} 庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[finalSize]?.allocatable} 件`);
                }
            }
            
            // 如果無法分配，記錄未分配學生
            if (!allocated) {
                // 添加分配失敗原因
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '沒有合適的尺寸或庫存不足';
                stats.failed++;
                
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 分配短褲失敗：沒有合適的尺寸或庫存不足`);
                if (hasShirtFailure) {
                    console.log(`%c關注：學生 ${student.name}(${student.className}-${student.number}) 短褲分配失敗，且此學生短袖上衣也分配失敗%c`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;', '');
                }
            }
        }
        
        // 在分配完成後立即添加詳細的統計日誌
        console.log('--------------------------------------------------');
        console.log(`${UNIFORM_TYPES[inventoryType]}分配結果詳細統計：`);
        console.log(`總分配成功: ${stats.allocated} 人`);
        console.log(`完全符合尺寸: ${stats.exact} 人`);
        console.log(`分配不同尺寸: ${stats.different} 人`);
        console.log(`因褲長調整尺寸: ${stats.pantsSizeAdjusted} 人`);
        console.log(`分配失敗: ${stats.failed} 人`);
        
        // 檢查所有類型的分配失敗
        const allFailedStudents = sortedStudentData.filter(student => 
            !student.allocatedPantsSize);
        
        if (allFailedStudents.length > 0) {
            console.log(`%c總共 ${allFailedStudents.length} 名學生未獲得 ${UNIFORM_TYPES[inventoryType]} 分配：`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
            
            // 按照失敗原因進行分類
            const failureGroups = {};
            
            allFailedStudents.forEach(student => {
                const reason = student.allocationFailReason?.[inventoryType] || '未知原因';
                failureGroups[reason] = failureGroups[reason] || [];
                failureGroups[reason].push(student);
            });
            
            // 顯示每種失敗原因的學生數量和詳情
            for (const [reason, students] of Object.entries(failureGroups)) {
                console.log(`失敗原因 "${reason}": ${students.length} 名學生`);
                students.forEach(student => {
                    console.log(`  - 學生 ${student.name}(${student.className}-${student.number}), 腰圍: ${student.waist}, 胸圍: ${student.chest}`);
                    console.log(`    * 短袖上衣分配狀態: ${student.allocatedShirtSize ? `已分配尺寸 ${student.allocatedShirtSize}` : '未分配'}`);
                    if (!student.allocatedShirtSize && student.allocationFailReason?.shortSleeveShirt) {
                        console.log(`    * 短袖上衣分配失敗原因: ${student.allocationFailReason.shortSleeveShirt}`);
                    }
                });
            }
        }
        
        // 輸出因褲長調整尺寸的學生列表
        const lengthAdjustedStudents = sortedStudentData.filter(student => student.isPantsLengthAdjusted);
        if (lengthAdjustedStudents.length > 0) {
            console.log(`%c因褲長調整尺寸的學生（共${lengthAdjustedStudents.length}名）：`, 'background: #2ecc71; color: white; font-size: 12px; padding: 3px;');
            lengthAdjustedStudents.forEach(student => {
                console.log(`  - 學生 ${student.name}(${student.className}-${student.number}), 腰圍: ${student.waist}, 褲長: ${student.pantsLength}, 分配尺寸: ${student.allocatedPantsSize}`);
            });
        }
        
        // 顯示最終庫存狀態
        console.log('最終庫存狀態：');
        Object.entries(remainingInventory).forEach(([size, data]) => {
            console.log(`尺寸 ${size}: 可分配=${data.allocatable}, 已分配=${data.allocated}, 預留=${data.reserved}`);
        });
        console.log('--------------------------------------------------');
        
        // 更新庫存資料
        inventoryData[inventoryType] = remainingInventory;
        
        console.log(`分配結果：成功=${stats.allocated}, 完全符合=${stats.exact}, 不同尺寸=${stats.different}, 因褲長調整=${stats.pantsSizeAdjusted}, 失敗=${stats.failed}`);
        
        resolve(true);
    });
}

/**
 * 分配長袖上衣
 */
function allocateLongShirts(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]}`);
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        // 獲取可用尺寸並排序（小到大）
        let availableSizes = getAvailableSizes(inventoryData, inventoryType);
        
        // 過濾掉可分配數為0的尺寸
        availableSizes = availableSizes.filter(size => size.available > 0);
        
        console.log(`可用尺寸：`, availableSizes.map(s => `${s.size}(${s.available})`).join(', '));
        
        if (availableSizes.length === 0) {
            console.warn(`沒有可用的 ${UNIFORM_TYPES[inventoryType]} 庫存`);
            resolve(false);
            return;
        }

        // 複製庫存資料用於分配
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        // 複製學生數組並根據有效胸圍排序
        const sortedStudents = [...sortedStudentData].map(student => ({
            student,
            effectiveChest: calculateEffectiveChest(student)
        })).sort((a, b) => a.effectiveChest - b.effectiveChest);

        console.log(`學生排序（前5名）：`, 
            sortedStudents.slice(0, 5).map(s => 
                `${s.student.name}(${s.student.className}-${s.student.number}): 有效胸圍=${s.effectiveChest}`
            )
        );
        
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
        console.log('第一階段分配開始：按胸圍順序分配');
        
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
            
            // 顯示部分學生詳情
            console.log('失敗學生樣本(最多顯示5名):');
            priorFailedStudents.slice(0, 5).forEach(({student}) => {
                const reasons = [];
                if (student.allocationFailReason.shortSleeveShirt) {
                    reasons.push(`短袖上衣: ${student.allocationFailReason.shortSleeveShirt}`);
                }
                if (student.allocationFailReason.shortSleevePants) {
                    reasons.push(`短褲: ${student.allocationFailReason.shortSleevePants}`);
                }
                console.log(`  - 學生 ${student.name}(${student.className}-${student.number}), 胸圍=${student.chest}, 先前失敗原因: ${reasons.join(', ')}`);
            });
        }
        
        // 第一階段：按照排序分配
        console.log('第一階段分配開始：按胸圍順序分配');
        for (const {student} of sortedStudents) {
            // 跳過已分配的學生
            if (student[allocatedField]) continue;
            
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
                console.log(`%c關注：學生 ${student.name}(${student.className}-${student.number}) 的先前分配有失敗記錄，但仍嘗試分配長衣：%c`, 'background: #e67e22; color: white; font-size: 12px; padding: 3px;', '');
                priorReasons.forEach(reason => console.log(`  - ${reason}`));
            }
            
            const requiredCount = student.longSleeveShirtCount || 1;
            
            // 如果學生不需要此類型制服（件數為0），標記為不需要
            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 不需要長衣（件數為0）`);
                continue;
            }
            
            let allocated = false;
            
            // 尋找可用的尺寸
            for (let i = 0; i < availableSizes.length; i++) {
                const size = availableSizes[i].size;
                
                // 檢查庫存是否足夠
                if (remainingInventory[size]?.allocatable < requiredCount) {
                    if (hasPriorFailure) {
                        console.log(`學生 ${student.name}(${student.className}-${student.number}) 尺寸 ${size} 庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[size]?.allocatable} 件（注意：此學生先前分配有失敗記錄）`);
                    }
                    continue;
                }
                
                // 分配此尺寸
                student[allocatedField] = size;
                student[specialField] = false;
                decreaseInventory(remainingInventory, size, requiredCount, inventoryType);
                stats.allocated++;
                if (i === 0) {
                    stats.exact++;
                } else {
                    stats.different++;
                }
                allocated = true;
                
                if (hasPriorFailure) {
                    console.log(`%c特別關注：儘管先前分配有失敗記錄，學生 ${student.name}(${student.className}-${student.number}) 仍然分配到長衣尺寸 ${size}%c`, 'background: #2ecc71; color: white; font-size: 12px; padding: 3px;', '');
                } else {
                    // 計算調整後的胸圍（使用學校配置的女生胸圍調整值）
                    let adjustedChest = student.chest || 0;
                    if (student.gender === '女') {
                        adjustedChest += getFemaleChestAdjustment();
                    }
                    // 獲取調整值，以便在日誌中顯示
                    const adjustment = getFemaleChestAdjustment();
                    console.log(`長衣：${student.name}（${student.gender}，座號：${student.number}，胸圍：${student.chest}${student.gender === '女' ? `(已調整${adjustment}cm=${adjustedChest})` : ''}）， 分配尺寸 ${size}, 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[size]?.allocatable}`);
                }
                break;
            }
            
            // 如果無法分配，記錄未分配學生及其適合尺寸
            if (!allocated) {
                unallocatedStudents.push({
                    student,
                    suitableSize: availableSizes.length > 0 ? availableSizes[0].size : null,
                    requiredCount
                });
                
                if (hasPriorFailure) {
                    console.log(`%c關注：學生 ${student.name}(${student.className}-${student.number}) 第一階段長衣分配失敗，且先前已有分配失敗記錄%c`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;', '');
                }
            }
            
            // 更新可用尺寸列表
            availableSizes = availableSizes.filter(s => remainingInventory[s.size]?.allocatable > 0);
            if (availableSizes.length === 0) {
                console.warn('所有尺寸庫存都已用完，結束第一階段分配');
                break;
            }
        }
        
        // 第二階段：處理未分配的學生（嘗試特殊分配）
        console.log(`第二階段分配開始：處理 ${unallocatedStudents.length} 名未分配學生`);
        
        // 重新獲取所有尺寸，包括庫存為0的尺寸（可能某些學生需求數量較小，可以分配）
        let allSizes = Object.keys(remainingInventory)
            .filter(size => remainingInventory[size] && remainingInventory[size].allocatable > 0)
            .map(size => ({
                size,
                available: remainingInventory[size].allocatable
            }))
            .sort((a, b) => SIZES.indexOf(a.size) - SIZES.indexOf(b.size));
        
        if (allSizes.length > 0) {
            for (const {student, requiredCount} of unallocatedStudents) {
                // 尋找庫存最多的尺寸
                const bestSize = allSizes.reduce((prev, curr) => 
                    (curr.available > prev.available) ? curr : prev, allSizes[0]);
                
                if (bestSize && bestSize.available >= requiredCount) {
                    // 分配庫存最多的尺寸
                    student[allocatedField] = bestSize.size;
                    student[specialField] = true; // 標記為特殊分配
                    decreaseInventory(remainingInventory, bestSize.size, requiredCount, inventoryType);
                stats.allocated++;
                stats.special++;
                    
                    // 清除失敗原因
                    if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                        delete student.allocationFailReason[inventoryType];
                    }
                    
                    console.log(`學生 ${student.name}(${student.className}-${student.number}) 特殊分配尺寸 ${bestSize.size}（庫存最多）`);

                    // 修改為以下格式
                    // 計算調整後的胸圍（使用學校配置的女生胸圍調整值）
                    let adjustedChest = student.chest || 0;
                    if (student.gender === '女') {
                        adjustedChest += getFemaleChestAdjustment();
                    }
                    // 獲取調整值，以便在日誌中顯示
                    const adjustment = getFemaleChestAdjustment();
                    console.log(`長衣：${student.name}（${student.gender}，座號：${student.number}，胸圍：${student.chest}${student.gender === '女' ? `(已調整${adjustment}cm=${adjustedChest})` : ''}）， 特殊分配尺寸 ${bestSize.size}, 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[bestSize.size]?.allocatable}`);
                    
                    // 更新尺寸可用數量
                    allSizes = allSizes.map(s => s.size === bestSize.size ? 
                        { ...s, available: remainingInventory[s.size].allocatable } : s)
                        .filter(s => s.available > 0);
                    
                    if (allSizes.length === 0) {
                        console.warn('所有尺寸庫存都已用完，結束第二階段分配');
                        break;
                    }
            } else {
                stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '分配失敗：沒有足夠的庫存';
                    console.log(`學生 ${student.name}(${student.className}-${student.number}) 分配失敗：沒有足夠的庫存`);
                }
            }
        } else {
            // 如果沒有任何庫存可用於第二階段，將所有未分配學生標記為失敗
            unallocatedStudents.forEach(({student}) => {
                if (!student[allocatedField]) {
                    stats.failed++;
                    student.allocationFailReason = student.allocationFailReason || {};
                    student.allocationFailReason[inventoryType] = '第二階段分配失敗：沒有可用庫存';
                }
            });
        }
        
        // 更新庫存資料
        inventoryData[inventoryType] = remainingInventory;
        
        console.log(`分配結果：成功=${stats.allocated}, 完全符合=${stats.exact}, 不同尺寸=${stats.different}, 特殊分配=${stats.special}, 因褲長調整=${stats.pantsSizeAdjusted}, 失敗=${stats.failed}`);
        
        resolve(true);
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
    // 獲取尺碼對應的實際厘米數
    const sizeLengthInCm = getPantsLengthInCm(pantsSize);
    
    // 學生可接受的最小尺碼為其褲長減去2厘米
    const minAcceptableLength = studentPantsLength - 2;
    
    // 褲子實際長度不能小於學生可接受的最小長度
    return sizeLengthInCm >= minAcceptableLength;
}

/**
 * 分配長袖褲子
 */
function allocateLongPants(inventoryType, allocatedField, specialField) {
    return new Promise((resolve) => {
        console.log(`開始分配 ${UNIFORM_TYPES[inventoryType]}`);
        
        // 確保有庫存資料
        if (!inventoryData[inventoryType]) {
            console.warn(`沒有 ${UNIFORM_TYPES[inventoryType]} 庫存資料`);
            resolve(false);
            return;
        }
        
        // 獲取可用尺寸並排序（小到大）
        let availableSizes = getAvailableSizes(inventoryData, inventoryType);
        
        // 過濾掉可分配數為0的尺寸
        availableSizes = availableSizes.filter(size => size.available > 0);
        
        console.log(`可用尺寸：`, availableSizes.map(s => `${s.size}(${s.available})`).join(', '));
        
        if (availableSizes.length === 0) {
            console.warn(`沒有可用的 ${UNIFORM_TYPES[inventoryType]} 庫存`);
            resolve(false);
            return;
        }
        
        // 複製庫存資料用於分配
        const remainingInventory = JSON.parse(JSON.stringify(inventoryData[inventoryType]));
        
        // 複製學生數組並根據腰圍和褲長排序
        const sortedStudents = [...sortedStudentData].map(student => ({
            student,
            waist: student.waist || 0,
            pantsLength: student.pantsLength || 0
        })).sort((a, b) => {
            if (a.waist === b.waist) {
                return a.pantsLength - b.pantsLength;
            }
            return a.waist - b.waist;
        });

        console.log(`學生排序（前5名）：`, 
            sortedStudents.slice(0, 5).map(s => 
                `${s.student.name}(${s.student.className}-${s.student.number}): 腰圍=${s.waist}, 褲長=${s.pantsLength}`
            )
        );
        
        // 用於記錄分配結果的統計資料
        const stats = {
            allocated: 0,
            exact: 0,
            different: 0,
            failed: 0,
            special: 0,
            pantsSizeAdjusted: 0 // 新增：因褲長調整尺寸的學生數
        };
        
        // 記錄未分配的學生及其適合尺寸
        const unallocatedStudents = [];
        
        // 第一階段：按照腰圍和褲長排序分配
        console.log('第一階段分配開始：按腰圍和褲長順序分配');
        
        // 處理第一階段分配
        for (const {student, waist, pantsLength} of sortedStudents) {
            // 檢查是否已經分配過 - 避免重複分配
            if (student[allocatedField]) continue;
            
            const requiredCount = student.longSleevePantsCount || 1;
            
            // 如果學生不需要此類型制服（件數為0），標記為不需要並跳過
            if (requiredCount <= 0) {
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = '學生不需要此制服（件數為0）';
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 不需要長褲（件數為0）`);
                continue;
            }
            
            // 日誌記錄
            console.log(`尋找學生 ${student.name}(${student.className}-${student.number}) 的長褲尺寸，腰圍=${waist}，褲長=${pantsLength}`);
            
            let allocated = false;
            let suitableSize = null;
            
            // 嘗試每個可用的尺寸
            for (let i = 0; i < availableSizes.length; i++) {
                const size = availableSizes[i].size;
                
                // 檢查庫存是否足夠
                if (remainingInventory[size]?.allocatable < requiredCount) {
                    console.log(`學生 ${student.name}(${student.className}-${student.number}) 尺寸 ${size} 庫存不足，需要 ${requiredCount} 件，但只剩 ${remainingInventory[size]?.allocatable} 件`);
                    continue;
                }
                
                // 使用新的褲長適合性檢查
                if (!isPantsLengthAcceptable(pantsLength, size)) {
                    // 記錄適合的尺寸，但不分配（褲長限制）
                    if (!suitableSize) {
                        suitableSize = size;
                        console.log(`學生 ${student.name}(${student.className}-${student.number}) 尺寸 ${size} 不符合褲長要求 ${pantsLength}cm（尺寸長度為 ${getPantsLengthInCm(size)}cm，學生最小可接受長度為 ${pantsLength - 2}cm）`);
                    }
                    continue;
                }
                
                // 分配此尺寸
                student[allocatedField] = size;
                student[specialField] = false;
                decreaseInventory(remainingInventory, size, requiredCount, inventoryType);
                stats.allocated++;
                if (i === 0) {
                    stats.exact++;
                } else {
                    stats.different++;
                }
                allocated = true;
                
                // 清除失敗原因（如果之前有）
                if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                    delete student.allocationFailReason[inventoryType];
                }
                
                console.log(`長褲：${student.name}(${student.gender}，座號：${student.number}，腰圍=${waist}，褲長=${pantsLength}) ，分配尺寸 ${size}(${getPantsLengthInCm(size)}cm), 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[size]?.allocatable}`);
                break;
            }
            
            // 如果無法分配，記錄未分配學生及其適合尺寸
            if (!allocated) {
                unallocatedStudents.push({
                    student,
                    suitableSize: suitableSize || (availableSizes.length > 0 ? availableSizes[0].size : null),
                    requiredCount,
                    pantsLength
                });
                
                // 更新失敗原因，明確註明是褲長需求超出可用尺碼範圍
                student.allocationFailReason = student.allocationFailReason || {};
                student.allocationFailReason[inventoryType] = `褲長需求(${pantsLength}cm)超出可用尺碼範圍(${getPantsLengthInCm(availableSizes[availableSizes.length-1].size)}cm)`;
                
                console.log(`學生 ${student.name}(${student.className}-${student.number}) 第一階段長褲分配失敗，褲長需求 ${pantsLength}cm，最小可接受長度 ${pantsLength - 2}cm，最大可用尺碼 ${availableSizes[availableSizes.length-1].size}(${getPantsLengthInCm(availableSizes[availableSizes.length-1].size)}cm)`);
            }
            
            // 更新可用尺寸列表 - 只移除庫存為0的尺寸，保留所有有庫存的尺寸
            availableSizes = availableSizes.filter(s => remainingInventory[s.size]?.allocatable > 0);
            if (availableSizes.length === 0) {
                console.warn('所有尺寸庫存都已用完，結束第一階段分配');
                break;
            }
        }
        
        // 第二階段：處理未分配的學生
        console.log(`第二階段分配開始：處理 ${unallocatedStudents.length} 名未分配學生`);
        
        // 記錄第一階段結束時的庫存狀態
        console.log(`第一階段結束後庫存狀態：`);
        Object.entries(remainingInventory).forEach(([size, data]) => {
            console.log(`尺寸 ${size}(${getPantsLengthInCm(size)}cm): 可分配=${data.allocatable}, 已分配=${data.allocated}, 預留=${data.reserved}`);
        });
        
        for (const {student, suitableSize, requiredCount, pantsLength} of unallocatedStudents) {
            // 查找適合尺寸及其相鄰尺寸的庫存
            const suitableIndex = SIZES.indexOf(suitableSize);
            const sizes = [];
            
            // 嘗試所有可能適合的尺寸（包括那些可能不完全符合褲長要求的）
            for (const size of Object.keys(remainingInventory)) {
                if (remainingInventory[size] && remainingInventory[size].allocatable >= requiredCount) {
                    sizes.push(size);
                }
            }
            
            // 按照尺寸大小排序
            sizes.sort((a, b) => {
                const aLength = getPantsLengthInCm(a);
                const bLength = getPantsLengthInCm(b);
                return bLength - aLength; // 優先選擇較大的尺寸（降序排列）
            });
            
            console.log(`嘗試為學生 ${student.name}(${student.className}-${student.number}) 進行特殊分配，褲長要求 ${pantsLength}cm`);
            console.log(`  - 考慮的尺寸: ${sizes.map(s => `${s}(${getPantsLengthInCm(s)}cm)`).join(', ')}`);
            
            // 找出最佳尺寸（優先選擇符合最小褲長要求的最小尺寸）
            let bestSize = null;
            
            // 先嘗試符合褲長要求的尺寸
            for (const size of sizes) {
                if (isPantsLengthAcceptable(pantsLength, size) && 
                    remainingInventory[size]?.allocatable >= requiredCount) {
                    bestSize = size;
                    console.log(`  - 選擇尺寸 ${size}(${getPantsLengthInCm(size)}cm)，符合褲長要求，庫存為 ${remainingInventory[size]?.allocatable}`);
                    break;
                }
            }
            
            // 修改這裡：不再分配不符合褲長要求的尺寸
            if (bestSize) {
                // 分配最佳替代尺寸
                student[allocatedField] = bestSize;
                student[specialField] = true; // 標記為特殊分配
                decreaseInventory(remainingInventory, bestSize, requiredCount, inventoryType);
                stats.allocated++;
                stats.special++;
                
                // 清除失敗原因（如果之前有）
                if (student.allocationFailReason && student.allocationFailReason[inventoryType]) {
                    delete student.allocationFailReason[inventoryType];
                }
                
                console.log(`長褲特殊分配：${student.name}（${student.gender}，座號：${student.number}，腰圍=${student.waist}，褲長=${pantsLength}）， 分配尺寸 ${bestSize}(${getPantsLengthInCm(bestSize)}cm), 需求數量 ${requiredCount}, 剩餘庫存 ${remainingInventory[bestSize]?.allocatable}`);
            } else {
                // 檢查是否有可能的尺寸（不符合褲長要求）
                if (sizes.length > 0) {
                    const largestSize = sizes[0]; // 已排序，第一個是最大的
                    console.log(`  - 無符合褲長要求的尺寸，有 ${sizes.length} 個不符合要求的尺寸可用`);
                    console.log(`  - 最大可用尺寸 ${largestSize}(${getPantsLengthInCm(largestSize)}cm) 小於學生褲長要求 ${pantsLength}cm`);
                    console.log(`  - 標記為分配失敗，不進行特殊分配`);
                }
                
                // 無法分配
                stats.failed++;
                student.allocationFailReason = student.allocationFailReason || {};
                // 明確標示分配失敗原因
                student.allocationFailReason[inventoryType] = `分配失敗：褲長限制`;
                
                console.log(`%c學生 ${student.name}(${student.className}-${student.number}) 分配失敗，褲長要求 ${pantsLength}cm 超出可用尺寸範圍（最大可用尺寸只有 ${sizes.length > 0 ? getPantsLengthInCm(sizes[0]) : 0}cm）%c`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;', '');
            }
        }
        
        // 在分配完成後立即添加詳細的統計日誌
        console.log('--------------------------------------------------');
        console.log(`${UNIFORM_TYPES[inventoryType]}分配結果詳細統計：`);
        console.log(`總分配成功: ${stats.allocated} 人`);
        console.log(`完全符合尺寸: ${stats.exact} 人`);
        console.log(`分配不同尺寸: ${stats.different} 人`);
        console.log(`特殊分配: ${stats.special} 人`);
        console.log(`分配失敗: ${stats.failed} 人`);
        
        // 檢查所有類型的分配失敗
        const allFailedStudents = sortedStudentData.filter(student => 
            !student[allocatedField]);
        
        if (allFailedStudents.length > 0) {
            console.log(`%c總共 ${allFailedStudents.length} 名學生未獲得 ${UNIFORM_TYPES[inventoryType]} 分配：`, 'background: #e74c3c; color: white; font-size: 12px; padding: 3px;');
            
            // 按照失敗原因進行分類
            const failureGroups = {};
            
            allFailedStudents.forEach(student => {
                const reason = student.allocationFailReason?.[inventoryType] || '未知原因';
                failureGroups[reason] = failureGroups[reason] || [];
                failureGroups[reason].push(student);
            });
            
            // 顯示每種失敗原因的學生數量和詳情
            for (const [reason, students] of Object.entries(failureGroups)) {
                console.log(`%c失敗原因 "${reason}": ${students.length} 名學生%c`, 'background: #f39c12; color: white; font-size: 12px; padding: 3px;', '');
                students.forEach(student => {
                    console.log(`  - 學生 ${student.name}(${student.className}-${student.number}), 腰圍: ${student.waist}, 褲長: ${student.pantsLength}`);
                });
            }
        }
        
        // 輸出因褲長調整尺寸的學生列表
        const lengthAdjustedStudents = sortedStudentData.filter(student => student.isPantsLengthAdjusted);
        if (lengthAdjustedStudents.length > 0) {
            console.log(`%c因褲長調整尺寸的學生（共${lengthAdjustedStudents.length}名）：`, 'background: #2ecc71; color: white; font-size: 12px; padding: 3px;');
            lengthAdjustedStudents.forEach(student => {
                console.log(`  - 學生 ${student.name}(${student.className}-${student.number}), 腰圍: ${student.waist}, 褲長: ${student.pantsLength}, 分配尺寸: ${student.allocatedLongPantsSize}`);
            });
        }
        
        // 顯示最終庫存狀態
        console.log('最終庫存狀態：');
        Object.entries(remainingInventory).forEach(([size, data]) => {
            console.log(`尺寸 ${size}: 可分配=${data.allocatable}, 已分配=${data.allocated}, 預留=${data.reserved}`);
        });
        console.log('--------------------------------------------------');
        
        // 更新庫存資料
        inventoryData[inventoryType] = remainingInventory;
        
        console.log(`分配結果：成功=${stats.allocated}, 完全符合=${stats.exact}, 不同尺寸=${stats.different}, 特殊分配=${stats.special}, 因褲長調整=${stats.pantsSizeAdjusted}, 失敗=${stats.failed}`);
        
        resolve(true);
    });
}

/**
 * 更新分配結果頁面
 */
export function updateAllocationResults() {
    console.log('開始更新分配結果頁面');
    
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
            </td>
            <td class="count-column">${student.allocatedShirtSize ? (student.shortSleeveShirtCount || 1) : '-'}</td>
            <td>
                ${formattedPantsSize}${student.isPantsLengthAdjusted ? '<sup>↑</sup>' : ''}
                ${shortPantsFailReason ? `<div class="failure-reason">${shortPantsFailReason}</div>` : ''}
                ${student.isPantsLengthAdjusted ? `<div class="adjustment-reason text-info" style="font-size: 0.85em;">因褲長調整</div>` : ''}
            </td>
            <td class="count-column">${student.allocatedPantsSize ? (student.shortSleevePantsCount || 1) : '-'}</td>
            <td>
                ${formattedLongShirtSize}
                ${longShirtFailReason ? `<div class="failure-reason">${longShirtFailReason}</div>` : ''}
            </td>
            <td class="count-column">${student.allocatedLongShirtSize ? (student.longSleeveShirtCount || 1) : '-'}</td>
            <td>
                ${formattedLongPantsSize}
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
        totalStudents: sortedStudentData.length,
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
    
    for (const size in inventoryData[inventoryType]) {
        if (!inventoryData[inventoryType].hasOwnProperty(size)) continue;
        
        const invData = inventoryData[inventoryType][size];
        const available = invData.allocatable || 0;
        
        if (available > 0) {
            availableSizes.push({
                size: size,
                available: available
            });
        }
    }
    
    // 按照尺寸索引排序
    availableSizes.sort((a, b) => {
        return SIZES.indexOf(a.size) - SIZES.indexOf(b.size);
    });
    
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
    
    // 尋找庫存最多的尺寸
    for (const size of suitableSizes) {
        const available = inventory[size]?.allocatable || 0;
        if (available > maxAvailable) {
            maxAvailable = available;
            bestSize = size;
        }
    }
    
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
        
        let shortPantsSize = student.allocatedPantsSize ? formatSize(student.allocatedPantsSize) : '-';
        let longShirtSize = student.allocatedLongShirtSize ? formatSize(student.allocatedLongShirtSize) : '-';
        let longPantsSize = student.allocatedLongPantsSize ? formatSize(student.allocatedLongPantsSize) : '-';
        
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