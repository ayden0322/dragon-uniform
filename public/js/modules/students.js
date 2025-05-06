// 學生資料相關功能模組
import { saveToLocalStorage, loadFromLocalStorage, showAlert, downloadCSV, downloadExcel } from './utils.js';
import { SIZES, UNIFORM_TYPES, formatSize, getFemaleChestAdjustment } from './config.js';
import { updateAllocationRatios } from './ui.js';

// 學生資料
export let studentData = [];
export let sortedStudentData = [];

// 需求資料
export let demandData = {
    shortSleeveShirt: { totalDemand: 0, sizeCount: {} },
    shortSleevePants: { totalDemand: 0, sizeCount: {} },
    longSleeveShirt: { totalDemand: 0, sizeCount: {} },
    longSleevePants: { totalDemand: 0, sizeCount: {} }
};

/**
 * 初始化學生資料表格
 */
export function initStudentTable() {
    const studentTable = document.getElementById('studentTable');
    if (!studentTable) return;
    
    const tbody = studentTable.querySelector('tbody');
    
    // 清空表格
    tbody.innerHTML = '';
    
    // 建立表格內容
    if (studentData.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="12" class="text-center">無資料，請匯入或手動新增</td>';
        tbody.appendChild(row);
    } else {
        studentData.forEach((student, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${index + 1}</td>
                <td>${student.class || ''}</td>
                <td>${student.number || ''}</td>
                <td>${student.name || ''}</td>
                <td>${student.gender || ''}</td>
                <td>${student.chest || ''}</td>
                <td>${student.waist || ''}</td>
                <td>${student.pantsLength || ''}</td>
                <td>${student.shortSleeveShirtCount || ''}</td>
                <td>${student.shortSleevePantsCount || ''}</td>
                <td>${student.longSleeveShirtCount || ''}</td>
                <td>${student.longSleevePantsCount || ''}</td>
            `;
            tbody.appendChild(row);
        });
    }
}

/**
 * 清除學生資料
 */
export function clearStudentData() {
    // 確認是否要清除資料
    if (confirm('確定要清除所有學生資料嗎？此操作無法復原！')) {
        // 清空學生資料
        studentData = [];
        sortedStudentData = [];
        
        // 重置需求資料
        demandData = {
            shortSleeveShirt: { totalDemand: 0, sizeCount: {} },
            shortSleevePants: { totalDemand: 0, sizeCount: {} },
            longSleeveShirt: { totalDemand: 0, sizeCount: {} },
            longSleevePants: { totalDemand: 0, sizeCount: {} }
        };
        
        // 儲存到本地儲存
        saveToLocalStorage('studentData', studentData);
        saveToLocalStorage('demandData', demandData);
        
        // 更新表格
        initStudentTable();
        
        // 提示用戶
        showAlert('已清除所有學生資料', 'success');
    }
}

/**
 * 處理Excel檔案匯入
 * @param {File} file - Excel檔案
 * @returns {Promise} - 成功時解析為true，失敗時拒絕並返回錯誤
 */
export function handleExcelImport(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                
                // 使用SheetJS解析Excel資料
                const workbook = XLSX.read(data, {type: 'binary'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // 轉換為JSON格式
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
                
                // 解析資料
                if (jsonData.length < 2) {
                    showAlert('Excel檔案內容不足', 'danger');
                    resolve(false);
                    return;
                }
                
                // 取得標頭
                const headers = jsonData[0];
                
                // 解析學生資料
                const newStudentData = [];
                const incompleteStudents = [];
                
                for (let i = 1; i < jsonData.length; i++) {
                    const row = jsonData[i];
                    
                    // 跳過空行
                    if (row.length === 0 || !row[0]) continue;
                    
                    const student = {};
                    let isComplete = true;
                    
                    // 依標頭配對資料
                    headers.forEach((header, index) => {
                        // 標準化標頭
                        const normalizedHeader = normalizeCellData(header);
                        
                        // 配對資料欄位
                        if (index < row.length) {
                            const cellValue = normalizeCellData(row[index]);
                            
                            switch (normalizedHeader) {
                                case '班級':
                                case '班':
                                    student.class = cellValue;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '座號':
                                case '號碼':
                                    student.number = cellValue;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '姓名':
                                case '學生姓名':
                                    student.name = cellValue;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '性別':
                                    student.gender = cellValue;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '胸圍':
                                case '胸':
                                    student.chest = parseFloat(cellValue) || 0;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '腰圍':
                                case '腰':
                                    student.waist = parseFloat(cellValue) || 0;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '褲長':
                                    student.pantsLength = parseFloat(cellValue) || 0;
                                    if (!cellValue) isComplete = false;
                                    break;
                                case '短衣件數':
                                case '短袖上衣件數':
                                    // 確保件數是非負整數
                                    let shortShirtCount = parseInt(cellValue);
                                    student.shortSleeveShirtCount = (shortShirtCount >= 0) ? shortShirtCount : 0;
                                    if (cellValue === '') isComplete = false;
                                    break;
                                case '短褲件數':
                                case '短袖褲子件數':
                                    let shortPantsCount = parseInt(cellValue);
                                    student.shortSleevePantsCount = (shortPantsCount >= 0) ? shortPantsCount : 0;
                                    if (cellValue === '') isComplete = false;
                                    break;
                                case '長衣件數':
                                case '長袖上衣件數':
                                    let longShirtCount = parseInt(cellValue);
                                    student.longSleeveShirtCount = (longShirtCount >= 0) ? longShirtCount : 0;
                                    if (cellValue === '') isComplete = false;
                                    break;
                                case '長褲件數':
                                case '長袖褲子件數':
                                    let longPantsCount = parseInt(cellValue);
                                    student.longSleevePantsCount = (longPantsCount >= 0) ? longPantsCount : 0;
                                    if (cellValue === '') isComplete = false;
                                    break;
                                default:
                                    break;
                            }
                        } else {
                            // 如果缺少欄位，則標記為不完整
                            isComplete = false;
                        }
                    });
                    
                    // 只有在資料完整時才添加學生資料
                    if (isComplete) {
                        // 確保所有件數欄位都有值，沒有的話默認為0
                        student.shortSleeveShirtCount = student.shortSleeveShirtCount || 0;
                        student.shortSleevePantsCount = student.shortSleevePantsCount || 0;
                        student.longSleeveShirtCount = student.longSleeveShirtCount || 0;
                        student.longSleevePantsCount = student.longSleevePantsCount || 0;
                        
                        newStudentData.push(student);
                    } else {
                        // 記錄不完整的學生資料
                        incompleteStudents.push({
                            row: i + 1,
                            data: row,
                            student: student
                        });
                    }
                }
                
                // 處理不完整的學生資料
                if (incompleteStudents.length > 0) {
                    // 整理不完整的學生資料以顯示
                    const incompleteMessage = incompleteStudents.map(item => {
                        let details = `第 ${item.row} 行: `;
                        if (item.student.class) details += `班級=${item.student.class}, `;
                        if (item.student.number) details += `號碼=${item.student.number}, `;
                        if (item.student.name) details += `姓名=${item.student.name}`;
                        return details;
                    }).join('\n');
                    
                    // 顯示不完整資料的通知
                    alert(`發現 ${incompleteStudents.length} 筆資料不完整，這些資料將被忽略：\n\n${incompleteMessage}`);
                }
                
                if (newStudentData.length === 0) {
                    showAlert('沒有找到有效的學生資料', 'warning');
                    resolve(false);
                    return;
                }
                
                // 更新學生資料
                studentData = newStudentData;
                
                // 計算制服尺寸
                calculateUniformSizes();
                
                // 初始化尺寸分佈資料
                initializeDemandData();
                
                // 排序學生資料
                sortStudentData();
                
                // 儲存學生資料
                saveToLocalStorage('studentData', studentData);
                
                // 不需要重複儲存需求資料，因為已經在 initializeDemandData 中儲存了
                // saveToLocalStorage('demandData', demandData);
                
                // 更新UI
                initStudentTable();
                
                showAlert(`成功匯入 ${newStudentData.length} 筆學生資料${incompleteStudents.length > 0 ? `（忽略 ${incompleteStudents.length} 筆不完整資料）` : ''}`, 'success');
                resolve(true);
            } catch (error) {
                console.error('匯入失敗:', error);
                showAlert(`匯入失敗: ${error.message}`, 'danger');
                reject(error);
            }
        };
        
        reader.onerror = function(error) {
            console.error('檔案讀取失敗:', error);
            showAlert('檔案讀取失敗', 'danger');
            reject(error);
        };
        
        reader.readAsBinaryString(file);
    });
}

/**
 * 標準化單元格資料
 * @param {any} value - 單元格資料
 * @returns {string} - 標準化後的字串
 */
function normalizeCellData(value) {
    if (value === null || value === undefined) return '';
    
    // 轉換為字串
    let strValue = String(value).trim();
    
    // 移除非打印字符
    strValue = strValue.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
    
    return strValue;
}

/**
 * 計算制服尺寸
 */
export function calculateUniformSizes() {
    studentData.forEach(student => {
        // 根據胸圍腰圍計算尺寸，但不直接設置到學生對象中
        // 只計算分配所需的內部參考值
        student._calculatedShirtSize = calculateShirtSize(student);
        student._calculatedPantsSize = calculatePantsSize(student);
    });
}

/**
 * 計算上衣尺寸
 * @param {Object} student - 學生資料
 * @returns {string} - 尺寸代碼
 */
function calculateShirtSize(student) {
    // 獲取有效胸圍（取胸圍和腰圍的最大值）
    let chest = student.chest || 0;
    let waist = student.waist || 0;
    let effectiveChest = Math.max(chest, waist);
    
    // 如果有效胸圍為0，返回空值
    if (effectiveChest <= 0) return '';
    
    // 根據性別選擇加值幅度
    let adjustment = 0;
    if (student.gender === '男') {
        // 男生: 有效胸圍 + 10~12
        // 如果有效胸圍是偶數，加10；如果是奇數，加11（確保結果為偶數）
        adjustment = (effectiveChest % 2 === 0) ? 10 : 11;
    } else {
        // 女生: 有效胸圍 + 8~10
        // 如果有效胸圍是偶數，加8；如果是奇數，加9（確保結果為偶數）
        adjustment = (effectiveChest % 2 === 0) ? 8 : 9;
    }
    
    // 計算最終尺碼值
    let sizeNumber = effectiveChest + adjustment;
    
    // 轉換為對應的尺碼代碼
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

/**
 * 計算褲子尺寸
 * @param {Object} student - 學生資料
 * @returns {string} - 尺寸代碼
 */
function calculatePantsSize(student) {
    let waist = student.waist || 0;
    
    // 根據腰圍判定尺寸
    if (waist <= 0) return '';
    
    if (waist < 69) return 'XS/34';
    if (waist < 73) return 'S/36';
    if (waist < 77) return 'M/38';
    if (waist < 81) return 'L/40';
    if (waist < 85) return 'XL/42';
    if (waist < 89) return '2L/44';
    if (waist < 93) return '3L/46';
    if (waist < 97) return '4L/48';
    if (waist < 101) return '5L/50';
    if (waist < 105) return '6L/52';
    if (waist < 109) return '7L/54';
    if (waist < 113) return '8L/56';
    return '9L/58';
}

/**
 * 初始化需求資料
 */
export function initializeDemandData() {
    // 重設需求資料
    demandData = {
        shortSleeveShirt: { totalDemand: 0, sizeCount: {} },
        shortSleevePants: { totalDemand: 0, sizeCount: {} },
        longSleeveShirt: { totalDemand: 0, sizeCount: {} },
        longSleevePants: { totalDemand: 0, sizeCount: {} }
    };
    
    // 為每個尺寸初始化計數器
    SIZES.forEach(size => {
        demandData.shortSleeveShirt.sizeCount[size] = 0;
        demandData.shortSleevePants.sizeCount[size] = 0;
        demandData.longSleeveShirt.sizeCount[size] = 0;
        demandData.longSleevePants.sizeCount[size] = 0;
    });
    
    // 統計各制服類型的總需求數量（根據學生的各件數欄位）
    studentData.forEach(student => {
        // 短袖上衣（短衣）需求
        const shortShirtCount = parseInt(student.shortSleeveShirtCount) || 0;
        demandData.shortSleeveShirt.totalDemand += shortShirtCount;
        
        // 短袖褲子（短褲）需求
        const shortPantsCount = parseInt(student.shortSleevePantsCount) || 0;
        demandData.shortSleevePants.totalDemand += shortPantsCount;
        
        // 長袖上衣（長衣）需求
        const longShirtCount = parseInt(student.longSleeveShirtCount) || 0;
        demandData.longSleeveShirt.totalDemand += longShirtCount;
        
        // 長袖褲子（長褲）需求
        const longPantsCount = parseInt(student.longSleevePantsCount) || 0;
        demandData.longSleevePants.totalDemand += longPantsCount;
        
        // 如果還需要維持尺寸統計，可以使用計算的尺寸搭配件數來更新
        if (student._calculatedShirtSize && SIZES.includes(student._calculatedShirtSize)) {
            // 短袖上衣尺寸統計
            demandData.shortSleeveShirt.sizeCount[student._calculatedShirtSize] += shortShirtCount;
            
            // 長袖上衣尺寸統計
            demandData.longSleeveShirt.sizeCount[student._calculatedShirtSize] += longShirtCount;
        }
        
        if (student._calculatedPantsSize && SIZES.includes(student._calculatedPantsSize)) {
            // 短袖褲子尺寸統計
            demandData.shortSleevePants.sizeCount[student._calculatedPantsSize] += shortPantsCount;
            
            // 長袖褲子尺寸統計
            demandData.longSleevePants.sizeCount[student._calculatedPantsSize] += longPantsCount;
        }
    });

    // 確保所有需求資料都是數字
    demandData.shortSleeveShirt.totalDemand = parseInt(demandData.shortSleeveShirt.totalDemand) || 0;
    demandData.shortSleevePants.totalDemand = parseInt(demandData.shortSleevePants.totalDemand) || 0;
    demandData.longSleeveShirt.totalDemand = parseInt(demandData.longSleeveShirt.totalDemand) || 0;
    demandData.longSleevePants.totalDemand = parseInt(demandData.longSleevePants.totalDemand) || 0;

    // 儲存更新後的需求資料
    saveToLocalStorage('demandData', demandData);
}

/**
 * 排序學生資料
 */
export function sortStudentData() {
    // 深拷貝學生資料
    sortedStudentData = [...studentData];
    
    // 按班級和座號排序
    sortedStudentData.sort((a, b) => {
        // 比較班級
        const classComparison = String(a.class || '').localeCompare(String(b.class || ''));
        
        if (classComparison !== 0) {
            return classComparison;
        }
        
        // 比較座號
        const aNumber = parseInt(a.number) || 9999;
        const bNumber = parseInt(b.number) || 9999;
        
        return aNumber - bNumber;
    });
}

/**
 * 匯出學生資料
 */
export function exportStudentData() {
    const headers = [
        '班級', '號碼', '姓名', '性別', '胸圍', '腰圍', '褲長',
        '短衣件數', '短褲件數', '長衣件數', '長褲件數'
    ];
    
    const rows = studentData.map(student => [
        student.class || '',
        student.number || '',
        student.name || '',
        student.gender || '',
        student.chest || '',
        student.waist || '',
        student.pantsLength || '',
        student.shortSleeveShirtCount || '',
        student.shortSleevePantsCount || '',
        student.longSleeveShirtCount || '',
        student.longSleevePantsCount || ''
    ]);
    
    const now = new Date();
    const timestamp = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
    const filename = `學生制服尺寸資料_${timestamp}.csv`;
    
    downloadCSV(filename, headers, rows);
}

/**
 * 下載匯入範本
 */
export function downloadTemplate() {
    const headers = [
        '班級', '號碼', '姓名', '性別', '胸圍', '腰圍', '褲長', 
        '短衣件數', '短褲件數', '長衣件數', '長褲件數'
    ];
    
    const rows = [
        ['七年一班', '1', '測試學生', '男', '85', '70', '75', '1', '1', '1', '1'],
        ['七年一班', '2', '測試學生2', '女', '83', '68', '72', '1', '1', '1', '1']
    ];
    
    downloadExcel('制服尺寸匯入範本.xlsx', headers, rows);
}

/**
 * 更新調整頁面資料
 */
export function updateAdjustmentPage() {
    try {
        if (!sortedStudentData || !demandData) {
            console.error('學生資料或需求資料未初始化');
            return;
        }
        
        // 更新學生總數
        const totalStudentsElem = document.getElementById('totalStudentsCount');
        if (totalStudentsElem) {
            totalStudentsElem.textContent = sortedStudentData.length;
        }
        
        // 更新各制服類型需求
        const typeElements = {
            shortSleeveShirt: document.getElementById('shortSleeveShirtDemand'),
            shortSleevePants: document.getElementById('shortSleevePantsDemand'),
            longSleeveShirt: document.getElementById('longSleeveShirtDemand'),
            longSleevePants: document.getElementById('longSleevePantsDemand')
        };
        
        // 如果需求資料不存在，先初始化
        if (!demandData.shortSleeveShirt || !demandData.shortSleevePants || 
            !demandData.longSleeveShirt || !demandData.longSleevePants) {
            initializeDemandData();
        }
        
        // 更新需求數量
        for (const type in typeElements) {
            if (typeElements[type] && demandData[type]) {
                typeElements[type].textContent = demandData[type].totalDemand;
            }
        }
        
        // 更新學生分配表格
        updateStudentAllocationUI();
        
        // 更新分配比率
        updateAllocationRatios();
    } catch (error) {
        console.error('更新調整頁面時發生錯誤:', error);
    }
}

/**
 * 更新學生分配UI
 */
export function updateStudentAllocationUI() {
    // 不再需要更新分配結果，因為表格中已經沒有這兩列
    // 保留此函數以維持兼容性
    console.log('學生分配UI更新 - 已停用表格中的分配列顯示');
}

/**
 * 載入學生資料
 */
export function loadStudentData() {
    // 載入學生資料
    const savedStudentData = loadFromLocalStorage('studentData', null);
    if (savedStudentData) {
        studentData = savedStudentData;
        sortStudentData();
    }
    
    // 載入需求資料
    const savedDemandData = loadFromLocalStorage('demandData', null);
    if (savedDemandData) {
        demandData = savedDemandData;
    } else {
        initializeDemandData();
    }
    
    // 更新UI
    initStudentTable();
} 