// 全局變數
const SIZES = ['XS/34', 'S/36', 'M/38', 'L/40', 'XL/42', '2L/44', '3L/46', '4L/48', '5L/50', '6L/52', '7L/54', '8L/56', '9L/58'];
const UNIFORM_TYPES = {
    shortSleeveShirt: '短袖上衣',
    shortSleevePants: '短袖褲子',
    longSleeveShirt: '長袖上衣',
    longSleevePants: '長袖褲子'
};

// 學校配置
const SCHOOL_CONFIGS = {
    'dragon': {
        name: '土庫國中',
        uniformTypes: {
            shortSleeveShirt: '短袖上衣',
            shortSleevePants: '短袖褲子',
            longSleeveShirt: '長袖上衣',
            longSleevePants: '長袖褲子'
        },
        sizes: ['XS/34', 'S/36', 'M/38', 'L/40', 'XL/42', '2L/44', '3L/46', '4L/48', '5L/50', '6L/52', '7L/54', '8L/56', '9L/58'],
        femaleChestAdjustment: -1.5
    },
    'sample': {
        name: '範例學校',
        uniformTypes: {
            summerShirt: '夏季上衣',
            summerPants: '夏季褲子',
            winterShirt: '冬季上衣',
            winterPants: '冬季褲子',
            sportShirt: '運動服'
        },
        sizes: ['XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', '3XL'],
        femaleChestAdjustment: -2.0
    }
};

// 當前選擇的學校
let currentSchool = 'dragon';

// 庫存數據
let inventory = {
    shortSleeveShirt: {},
    shortSleevePants: {},
    longSleeveShirt: {},
    longSleevePants: {}
};

// 學生數據
let students = [];

// 分配結果
let allocationResults = [];

// DOM 載入完成後執行
document.addEventListener('DOMContentLoaded', function() {
    // 綁定學校選擇事件
    document.getElementById('schoolSelector').addEventListener('change', function() {
        const selectedSchool = this.value;
        if (selectedSchool && SCHOOL_CONFIGS[selectedSchool]) {
            // 更新當前學校
            currentSchool = selectedSchool;
            
            // 更新學校名稱顯示
            document.getElementById('currentSchoolName').textContent = SCHOOL_CONFIGS[currentSchool].name;
            
            // 如果切換到範例學校，顯示提示訊息
            if (currentSchool === 'sample') {
                alert('您已切換到範例學校。這是一個示範用的學校配置，實際功能尚未實現。');
            } else {
                // 重新載入龍頭學校的數據
                loadDataFromLocalStorage();
            }
        }
    });
    
    // 初始化庫存表格
    initInventoryTables();
    
    // 初始化學生表格
    initStudentsTable();
    
    // 綁定按鈕事件
    bindButtonEvents();
    
    // 從本地存儲加載數據
    loadDataFromLocalStorage();
});

// 初始化庫存表格
function initInventoryTables() {
    const tableIds = ['shortSleeveShirtTable', 'shortSleevePantsTable', 'longSleeveShirtTable', 'longSleevePantsTable'];
    const inventoryTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    
    // 初始化庫存數據結構
    if (!inventory) {
        inventory = {
            shortSleeveShirt: {},
            shortSleevePants: {},
            longSleeveShirt: {},
            longSleevePants: {}
        };
    }
    
    // 確保每個制服類型都有 setsPerStudent 屬性
    inventoryTypes.forEach(type => {
        if (inventory[type].setsPerStudent === undefined) {
            inventory[type].setsPerStudent = 1;
        }
    });
    
    // 初始化每個表格
    tableIds.forEach((tableId, index) => {
        const type = inventoryTypes[index];
        const tbody = document.getElementById(tableId);
        tbody.innerHTML = '';
        
        // 為每個尺寸創建一行
        SIZES.forEach(size => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${size}</td>
                <td><input type="number" class="form-control total-inventory" data-size="${size}" min="0" value="${inventory[type][size]?.total || 0}" onfocus="this.select();"></td>
                <td><input type="number" class="form-control reserve-percent" data-size="${size}" min="0" max="100" value="${inventory[type][size]?.reservePercent || 0}" onfocus="this.select();"></td>
            `;
            tbody.appendChild(tr);
            
            // 初始化庫存數據
            if (!inventory[type][size]) {
                inventory[type][size] = {
                    total: 0,
                    reservePercent: 0,
                    reserved: 0,
                    used: 0
                };
            }
        });
        
        // 添加事件監聽器
        const totalInputs = tbody.querySelectorAll('.total-inventory');
        totalInputs.forEach(input => {
            input.addEventListener('input', function() {
                const size = this.getAttribute('data-size');
                inventory[type][size].total = parseInt(this.value) || 0;
                updateInventoryTotals(tableId);
                updateInventoryCalculations(tableId, size);
            });
        });
        
        const reserveInputs = tbody.querySelectorAll('.reserve-percent');
        reserveInputs.forEach(input => {
            input.addEventListener('input', function() {
                const size = this.getAttribute('data-size');
                inventory[type][size].reservePercent = parseInt(this.value) || 0;
                updateInventoryCalculations(tableId, size);
            });
        });
        
        // 初始化每人套數輸入框
        const setsPerStudentInput = document.getElementById(`${type}SetsPerStudent`);
        if (setsPerStudentInput) {
            setsPerStudentInput.value = inventory[type].setsPerStudent || 1;
            setsPerStudentInput.addEventListener('input', function() {
                inventory[type].setsPerStudent = parseInt(this.value) || 1;
                console.log(`${type} 每人套數更新為: ${inventory[type].setsPerStudent}`);
            });
            // 添加 onfocus 屬性，使點擊時自動選中
            setsPerStudentInput.setAttribute('onfocus', 'this.select();');
        }
        
        // 添加加總行
        const totalRow = document.createElement('tr');
        totalRow.className = 'table-secondary';
        totalRow.innerHTML = `
            <td><strong>加總</strong></td>
            <td class="total-sum">0</td>
            <td></td>
        `;
        tbody.appendChild(totalRow);
        
        // 更新總計
        updateInventoryTotals(tableId);
    });
}

// 更新庫存加總
function updateInventoryTotals(tableId) {
    const tbody = document.getElementById(tableId);
    if (!tbody) return;
    
    let totalSum = 0;
    
    // 遍歷所有尺寸行（除了最後一行加總行）
    const rows = tbody.querySelectorAll('tr:not(:last-child)');
    rows.forEach(row => {
        const totalInput = row.querySelector('.total-inventory');
        
        if (totalInput) {
            totalSum += parseInt(totalInput.value) || 0;
        }
    });
    
    // 更新加總行
    const totalRow = tbody.querySelector('tr:last-child');
    if (totalRow) {
        const totalSumCell = totalRow.querySelector('.total-sum');
        
        if (totalSumCell) totalSumCell.textContent = totalSum;
    }
}

// 更新庫存計算
function updateInventoryCalculations(tableId, size) {
    // 找到對應的行
    const rows = document.querySelectorAll(`#${tableId} tr`);
    let row = null;
    
    for (let i = 0; i < rows.length; i++) {
        if (rows[i].cells && rows[i].cells[0] && rows[i].cells[0].textContent.trim() === size) {
            row = rows[i];
            break;
        }
    }
    
    if (!row) return;
    
    const totalInput = row.querySelector('.total-inventory');
    const reserveInput = row.querySelector('.reserve-percent');
    
    // 處理空值，將空字符串轉換為0
    const total = totalInput.value === '' ? 0 : parseInt(totalInput.value) || 0;
    const reservePercent = reserveInput.value === '' ? 0 : parseInt(reserveInput.value) || 0;
    
    // 更新庫存數據
    const inventoryType = tableIdToInventoryType(tableId);
    if (inventoryType) {
        inventory[inventoryType][size].total = total;
        inventory[inventoryType][size].reservePercent = reservePercent;
    }
}

// 表格ID轉換為庫存類型
function tableIdToInventoryType(tableId) {
    switch (tableId) {
        case 'shortSleeveShirtTable': return 'shortSleeveShirt';
        case 'shortSleevePantsTable': return 'shortSleevePants';
        case 'longSleeveShirtTable': return 'longSleeveShirt';
        case 'longSleevePantsTable': return 'longSleevePants';
        default: return null;
    }
}

// 初始化學生表格
function initStudentsTable() {
    const tbody = document.getElementById('studentsTable');
    tbody.innerHTML = '';
}

// 添加學生行
function addStudentRow(student = null) {
    const tbody = document.getElementById('studentsTable');
    const rowIndex = tbody.rows.length + 1;
    
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${rowIndex}</td>
        <td>${student ? student.class : ''}</td>
        <td>${student ? student.number : ''}</td>
        <td>${student ? student.name : ''}</td>
        <td>${student ? student.gender : ''}</td>
        <td>${student ? student.chest : ''}</td>
        <td>${student ? student.waist : ''}</td>
        <td>${student ? student.pantsLength : ''}</td>
        <td>
            <button class="btn btn-sm btn-danger delete-student">刪除</button>
        </td>
    `;
    tbody.appendChild(tr);
    
    // 綁定刪除按鈕事件
    tr.querySelector('.delete-student').addEventListener('click', function() {
        tr.remove();
        // 重新編號
        updateStudentNumbers();
    });
}

// 更新學生編號
function updateStudentNumbers() {
    const rows = document.querySelectorAll('#studentsTable tr');
    rows.forEach((row, index) => {
        const numberCell = row.cells[0];
        if (numberCell) {
            numberCell.textContent = index + 1;
        }
    });
}

// 綁定按鈕事件
function bindButtonEvents() {
    // 添加學生按鈕 - 打開模態框
    document.getElementById('addStudentBtn').addEventListener('click', function() {
        // 清空表單
        document.getElementById('studentClass').value = '';
        document.getElementById('studentNumber').value = '';
        document.getElementById('studentName').value = '';
        document.getElementById('studentGender').value = '';
        document.getElementById('studentChest').value = '';
        document.getElementById('studentWaist').value = '';
        document.getElementById('studentPantLength').value = '';
        
        // 顯示模態框
        const modal = new bootstrap.Modal(document.getElementById('addStudentModal'));
        modal.show();
    });
    
    // 儲存學生按鈕 - 模態框中
    document.getElementById('saveStudentBtn').addEventListener('click', function() {
        const classValue = document.getElementById('studentClass').value.trim();
        const number = parseInt(document.getElementById('studentNumber').value) || 0;
        const name = document.getElementById('studentName').value.trim();
        const gender = document.getElementById('studentGender').value;
        const chest = parseInt(document.getElementById('studentChest').value) || 0;
        const waist = parseInt(document.getElementById('studentWaist').value) || 0;
        const pantsLength = parseInt(document.getElementById('studentPantLength').value) || 0;
        
        // 驗證數據
        if (!name) {
            alert('請輸入學生姓名！');
            return;
        }
        
        if (chest <= 0 || waist <= 0 || pantsLength <= 0) {
            alert('請輸入有效的三圍數據！');
            return;
        }
        
        // 添加學生
        const student = {
            class: classValue,
            number: number,
            name: name,
            gender: gender,
            chest: chest,
            waist: waist,
            pantsLength: pantsLength
        };
        
        students.push(student);
        addStudentRow(student);
        
        // 儲存到本地存儲
        localStorage.setItem('uniformStudents', JSON.stringify(students));
        
        // 關閉模態框
        const modal = bootstrap.Modal.getInstance(document.getElementById('addStudentModal'));
        modal.hide();
    });
    
    // 清空學生數據按鈕
    document.getElementById('clearStudentsBtn').addEventListener('click', function() {
        if (confirm('確定要清空所有學生資料嗎？')) {
            students = [];
            document.getElementById('studentsTable').innerHTML = '';
            localStorage.removeItem('uniformStudents');
        }
    });
    
    // 儲存庫存設定
    document.getElementById('saveInventoryBtn').addEventListener('click', saveInventoryData);
    
    // 清空庫存數字
    document.getElementById('clearInventoryBtn').addEventListener('click', clearInventoryData);
    
    // 儲存學生資料按鈕
    document.getElementById('saveStudentsBtn').addEventListener('click', function() {
        saveStudentsData();
    });
    
    // 執行制服分配按鈕
    document.getElementById('allocateBtn').addEventListener('click', function() {
        if (students.length === 0) {
            alert('請先添加學生資料！');
            return;
        }
        
        allocateUniforms();
        showAllocationResults();
        
        // 切換到結果頁籤
        const resultsTab = document.getElementById('results-tab');
        const resultsTabInstance = bootstrap.Tab.getInstance(resultsTab);
        if (resultsTabInstance) {
            resultsTabInstance.show();
        } else {
            // 如果Tab實例不存在，則創建一個新的實例並顯示
            new bootstrap.Tab(resultsTab).show();
        }
    });
    
    // 匯出Excel按鈕
    document.getElementById('exportExcelBtn').addEventListener('click', function() {
        exportToExcel();
    });

    // 匯入 Excel/CSV 按鈕
    document.getElementById('importExcelBtn').addEventListener('click', function() {
        document.getElementById('fileInput').click();
    });

    // 檔案選擇事件
    document.getElementById('fileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // 假設第一個工作表包含學生資料
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            // 處理匯入的資料
            importStudents(jsonData);
        };
        reader.readAsArrayBuffer(file);
    });
    
    // 下載匯入範本按鈕
    document.getElementById('downloadTemplateBtn').addEventListener('click', function() {
        downloadTemplate();
    });
}

// 清空庫存數字
function clearInventoryData() {
    if (confirm('確定要清空所有庫存數字嗎？這將把所有總庫存量和預留比例設為0。')) {
        const inventoryTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    const tableIds = ['shortSleeveShirtTable', 'shortSleevePantsTable', 'longSleeveShirtTable', 'longSleevePantsTable'];
    
        // 清空庫存數據
        inventoryTypes.forEach((type, index) => {
            const tableId = tableIds[index];
            const rows = document.querySelectorAll(`#${tableId} tr:not(:last-child)`);
            
        rows.forEach(row => {
                if (row.cells && row.cells[0]) {
                    const size = row.cells[0].textContent.trim();
                    if (SIZES.includes(size)) {
                        const totalInput = row.querySelector('.total-inventory');
                        const reserveInput = row.querySelector('.reserve-percent');
                        
                        if (totalInput) totalInput.value = 0;
                        if (reserveInput) reserveInput.value = 0;
                        
                        // 更新庫存數據
                        if (inventory[type][size]) {
                            inventory[type][size].total = 0;
                            inventory[type][size].reservePercent = 0;
                            inventory[type][size].reserved = 0;
                            inventory[type][size].used = 0;
                        }
                    }
                }
            });
            
            // 更新加總行
            updateInventoryTotals(tableId);
    });
    
    // 儲存到本地存儲
    localStorage.setItem('uniformInventory', JSON.stringify(inventory));
        alert('庫存數字已清空！');
    }
}

// 儲存庫存數據
function saveInventoryData() {
    // 保存每人套數設定
    const inventoryTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    inventoryTypes.forEach(type => {
        const setsPerStudentInput = document.getElementById(`${type}SetsPerStudent`);
        if (setsPerStudentInput) {
            inventory[type].setsPerStudent = parseInt(setsPerStudentInput.value) || 1;
        }
    });
    
    // 遍歷所有庫存表格，更新庫存數據
    const tableIds = ['shortSleeveShirtTable', 'shortSleevePantsTable', 'longSleeveShirtTable', 'longSleevePantsTable'];
    
    tableIds.forEach((tableId, index) => {
        const type = inventoryTypes[index];
        const rows = document.querySelectorAll(`#${tableId} tr:not(:last-child)`);
    rows.forEach(row => {
            if (row.cells && row.cells[0]) {
                const size = row.cells[0].textContent.trim();
                if (SIZES.includes(size)) {
                    const totalInput = row.querySelector('.total-inventory');
                    const reserveInput = row.querySelector('.reserve-percent');
                    
                    if (totalInput && reserveInput) {
                        inventory[type][size].total = parseInt(totalInput.value) || 0;
                        inventory[type][size].reservePercent = parseInt(reserveInput.value) || 0;
                    }
                }
            }
        });
        
        // 更新加總行
        updateInventoryTotals(tableId);
    });
    
    // 儲存到本地存儲
    localStorage.setItem('uniformInventory', JSON.stringify(inventory));
    alert('庫存設定已儲存！');
}

// 儲存學生數據
function saveStudentsData() {
    // 已經在添加學生時保存了數據，這裡只需要再次確認
    localStorage.setItem('uniformStudents', JSON.stringify(students));
    alert('學生資料已儲存！');
}

// 從本地存儲加載數據
function loadDataFromLocalStorage() {
    // 加載庫存數據
    const savedInventory = localStorage.getItem('uniformInventory');
    if (savedInventory) {
        inventory = JSON.parse(savedInventory);
        
        // 確保每個制服類型都有 setsPerStudent 屬性
        const inventoryTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
        inventoryTypes.forEach(type => {
            if (inventory[type].setsPerStudent === undefined) {
                inventory[type].setsPerStudent = 1;
            }
            
            // 更新每人套數輸入框
            const setsPerStudentInput = document.getElementById(`${type}SetsPerStudent`);
            if (setsPerStudentInput) {
                setsPerStudentInput.value = inventory[type].setsPerStudent;
            }
        
        // 更新庫存表格
            const tableId = inventoryTypeToTableId(type);
            if (tableId) {
                const rows = document.querySelectorAll(`#${tableId} tr:not(:last-child)`);
                rows.forEach(row => {
                    if (row.cells && row.cells[0]) {
                        const size = row.cells[0].textContent.trim();
                        if (SIZES.includes(size) && inventory[type][size]) {
                            const totalInput = row.querySelector('.total-inventory');
                            const reserveInput = row.querySelector('.reserve-percent');
                            
                            if (totalInput) totalInput.value = inventory[type][size].total || 0;
                            if (reserveInput) reserveInput.value = inventory[type][size].reservePercent || 0;
                        }
                    }
                });
                
                // 更新加總行
                updateInventoryTotals(tableId);
            }
        });
    }
    
    // 加載學生數據
    const savedStudents = localStorage.getItem('uniformStudents');
    if (savedStudents) {
        students = JSON.parse(savedStudents);
        
        // 更新學生表格
        const tbody = document.getElementById('studentsTable');
        tbody.innerHTML = '';
        
        students.forEach(student => {
            addStudentRow(student);
        });
    }
    
    // 加載分配結果
    const savedResults = localStorage.getItem('uniformAllocationResults');
    if (savedResults) {
        allocationResults = JSON.parse(savedResults);
    }
}

// 庫存類型轉換為表格ID
function inventoryTypeToTableId(type) {
    switch (type) {
        case 'shortSleeveShirt': return 'shortSleeveShirtTable';
        case 'shortSleevePants': return 'shortSleevePantsTable';
        case 'longSleeveShirt': return 'longSleeveShirtTable';
        case 'longSleevePants': return 'longSleevePantsTable';
        default: return null;
    }
}

// 分配制服
function allocateUniforms() {
    console.log("開始分配制服...");
    
    // 重置已使用的庫存
    resetUsedInventory();
    
    // 計算預留數量
    calculateReservedQuantities();
    
    // 從本地存儲加載學生數據
    const studentsData = JSON.parse(localStorage.getItem('uniformStudents') || '[]');
    console.log(`加載了 ${studentsData.length} 名學生的數據`);
    
    if (studentsData.length === 0) {
        console.log("沒有學生數據，無法進行分配");
        return;
    }
    
    // 分配制服
    console.log("\n===== 第一階段：按照尺寸順序分配 =====");
    allocateShortSleeveShirts(studentsData);
    allocateShortSleevePants(studentsData);
    allocateLongSleeveShirts(studentsData);
    allocateLongSleevePants(studentsData);
    
    // 處理未分配到制服的學生
    console.log("\n===== 第二階段：處理未分配到制服的學生 =====");
    const uniformTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    
    uniformTypes.forEach(uniformType => {
        handleUnallocatedStudents(studentsData, uniformType);
    });
    
    // 統計分配結果
    console.log("\n===== 分配結果統計 =====");
    uniformTypes.forEach(uniformType => {
        const allocatedCount = studentsData.filter(student => {
            if (uniformType === 'shortSleeveShirt') {
                return student.shortSleeveShirt !== "無庫存";
            } else if (uniformType === 'shortSleevePants') {
                return student.shortSleevePants !== "無庫存";
            } else if (uniformType === 'longSleeveShirt') {
                return student.longSleeveShirt !== "無庫存";
            } else if (uniformType === 'longSleevePants') {
                return student.longSleevePants !== "無庫存";
            }
        }).length;
        
        const unallocatedCount = studentsData.length - allocatedCount;
        console.log(`${uniformType} - 已分配: ${allocatedCount} 名學生, 未分配: ${unallocatedCount} 名學生`);
    });
    
    // 保存分配結果
    allocationResults = studentsData;
    localStorage.setItem('uniformAllocationResults', JSON.stringify(allocationResults));
    
    // 顯示分配結果
    showAllocationResults();
}

// 計算預留數量
function calculateReservedQuantities() {
    // 計算學生總數
    const totalStudents = students.length;
    console.log(`學生總數: ${totalStudents}`);
    
    // 為每個制服類型計算預留數量
    const uniformTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    
    // 儲存每個制服項目的實際總預留量
    inventory.actualTotalReserved = {
        shortSleeveShirt: 0,
        shortSleevePants: 0,
        longSleeveShirt: 0,
        longSleevePants: 0
    };
    
    uniformTypes.forEach(uniformType => {
        // 計算該類型的總庫存
        let totalInventory = 0;
        Object.keys(inventory[uniformType]).forEach(size => {
            if (size !== 'setsPerStudent') {
                totalInventory += inventory[uniformType][size].total || 0;
            }
        });
        
        // 計算該類型的總需求（學生數量 * 每人套數）
        const setsPerStudent = inventory[uniformType].setsPerStudent || 1;
        
        // 計算有庫存的尺寸數量，用於計算平均每人套數
        let sizesWithInventory = 0;
        Object.keys(inventory[uniformType]).forEach(size => {
            if (size !== 'setsPerStudent' && inventory[uniformType][size].total > 0) {
                sizesWithInventory++;
            }
        });
        
        // 計算總需求
        const totalDemand = totalStudents * setsPerStudent;
        console.log(`${uniformType} - 總庫存: ${totalInventory}, 總需求: ${totalDemand}`);
        
        // 計算總預留量（總庫存 - 總需求）
        let totalReserve = totalInventory - totalDemand;
        if (totalReserve < 0) totalReserve = 0;
        console.log(`${uniformType} - 理論總預留量: ${totalReserve}`);
        
        // 計算總預留百分比
        let totalReservePercent = 0;
        Object.keys(inventory[uniformType]).forEach(size => {
            if (size !== 'setsPerStudent') {
                totalReservePercent += inventory[uniformType][size].reservePercent || 0;
            }
        });
        
        // 如果總預留百分比不為100%，調整各尺寸的預留百分比
        if (totalReservePercent !== 100 && totalReservePercent !== 0) {
            console.log(`${uniformType} - 總預留百分比 ${totalReservePercent}% 不為100%，進行調整`);
            const adjustmentFactor = 100 / totalReservePercent;
            
            Object.keys(inventory[uniformType]).forEach(size => {
                if (size !== 'setsPerStudent') {
                    const originalPercent = inventory[uniformType][size].reservePercent || 0;
                    inventory[uniformType][size].reservePercent = originalPercent * adjustmentFactor;
                    console.log(`${uniformType} ${size} - 預留百分比從 ${originalPercent}% 調整為 ${inventory[uniformType][size].reservePercent}%`);
                }
            });
            
            totalReservePercent = 100;
        }
        
        // 實際計算的預留量總和
        let totalCalculatedReserve = 0;
        
        // 計算每個尺寸的預留數量
        Object.keys(inventory[uniformType]).forEach(size => {
            if (size !== 'setsPerStudent') {
                const data = inventory[uniformType][size];
                
                if (totalReservePercent === 100 || totalReservePercent === 0) {
                    // 如果總預留百分比為100%或0%，直接使用原始值
                    data.reservePercent = data.reservePercent || 0;
                }
                
                // 計算該尺寸的預留數量
                if (data.total > 0) {
                    // 根據預留百分比計算預留數量
                    if (totalReserve > 0 && totalReservePercent > 0) {
                        // 計算該尺寸佔總庫存的比例
                        const inventoryRatio = data.total / totalInventory;
                        
                        // 計算該尺寸的預留數量（根據預留百分比或庫存比例）
                        const reserveRatio = data.reservePercent / 100;
                        const reserveByPercent = totalReserve * reserveRatio;
                        const reserveByRatio = totalReserve * inventoryRatio;
                        
                        // 使用預留百分比計算預留數量，無條件進位
                        data.reserved = Math.ceil(reserveByPercent);
                        
                        console.log(`${uniformType} ${size} - 總庫存: ${data.total}, 預留百分比: ${data.reservePercent}%, 理論預留量: ${reserveByPercent}, 實際預留量: ${data.reserved}`);
                    } else {
                        data.reserved = 0;
                        console.log(`${uniformType} ${size} - 總庫存: ${data.total}, 預留量: 0 (總預留量為0或總預留百分比為0)`);
                    }
                    
                    // 確保預留數量不超過總庫存
                    if (data.reserved > data.total) {
                        data.reserved = data.total;
                        console.log(`${uniformType} ${size} - 預留量調整為 ${data.reserved} (不超過總庫存)`);
                    }
                    
                    // 確保可用庫存（總庫存 - 預留量）能被套數整除
                    const setsPerStudent = inventory[uniformType].setsPerStudent || 1;
                    const availableAfterReserve = data.total - data.reserved;
                    const remainder = availableAfterReserve % setsPerStudent;
                    
                    // 如果有餘數，從預留量中取出足夠數量以確保完整分配
                    if (remainder > 0 && data.reserved >= remainder) {
                        const oldReserved = data.reserved;
                        data.reserved -= remainder;
                        console.log(`${uniformType} ${size} - 預留量調整: ${oldReserved} - ${remainder} = ${data.reserved} (確保可用庫存能被套數整除)`);
                    }
                } else {
                    data.reserved = 0;
                }
                
                // 累計實際計算的預留量
                totalCalculatedReserve += data.reserved;
                
                // 累計到對應制服項目的實際總預留量
                inventory.actualTotalReserved[uniformType] += data.reserved;
            }
        });
        
        console.log(`${uniformType} - 實際總預留量: ${inventory.actualTotalReserved[uniformType]} (可能因無條件進位而大於理論總預留量 ${totalReserve})`);
    });
}

// 重置已使用庫存
function resetUsedInventory() {
    Object.keys(inventory).forEach(type => {
        Object.keys(inventory[type]).forEach(size => {
            inventory[type][size].used = 0;
        });
    });
}

// 分配短袖襯衫
function allocateShortSleeveShirts(students) {
    console.log("開始分配短袖襯衫...");
    
    // 根據胸圍排序學生（從小到大）
    students.sort((a, b) => {
        // 女生胸圍減去1.5公分進行比較
        const adjustedChestA = a.gender === '女' ? a.chest - 1.5 : a.chest;
        const adjustedChestB = b.gender === '女' ? b.chest - 1.5 : b.chest;
        return adjustedChestA - adjustedChestB; // 從小到大排序
    });
    
    // 獲取每人套數
    const setsPerStudent = inventory.shortSleeveShirt.setsPerStudent || 1;
    console.log(`短袖襯衫每人套數: ${setsPerStudent}`);
    
    // 按照尺寸順序分配
    students.forEach(student => {
        // 按照尺寸順序嘗試分配
        let allocated = false;
        const adjustedChest = student.gender === '女' ? student.chest - 1.5 : student.chest;
        
        for (let i = 0; i < SIZES.length; i++) {
            const size = SIZES[i];
            if (inventory.shortSleeveShirt[size] && 
                inventory.shortSleeveShirt[size].total - 
                inventory.shortSleeveShirt[size].reserved - 
                inventory.shortSleeveShirt[size].used >= setsPerStudent) {
                
                student.shortSleeveShirt = size;
                student.shortSleeveShirtSets = setsPerStudent;
                inventory.shortSleeveShirt[size].used += setsPerStudent;
                allocated = true;
                console.log(`學生 ${student.name} (胸圍: ${student.chest}cm, 調整後: ${adjustedChest.toFixed(1)}cm) 分配到短袖襯衫尺寸 ${size}，原因: 按照尺寸順序分配，這是第一個有足夠庫存的尺寸`);
                break;
            }
        }
        
        // 如果沒有找到合適的尺寸，標記為無庫存
        if (!allocated) {
            student.shortSleeveShirt = "無庫存";
            student.shortSleeveShirtSets = 0;
            console.log(`學生 ${student.name} (胸圍: ${student.chest}cm, 調整後: ${adjustedChest.toFixed(1)}cm) 無法分配短袖襯衫，原因: 所有尺寸庫存不足`);
        }
    });
}

// 分配短袖褲子
function allocateShortSleevePants(students) {
    console.log("開始分配短袖褲子...");
    
    // 根據腰圍排序學生（從小到大）
    students.sort((a, b) => {
        // 女生腰圍減去1.5公分進行比較
        const adjustedWaistA = a.gender === '女' ? a.waist - 1.5 : a.waist;
        const adjustedWaistB = b.gender === '女' ? b.waist - 1.5 : b.waist;
        return adjustedWaistA - adjustedWaistB; // 從小到大排序
    });
    
    // 獲取每人套數
    const setsPerStudent = inventory.shortSleevePants.setsPerStudent || 1;
    console.log(`短袖褲子每人套數: ${setsPerStudent}`);
    
    // 按照尺寸順序分配
    students.forEach(student => {
        // 檢查學生是否已經分配到短袖上衣
        if (student.shortSleeveShirt === undefined || student.shortSleeveShirt === "無庫存") {
            console.log(`學生 ${student.name} 尚未分配到短袖上衣，無法檢查尺寸差距`);
            student.shortSleevePants = "無庫存";
            student.shortSleevePantsSets = 0;
            return;
        }
        
        // 獲取學生的短袖上衣尺寸在SIZES陣列中的索引
        const shirtSizeIndex = SIZES.indexOf(student.shortSleeveShirt);
        if (shirtSizeIndex === -1) {
            console.log(`學生 ${student.name} 的短袖上衣尺寸 ${student.shortSleeveShirt} 無效`);
            student.shortSleevePants = "無庫存";
            student.shortSleevePantsSets = 0;
            return;
        }
        
        // 計算允許的褲子尺寸範圍（上衣尺寸的前後一個階級內）
        const minAllowedIndex = Math.max(0, shirtSizeIndex - 1);
        const maxAllowedIndex = Math.min(SIZES.length - 1, shirtSizeIndex + 1);
        const allowedSizes = SIZES.slice(minAllowedIndex, maxAllowedIndex + 1);
        
        console.log(`學生 ${student.name} 的短袖上衣尺寸為 ${student.shortSleeveShirt}，允許的短袖褲子尺寸範圍: ${allowedSizes.join(', ')}`);
        
        // 按照尺寸順序嘗試分配
        let allocated = false;
        const adjustedWaist = student.gender === '女' ? student.waist - 1.5 : student.waist;
        
                for (let i = 0; i < SIZES.length; i++) {
                        const size = SIZES[i];
            
            // 檢查尺寸是否在允許範圍內
            if (!allowedSizes.includes(size)) {
                console.log(`尺寸 ${size} 不在允許範圍內，跳過`);
                continue;
            }
            
                        if (inventory.shortSleevePants[size] && 
                inventory.shortSleevePants[size].total - 
                inventory.shortSleevePants[size].reserved - 
                inventory.shortSleevePants[size].used >= setsPerStudent) {
                
                            student.shortSleevePants = size;
                student.shortSleevePantsSets = setsPerStudent;
                inventory.shortSleevePants[size].used += setsPerStudent;
                            allocated = true;
                console.log(`學生 ${student.name} (腰圍: ${student.waist}cm, 調整後: ${adjustedWaist.toFixed(1)}cm) 分配到短袖褲子尺寸 ${size}，原因: 按照尺寸順序分配，這是第一個有足夠庫存且與上衣尺寸差距不超過一個階級的尺寸`);
                            break;
            }
        }
        
        // 如果沒有找到合適的尺寸，標記為無庫存
        if (!allocated) {
            student.shortSleevePants = "無庫存";
            student.shortSleevePantsSets = 0;
            console.log(`學生 ${student.name} (腰圍: ${student.waist}cm, 調整後: ${adjustedWaist.toFixed(1)}cm) 無法分配短袖褲子，原因: 沒有符合條件的尺寸或庫存不足`);
        }
    });
}

// 分配長袖襯衫
function allocateLongSleeveShirts(students) {
    console.log("開始分配長袖襯衫...");
    
    // 根據胸圍排序學生（從小到大）
    students.sort((a, b) => {
        // 女生胸圍減去1.5公分進行比較
        const adjustedChestA = a.gender === '女' ? a.chest - 1.5 : a.chest;
        const adjustedChestB = b.gender === '女' ? b.chest - 1.5 : b.chest;
        return adjustedChestA - adjustedChestB; // 從小到大排序
    });
    
    // 獲取每人套數
    const setsPerStudent = inventory.longSleeveShirt.setsPerStudent || 1;
    console.log(`長袖襯衫每人套數: ${setsPerStudent}`);
    
    // 按照尺寸順序分配
    students.forEach(student => {
        // 按照尺寸順序嘗試分配
        let allocated = false;
        const adjustedChest = student.gender === '女' ? student.chest - 1.5 : student.chest;
        
        for (let i = 0; i < SIZES.length; i++) {
            const size = SIZES[i];
            if (inventory.longSleeveShirt[size] && 
                inventory.longSleeveShirt[size].total - 
                inventory.longSleeveShirt[size].reserved - 
                inventory.longSleeveShirt[size].used >= setsPerStudent) {
                
                student.longSleeveShirt = size;
                student.longSleeveShirtSets = setsPerStudent;
                inventory.longSleeveShirt[size].used += setsPerStudent;
                allocated = true;
                console.log(`學生 ${student.name} (胸圍: ${student.chest}cm, 調整後: ${adjustedChest.toFixed(1)}cm) 分配到長袖襯衫尺寸 ${size}，原因: 按照尺寸順序分配，這是第一個有足夠庫存的尺寸`);
                break;
            }
        }
        
        // 如果沒有找到合適的尺寸，標記為無庫存
        if (!allocated) {
            student.longSleeveShirt = "無庫存";
            student.longSleeveShirtSets = 0;
            console.log(`學生 ${student.name} (胸圍: ${student.chest}cm, 調整後: ${adjustedChest.toFixed(1)}cm) 無法分配長袖襯衫，原因: 所有尺寸庫存不足`);
        }
    });
}

// 分配長袖褲子
function allocateLongSleevePants(students) {
    console.log("開始分配長袖褲子...");
    
    // 根據腰圍排序學生（從小到大）
    students.sort((a, b) => {
        // 不再對女生腰圍進行調整
        return a.waist - b.waist; // 從小到大排序
    });
    
    // 獲取每人套數
    const setsPerStudent = inventory.longSleevePants.setsPerStudent || 1;
    console.log(`長袖褲子每人套數: ${setsPerStudent}`);
    
    // 按照尺寸順序分配
    students.forEach(student => {
        // 按照尺寸順序嘗試分配
        let allocated = false;
        // 不再對女生腰圍進行調整
        const adjustedWaist = student.waist;
        
        for (let i = 0; i < SIZES.length; i++) {
            const size = SIZES[i];
            if (inventory.longSleevePants[size] && 
                inventory.longSleevePants[size].total - 
                inventory.longSleevePants[size].reserved - 
                inventory.longSleevePants[size].used >= setsPerStudent) {
                
                    student.longSleevePants = size;
                student.longSleevePantsSets = setsPerStudent;
                inventory.longSleevePants[size].used += setsPerStudent;
                    allocated = true;
                console.log(`學生 ${student.name} (腰圍: ${student.waist}cm) 分配到長袖褲子尺寸 ${size}，原因: 按照尺寸順序分配，這是第一個有足夠庫存的尺寸`);
                    break;
            }
        }
        
        // 如果沒有找到合適的尺寸，標記為無庫存
        if (!allocated) {
            student.longSleevePants = "無庫存";
            student.longSleevePantsSets = 0;
            console.log(`學生 ${student.name} (腰圍: ${student.waist}cm) 無法分配長袖褲子，原因: 所有尺寸庫存不足`);
        }
    });
}

// 顯示分配結果
function showAllocationResults() {
    const tbody = document.getElementById('resultsTable');
    tbody.innerHTML = '';
    
    // 定義中文數字順序
    const chineseNumbers = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, 
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
    };
    
    // 解析班級名稱的函數
    function parseClassName(className) {
        // 預設值
        let year = 0;
        let classNum = 0;
        
        console.log(`解析班級名稱: ${className}`);
        
        // 嘗試解析 (一)年(一)班 格式
        let match = className.match(/\(([一二三四五六七八九十]+)\)年\(([一二三四五六七八九十]+)\)班/);
        
        // 如果上面的格式不匹配，嘗試解析 一年三班 格式
        if (!match) {
            match = className.match(/([一二三四五六七八九十]+)年([一二三四五六七八九十]+)班/);
        }
        
        if (match) {
            year = chineseNumbers[match[1]] || 0;
            classNum = chineseNumbers[match[2]] || 0;
            console.log(`解析結果 - 年級: ${year}, 班級: ${classNum}`);
        } else {
            console.log(`無法解析班級名稱: ${className}`);
        }
        
        return { year, classNum };
    }
    
    // 先根據班級和號碼排序
    const sortedResults = [...allocationResults].sort((a, b) => {
        // 解析班級名稱
        const aClass = parseClassName(a.class || '');
        const bClass = parseClassName(b.class || '');
        
        // 先按年級排序
        if (aClass.year !== bClass.year) {
            return aClass.year - bClass.year;
        }
        
        // 年級相同時按班級排序
        if (aClass.classNum !== bClass.classNum) {
            return aClass.classNum - bClass.classNum;
        }
        
        // 班級相同時按號碼排序（轉為數字比較）
        const aNumber = parseInt(a.number) || 0;
        const bNumber = parseInt(b.number) || 0;
        return aNumber - bNumber;
    });
    
    // 更新學生總人數顯示
    document.getElementById('totalStudentsCount').textContent = sortedResults.length;
    
    // 按班級分組
    const groupedByClass = {};
    sortedResults.forEach(student => {
        const className = student.class || '未分類';
        if (!groupedByClass[className]) {
            groupedByClass[className] = [];
        }
        groupedByClass[className].push(student);
    });
    
    // 自定義班級排序函數
    function compareClassNames(a, b) {
        const aClass = parseClassName(a);
        const bClass = parseClassName(b);
        
        // 先按年級排序
        if (aClass.year !== bClass.year) {
            return aClass.year - bClass.year;
        }
        
        // 年級相同時按班級排序
        return aClass.classNum - bClass.classNum;
    }
    
    // 遍歷每個班級（按自定義順序排序）
    Object.keys(groupedByClass).sort(compareClassNames).forEach(className => {
        // 添加班級標題行
        const headerRow = document.createElement('tr');
        headerRow.classList.add('table-primary');
        headerRow.innerHTML = `
            <td colspan="15" class="text-center fw-bold">${className} - 共 ${groupedByClass[className].length} 名學生</td>
        `;
        tbody.appendChild(headerRow);
        
        // 添加該班級的學生
        groupedByClass[className].forEach((student, index) => {
        const tr = document.createElement('tr');
            
            // 檢查是否有未分配的項目
            const hasUnallocated = 
                student.shortSleeveShirt === "無庫存" || 
                student.shortSleevePants === "無庫存" || 
                student.longSleeveShirt === "無庫存" || 
                student.longSleevePants === "無庫存";
            
            // 如果有未分配項目，添加樣式
            if (hasUnallocated) {
                tr.classList.add('unallocated-student');
            }
            
            // 準備顯示的尺寸和套數
            const shortSleeveShirtSize = student.shortSleeveShirt !== "無庫存" ? student.shortSleeveShirt : '未分配';
            const shortSleeveShirtSets = student.shortSleeveShirt !== "無庫存" ? student.shortSleeveShirtSets : 0;
            
            const shortSleevePantsSize = student.shortSleevePants !== "無庫存" ? student.shortSleevePants : '未分配';
            const shortSleevePantsSets = student.shortSleevePants !== "無庫存" ? student.shortSleevePantsSets : 0;
            
            const longSleeveShirtSize = student.longSleeveShirt !== "無庫存" ? student.longSleeveShirt : '未分配';
            const longSleeveShirtSets = student.longSleeveShirt !== "無庫存" ? student.longSleeveShirtSets : 0;
            
            const longSleevePantsSize = student.longSleevePants !== "無庫存" ? student.longSleevePants : '未分配';
            const longSleevePantsSets = student.longSleevePants !== "無庫存" ? student.longSleevePantsSets : 0;
            
        tr.innerHTML = `
            <td>${student.class || ''}</td>
            <td>${student.number || ''}</td>
            <td>${student.name}</td>
            <td>${student.gender}</td>
            <td>${student.chest}</td>
            <td>${student.waist}</td>
            <td>${student.pantsLength}</td>
            <td class="${student.shortSleeveShirtSecondAllocation ? 'second-allocation' : ''}">${shortSleeveShirtSize}</td>
            <td class="${student.shortSleeveShirtSecondAllocation ? 'second-allocation' : ''}">${shortSleeveShirtSets}</td>
            <td class="${student.shortSleevePantsSecondAllocation ? 'second-allocation' : ''}">${shortSleevePantsSize}</td>
            <td class="${student.shortSleevePantsSecondAllocation ? 'second-allocation' : ''}">${shortSleevePantsSets}</td>
            <td class="${student.longSleeveShirtSecondAllocation ? 'second-allocation' : ''}">${longSleeveShirtSize}</td>
            <td class="${student.longSleeveShirtSecondAllocation ? 'second-allocation' : ''}">${longSleeveShirtSets}</td>
            <td class="${student.longSleevePantsSecondAllocation ? 'second-allocation' : ''}">${longSleevePantsSize}</td>
            <td class="${student.longSleevePantsSecondAllocation ? 'second-allocation' : ''}">${longSleevePantsSets}</td>
        `;
        tbody.appendChild(tr);
    });
    });
    
    // 更新庫存摘要
    updateInventorySummary();
}

// 更新庫存摘要
function updateInventorySummary() {
    const summaryContainer = document.getElementById('inventorySummaryContainer');
    summaryContainer.innerHTML = '';
    
    // 定義有效的制服類型
    const validTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    
    // 為每種制服類型創建單獨的表格
    validTypes.forEach(type => {
        // 確保該類型在庫存中存在
        if (!inventory[type]) return;
        
        const typeName = UNIFORM_TYPES[type];
        
        // 如果類型名稱未定義，跳過
        if (!typeName) return;
        
        // 創建表格標題
        const titleDiv = document.createElement('div');
        titleDiv.className = 'mt-4';
        titleDiv.innerHTML = `<h6>${typeName}庫存摘要</h6>`;
        summaryContainer.appendChild(titleDiv);
        
        // 創建表格容器
        const tableResponsive = document.createElement('div');
        tableResponsive.className = 'table-responsive mb-4';
        
        // 創建表格
        const table = document.createElement('table');
        table.className = 'table table-bordered table-sm';
        
        // 創建表頭
        const thead = document.createElement('thead');
        thead.innerHTML = `
            <tr class="table-light">
                <th>尺寸</th>
                <th>總庫存</th>
                <th>已分配套數</th>
                <th>已分配人數</th>
                <th>預留數量</th>
            </tr>
        `;
        table.appendChild(thead);
        
        // 創建表格內容
        const tbody = document.createElement('tbody');
        
        // 添加每個尺寸的行
        let hasData = false;
        Object.keys(inventory[type]).forEach(size => {
            if (size !== 'setsPerStudent') {
                const data = inventory[type][size];
                if (data && data.total > 0) {
                    hasData = true;
                    // 計算已分配人數 = 已分配套數 / 每人套數
                    const setsPerStudent = inventory[type].setsPerStudent || 1;
                    const allocatedPeople = Math.floor(data.used / setsPerStudent);
                    
                    // 計算剩餘可用數量
                    const available = data.total - data.reserved - data.used;
                    
                    // 計算預留數量（剩餘可用 + 原預留數量）
                    const totalReserved = available + (data.reserved || 0);
                    
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${size}</td>
                        <td>${data.total}</td>
                        <td>${data.used}</td>
                        <td>${allocatedPeople}</td>
                        <td>${totalReserved}</td>
                    `;
                    tbody.appendChild(tr);
                }
            }
        });
        
        // 如果沒有數據，顯示提示信息
        if (!hasData) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td colspan="5" class="text-center">暫無庫存數據</td>
            `;
            tbody.appendChild(tr);
        }
        
        table.appendChild(tbody);
        tableResponsive.appendChild(table);
        summaryContainer.appendChild(tableResponsive);
    });
}

// 匯出為Excel
function exportToExcel() {
    // 檢查是否有分配結果
    if (allocationResults.length === 0) {
        alert('沒有分配結果可匯出！');
        return;
    }
    
    // 定義中文數字順序（與 showAllocationResults 函數中相同）
    const chineseNumbers = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, 
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
    };
    
    // 解析班級名稱的函數（與 showAllocationResults 函數中相同）
    function parseClassName(className) {
        // 預設值
        let year = 0;
        let classNum = 0;
        
        // 嘗試解析 (一)年(一)班 格式
        let match = className.match(/\(([一二三四五六七八九十]+)\)年\(([一二三四五六七八九十]+)\)班/);
        
        // 如果上面的格式不匹配，嘗試解析 一年三班 格式
        if (!match) {
            match = className.match(/([一二三四五六七八九十]+)年([一二三四五六七八九十]+)班/);
        }
        
        if (match) {
            year = chineseNumbers[match[1]] || 0;
            classNum = chineseNumbers[match[2]] || 0;
        }
        
        return { year, classNum };
    }
    
    // 先根據班級和號碼排序（使用與 showAllocationResults 函數相同的排序邏輯）
    const sortedResults = [...allocationResults].sort((a, b) => {
        // 解析班級名稱
        const aClass = parseClassName(a.class || '');
        const bClass = parseClassName(b.class || '');
        
        // 先按年級排序
        if (aClass.year !== bClass.year) {
            return aClass.year - bClass.year;
        }
        
        // 年級相同時按班級排序
        if (aClass.classNum !== bClass.classNum) {
            return aClass.classNum - bClass.classNum;
        }
        
        // 班級相同時按號碼排序（轉為數字比較）
        const aNumber = parseInt(a.number) || 0;
        const bNumber = parseInt(b.number) || 0;
        return aNumber - bNumber;
    });
    
    // 準備分配結果工作表數據
    const allocation_data = [
        ['班級', '號碼', '姓名', '性別', '胸圍', '腰圍', '褲長', '短袖上衣', '套數', '短袖褲子', '套數', '長袖上衣', '套數', '長袖褲子', '套數']
    ];
    
    sortedResults.forEach(student => {
        // 準備顯示的尺寸和套數
        const shortSleeveShirtSize = student.shortSleeveShirt !== "無庫存" ? student.shortSleeveShirt : '未分配';
        const shortSleeveShirtSets = student.shortSleeveShirt !== "無庫存" ? student.shortSleeveShirtSets : 0;
        
        const shortSleevePantsSize = student.shortSleevePants !== "無庫存" ? student.shortSleevePants : '未分配';
        const shortSleevePantsSets = student.shortSleevePants !== "無庫存" ? student.shortSleevePantsSets : 0;
        
        const longSleeveShirtSize = student.longSleeveShirt !== "無庫存" ? student.longSleeveShirt : '未分配';
        const longSleeveShirtSets = student.longSleeveShirt !== "無庫存" ? student.longSleeveShirtSets : 0;
        
        const longSleevePantsSize = student.longSleevePants !== "無庫存" ? student.longSleevePants : '未分配';
        const longSleevePantsSets = student.longSleevePants !== "無庫存" ? student.longSleevePantsSets : 0;
        
        allocation_data.push([
            student.class || '',
            student.number || '',
            student.name,
            student.gender,
            student.chest,
            student.waist,
            student.pantsLength,
            shortSleeveShirtSize,
            shortSleeveShirtSets,
            shortSleevePantsSize,
            shortSleevePantsSets,
            longSleeveShirtSize,
            longSleeveShirtSets,
            longSleevePantsSize,
            longSleevePantsSets
        ]);
    });
    
    // 準備庫存摘要工作表數據
    const inventory_data = [];
    
    // 定義有效的制服類型
    const validTypes = ['shortSleeveShirt', 'shortSleevePants', 'longSleeveShirt', 'longSleevePants'];
    
    // 為每種制服類型添加數據
    validTypes.forEach(type => {
        // 確保該類型在庫存中存在
        if (!inventory[type]) return;
        
        const typeName = UNIFORM_TYPES[type];
        
        // 如果類型名稱未定義，跳過
        if (!typeName) return;
        
        // 添加制服類型標題行
        inventory_data.push([`${typeName}庫存摘要`]);
        
        // 添加表頭
        inventory_data.push(['尺寸', '總庫存', '已分配套數', '已分配人數', '預留數量']);
        
        // 添加每個尺寸的數據
        let hasData = false;
        Object.keys(inventory[type]).forEach(size => {
            if (size !== 'setsPerStudent') {
                const data = inventory[type][size];
                if (data && data.total > 0) {
                    hasData = true;
                    // 計算已分配人數 = 已分配套數 / 每人套數
                    const setsPerStudent = inventory[type].setsPerStudent || 1;
                    const allocatedPeople = Math.floor(data.used / setsPerStudent);
                    
                    // 計算剩餘可用數量
                    const available = data.total - data.reserved - data.used;
                    
                    // 計算預留數量（剩餘可用 + 原預留數量）
                    const totalReserved = available + (data.reserved || 0);
                    
                    inventory_data.push([
                        size,
                        data.total,
                        data.used,
                        allocatedPeople,
                        totalReserved
                    ]);
                }
            }
        });
        
        // 如果沒有數據，添加提示行
        if (!hasData) {
            inventory_data.push(['暫無庫存數據', '', '', '', '']);
        }
        
        // 添加空行分隔不同制服類型
        inventory_data.push(['', '', '', '', '']);
    });
    
    // 創建工作簿
    const wb = XLSX.utils.book_new();
    
    // 創建並添加分配結果工作表
    const allocation_ws = XLSX.utils.aoa_to_sheet(allocation_data);
    XLSX.utils.book_append_sheet(wb, allocation_ws, '制服分配結果');
    
    // 創建並添加庫存摘要工作表
    const inventory_ws = XLSX.utils.aoa_to_sheet(inventory_data);
    XLSX.utils.book_append_sheet(wb, inventory_ws, '剩餘庫存摘要');
    
    // 保存Excel文件
    XLSX.writeFile(wb, '學生制服分配結果.xlsx');
}

// 匯入學生資料
function importStudents(jsonData) {
    // 確認是否清空現有資料
    let shouldClear = false;
    if (students.length > 0) {
        shouldClear = confirm('是否清空現有學生資料？');
        if (shouldClear) {
            students = [];
            document.getElementById('studentsTable').innerHTML = '';
        }
    }
    
    // 將匯入的資料轉換為學生對象並添加
    const importedStudents = [];
    const errors = [];
    
    jsonData.forEach((row, index) => {
        // 嘗試從不同可能的欄位名稱獲取資料
        const classValue = row['班級'] || row['class'] || '';
        const number = parseInt(row['號碼'] || row['number']) || 0;
        const name = row['姓名'] || row['name'] || '';
        const gender = row['性別'] || row['gender'] || '男';
        const chest = parseInt(row['胸圍'] || row['chest']) || 0;
        const waist = parseInt(row['腰圍'] || row['waist']) || 0;
        const pantsLength = parseInt(row['褲長'] || row['pantsLength']) || 0;
        
        // 驗證必要欄位
        if (!name) {
            errors.push(`第 ${index + 2} 行：缺少姓名`);
            return;
        }
        
        if (chest <= 0 || waist <= 0 || pantsLength <= 0) {
            errors.push(`第 ${index + 2} 行 (${name})：三圍數據不完整或無效`);
            return;
        }
        
        const student = {
            class: classValue,
            number: number,
            name: name,
            gender: gender === '女' ? '女' : '男', // 確保性別值有效
            chest: chest,
            waist: waist,
            pantsLength: pantsLength
        };
        
        importedStudents.push(student);
    });
    
    // 添加有效的學生資料
    importedStudents.forEach(student => {
        students.push(student);
        addStudentRow(student);
    });
    
    // 儲存到本地存儲
    localStorage.setItem('uniformStudents', JSON.stringify(students));
    
    // 顯示結果
    let message = `成功匯入 ${importedStudents.length} 筆學生資料`;
    if (errors.length > 0) {
        message += `\n\n以下資料有誤，未能匯入：\n${errors.join('\n')}`;
    }
    
    alert(message);
}

// 下載匯入範本
function downloadTemplate() {
    // 創建範本數據
    const templateData = [
        {
            '班級': '一年一班',
            '號碼': 1,
            '姓名': '王小明',
            '性別': '男',
            '胸圍': 80,
            '腰圍': 70,
            '褲長': 90
        },
        {
            '班級': '一年一班',
            '號碼': 2,
            '姓名': '李小華',
            '性別': '女',
            '胸圍': 75,
            '腰圍': 65,
            '褲長': 85
        }
    ];
    
    // 創建工作表
    const ws = XLSX.utils.json_to_sheet(templateData);
    
    // 創建工作簿
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '學生資料範本');
    
    // 保存 Excel 檔案
    XLSX.writeFile(wb, '學生資料匯入範本.xlsx');
}

// 處理未分配到制服的學生
function handleUnallocatedStudents(students, uniformType) {
    console.log(`開始處理未分配到 ${uniformType} 的學生...`);
    
    // 計算總庫存
    let totalInventory = 0;
    for (const size in inventory[uniformType]) {
        if (size !== 'setsPerStudent') {
            totalInventory += inventory[uniformType][size].total;
        }
    }
    
    // 獲取每人套數
    const setsPerStudent = inventory[uniformType].setsPerStudent || 1;
    
    // 計算總需求（所有學生的需求）
    const totalStudents = students.length;
    const totalDemand = totalStudents * setsPerStudent;
    
    // 計算理論總預留量（總庫存 - 總需求）
    const theoreticalReserve = Math.max(0, totalInventory - totalDemand);
    
    // 獲取該制服項目的實際總預留量
    const actualTotalReserved = inventory.actualTotalReserved[uniformType];
    
    console.log(`${uniformType} - 總庫存: ${totalInventory}, 總學生數: ${totalStudents}, 每人套數: ${setsPerStudent}`);
    console.log(`${uniformType} - 總需求: ${totalDemand}, 理論總預留量: ${theoreticalReserve}, 實際總預留量: ${actualTotalReserved}`);
    
    // 找出未分配到制服的學生
    const unallocatedStudents = students.filter(student => {
        if (uniformType === 'shortSleeveShirt') {
            return student.shortSleeveShirt === "無庫存";
        } else if (uniformType === 'shortSleevePants') {
            return student.shortSleevePants === "無庫存";
        } else if (uniformType === 'longSleeveShirt') {
            return student.longSleeveShirt === "無庫存";
        } else if (uniformType === 'longSleevePants') {
            return student.longSleevePants === "無庫存";
        }
    });
    
    const unallocatedCount = unallocatedStudents.length;
    console.log(`${uniformType} - 未分配學生數: ${unallocatedCount}`);
    
    if (unallocatedCount === 0) {
        console.log(`${uniformType} - 所有學生都已分配到制服，無需處理`);
        return;
    }
    
    // 計算未分配學生的需求量
    const unallocatedDemand = unallocatedCount * setsPerStudent;
    console.log(`${uniformType} - 未分配學生需求量: ${unallocatedDemand} 套`);
    
    // 檢查是否有額外的預留庫存可用
    let useExtraReserve = false;
    let useOriginalReserve = false;
    
    if (actualTotalReserved > theoreticalReserve) {
        const extraReserve = actualTotalReserved - theoreticalReserve;
        console.log(`${uniformType} - 額外可用預留庫存: ${extraReserve} 套 (實際總預留量 ${actualTotalReserved} - 理論總預留量 ${theoreticalReserve})`);
        
        // 檢查額外預留庫存是否足夠分配給所有未分配的學生
        if (extraReserve >= unallocatedDemand) {
            console.log(`${uniformType} - 額外預留庫存足夠分配給所有未分配的學生 (${extraReserve} >= ${unallocatedDemand})`);
            useExtraReserve = true;
        } else {
            console.log(`${uniformType} - 額外預留庫存不足以分配給所有未分配的學生 (${extraReserve} < ${unallocatedDemand})`);
            console.log(`${uniformType} - 將嘗試從原本預留的庫存中分配`);
            useOriginalReserve = true;
        }
    } else {
        console.log(`${uniformType} - 沒有額外的預留庫存可用 (實際總預留量 ${actualTotalReserved} <= 理論總預留量 ${theoreticalReserve})`);
        console.log(`${uniformType} - 將嘗試從原本預留的庫存中分配`);
        useOriginalReserve = true;
    }
    
    // 顯示未分配學生的詳細資訊
    console.log(`\n${uniformType} - 未分配學生詳細資訊:`);
    unallocatedStudents.forEach((student, index) => {
        let measurementInfo = "";
        if (uniformType === 'shortSleeveShirt' || uniformType === 'longSleeveShirt') {
            const adjustedChest = student.gender === '女' ? student.chest - 1.5 : student.chest;
            measurementInfo = `胸圍: ${student.chest}cm, 調整後: ${adjustedChest.toFixed(1)}cm`;
        } else if (uniformType === 'shortSleevePants') {
            const adjustedWaist = student.gender === '女' ? student.waist - 1.5 : student.waist;
            measurementInfo = `腰圍: ${student.waist}cm, 調整後: ${adjustedWaist.toFixed(1)}cm`;
        } else if (uniformType === 'longSleevePants') {
            measurementInfo = `腰圍: ${student.waist}cm`;
        }
        
        console.log(`  ${index + 1}. 學生: ${student.name}, 班級: ${student.class}, 學號: ${student.number}, 性別: ${student.gender}, ${measurementInfo}`);
    });
    
    // 為每個未分配的學生尋找最合適的尺寸並分配
    console.log(`\n${uniformType} - 開始嘗試分配:`);
    let successfullyAllocated = 0;
    
    unallocatedStudents.forEach((student, index) => {
        const suitableSize = findSuitableSizeForStudent(student, uniformType);
        if (suitableSize) {
            // 檢查該尺寸的預留庫存是否足夠
            if (inventory[uniformType][suitableSize].reserved >= setsPerStudent) {
                console.log(`  ${index + 1}. 學生 ${student.name} - 找到合適的尺寸: ${suitableSize}, 預留庫存: ${inventory[uniformType][suitableSize].reserved} 套`);
                
                // 從預留庫存中分配
                const allocateResult = allocateFromReserved(student, uniformType, suitableSize);
                if (allocateResult) {
                    successfullyAllocated++;
                    console.log(`  ${index + 1}. 學生 ${student.name} - 成功分配尺寸 ${suitableSize}, 原因: ${useExtraReserve ? '使用額外預留庫存' : '使用原本預留庫存'}`);
                } else {
                    console.log(`  ${index + 1}. 學生 ${student.name} - 分配失敗，原因: 預留庫存不足或其他錯誤`);
                }
            } else {
                console.log(`  ${index + 1}. 學生 ${student.name} - 尺寸 ${suitableSize} 預留庫存不足 (${inventory[uniformType][suitableSize].reserved} < ${setsPerStudent})`);
            }
        } else {
            console.log(`  ${index + 1}. 學生 ${student.name} - 無法找到合適的尺寸，原因: 預留庫存中沒有可用尺寸或沒有符合與上衣尺寸差距要求的尺寸`);
        }
    });
    
    console.log(`\n${uniformType} - 分配結果: 成功分配 ${successfullyAllocated}/${unallocatedCount} 名未分配的學生`);
    
    // 檢查是否仍有未分配的學生
    if (successfullyAllocated < unallocatedCount) {
        console.log(`${uniformType} - 仍有 ${unallocatedCount - successfullyAllocated} 名學生未能分配到制服，原因可能是:`);
        console.log(`  1. 找不到合適的尺寸`);
        console.log(`  2. 特定尺寸的預留庫存不足`);
        console.log(`  3. 預留庫存總量不足`);
    }
}

// 為未分配到制服的學生找到最合適的尺寸
function findSuitableSizeForStudent(student, uniformType) {
    let measurement;
    let adjustedMeasurement;
    
    if (uniformType === 'shortSleeveShirt' || uniformType === 'longSleeveShirt') {
        measurement = student.chest;
        adjustedMeasurement = student.gender === '女' ? student.chest - 1.5 : student.chest;
        console.log(`為學生 ${student.name} 尋找合適的 ${uniformType} 尺寸，原始胸圍: ${measurement}cm，調整後胸圍: ${adjustedMeasurement.toFixed(1)}cm`);
    } else if (uniformType === 'shortSleevePants') {
        // 檢查學生是否已經分配到短袖上衣
        if (student.shortSleeveShirt === undefined || student.shortSleeveShirt === "無庫存") {
            console.log(`學生 ${student.name} 尚未分配到短袖上衣，無法檢查尺寸差距`);
            return null;
        }
        
        // 獲取學生的短袖上衣尺寸在SIZES陣列中的索引
        const shirtSizeIndex = SIZES.indexOf(student.shortSleeveShirt);
        if (shirtSizeIndex === -1) {
            console.log(`學生 ${student.name} 的短袖上衣尺寸 ${student.shortSleeveShirt} 無效`);
            return null;
        }
        
        // 計算允許的褲子尺寸範圍（上衣尺寸的前後一個階級內）
        const minAllowedIndex = Math.max(0, shirtSizeIndex - 1);
        const maxAllowedIndex = Math.min(SIZES.length - 1, shirtSizeIndex + 1);
        const allowedSizes = SIZES.slice(minAllowedIndex, maxAllowedIndex + 1);
        
        console.log(`學生 ${student.name} 的短袖上衣尺寸為 ${student.shortSleeveShirt}，允許的短袖褲子尺寸範圍: ${allowedSizes.join(', ')}`);
        
        measurement = student.waist;
        adjustedMeasurement = student.gender === '女' ? student.waist - 1.5 : student.waist;
        console.log(`為學生 ${student.name} 尋找合適的 ${uniformType} 尺寸，原始腰圍: ${measurement}cm，調整後腰圍: ${adjustedMeasurement.toFixed(1)}cm`);
        
        // 從允許的尺寸中找出預留庫存量最多的尺寸
        let bestSize = null;
        let maxReserved = -1;
        let sizesWithReserve = [];
        
        // 先收集所有有預留庫存的尺寸
        for (const size of allowedSizes) {
            if (inventory[uniformType][size] && inventory[uniformType][size].reserved > 0) {
                const reserved = inventory[uniformType][size].reserved;
                sizesWithReserve.push({ size, reserved });
                console.log(`  檢查尺寸 ${size}，預留庫存: ${reserved} 套`);
            }
        }
        
        // 如果沒有任何尺寸有預留庫存，返回 null
        if (sizesWithReserve.length === 0) {
            console.log(`  沒有任何允許的尺寸有預留庫存`);
            return null;
        }
        
        // 按預留庫存量從大到小排序
        sizesWithReserve.sort((a, b) => b.reserved - a.reserved);
        
        // 選擇預留庫存量最多的尺寸
        bestSize = sizesWithReserve[0].size;
        console.log(`  選擇預留庫存量最多的尺寸: ${bestSize}，預留庫存: ${sizesWithReserve[0].reserved} 套`);
        
        return bestSize;
    } else if (uniformType === 'longSleevePants') {
        measurement = student.waist;
        adjustedMeasurement = student.waist; // 不對長袖褲子的女生腰圍進行調整
        console.log(`為學生 ${student.name} 尋找合適的 ${uniformType} 尺寸，腰圍: ${measurement}cm`);
    } else {
        measurement = student.waist;
        adjustedMeasurement = student.gender === '女' ? student.waist - 1.5 : student.waist;
        console.log(`為學生 ${student.name} 尋找合適的 ${uniformType} 尺寸，原始腰圍: ${measurement}cm，調整後腰圍: ${adjustedMeasurement.toFixed(1)}cm`);
    }
    
    // 尋找最接近的尺寸（用於非短袖褲子的情況）
    if (uniformType !== 'shortSleevePants') {
        // 從所有有預留庫存的尺寸中找出預留庫存量最多的尺寸
        let sizesWithReserve = [];
        
        for (const size of SIZES) {
            if (inventory[uniformType][size] && inventory[uniformType][size].reserved > 0) {
                // 獲取尺寸的數字部分
                let sizeValue;
                if (size.includes('/')) {
                    // 如果是 "S/36" 這種格式，取 "/" 後面的數字
                    sizeValue = parseInt(size.split('/')[1]);
                } else {
                    // 如果只是 "S" 這種格式，使用預設值
                    const sizeMap = { 'XS': 34, 'S': 36, 'M': 38, 'L': 40, 'XL': 42, 'XXL': 44 };
                    sizeValue = sizeMap[size] || 0;
                }
                
                // 計算差異
                const difference = Math.abs(adjustedMeasurement - sizeValue);
                const reserved = inventory[uniformType][size].reserved;
                
                sizesWithReserve.push({ size, difference, reserved });
                console.log(`  檢查尺寸 ${size} (數值: ${sizeValue})，與學生測量值差異: ${difference.toFixed(1)}cm，預留庫存: ${reserved} 套`);
            }
        }
        
        // 如果沒有任何尺寸有預留庫存，返回 null
        if (sizesWithReserve.length === 0) {
            console.log(`  沒有任何尺寸有預留庫存`);
            return null;
        }
        
        // 按預留庫存量從大到小排序
        sizesWithReserve.sort((a, b) => b.reserved - a.reserved);
        
        // 選擇預留庫存量最多的尺寸
        const bestSize = sizesWithReserve[0].size;
        console.log(`  選擇預留庫存量最多的尺寸: ${bestSize}，預留庫存: ${sizesWithReserve[0].reserved} 套，與測量值差異: ${sizesWithReserve[0].difference.toFixed(1)}cm`);
        
        return bestSize;
    }
}

// 從預留庫存中分配尺寸
function allocateFromReserved(student, uniformType, size) {
    // 檢查該尺寸是否有預留庫存
    if (!inventory[uniformType][size] || inventory[uniformType][size].reserved <= 0) {
        console.log(`${uniformType} - 無法從預留庫存分配尺寸 ${size} 給學生 ${student.name}，該尺寸沒有預留庫存`);
        return false;
    }
    
    // 獲取每人套數
    const setsPerStudent = inventory[uniformType].setsPerStudent || 1;
    
    // 檢查預留庫存是否足夠
    if (inventory[uniformType][size].reserved < setsPerStudent) {
        console.log(`${uniformType} - 無法從預留庫存分配尺寸 ${size} 給學生 ${student.name}，預留庫存不足 (${inventory[uniformType][size].reserved} < ${setsPerStudent})`);
        return false;
    }
    
    // 從預留庫存中分配，並標記為第二次分配
    if (uniformType === 'shortSleeveShirt') {
        student.shortSleeveShirt = size;
        student.shortSleeveShirtSets = setsPerStudent;
        student.shortSleeveShirtSecondAllocation = true;  // 添加標記
    } else if (uniformType === 'shortSleevePants') {
        student.shortSleevePants = size;
        student.shortSleevePantsSets = setsPerStudent;
        student.shortSleevePantsSecondAllocation = true;  // 添加標記
    } else if (uniformType === 'longSleeveShirt') {
        student.longSleeveShirt = size;
        student.longSleeveShirtSets = setsPerStudent;
        student.longSleeveShirtSecondAllocation = true;  // 添加標記
    } else if (uniformType === 'longSleevePants') {
        student.longSleevePants = size;
        student.longSleevePantsSets = setsPerStudent;
        student.longSleevePantsSecondAllocation = true;  // 添加標記
    }
    
    // 更新預留庫存
    inventory[uniformType][size].reserved -= setsPerStudent;
    
    // 更新實際總預留量
    inventory.actualTotalReserved[uniformType] -= setsPerStudent;
    
    console.log(`${uniformType} - 成功從預留庫存分配尺寸 ${size} 給學生 ${student.name}，使用 ${setsPerStudent} 套`);
    console.log(`${uniformType} - 尺寸 ${size} 剩餘預留庫存: ${inventory[uniformType][size].reserved} 套`);
    console.log(`${uniformType} - 更新後實際總預留量: ${inventory.actualTotalReserved[uniformType]} 套`);
    
    return true;
} 