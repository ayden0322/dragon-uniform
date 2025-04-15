const fs = require('fs');
const path = require('path');
const { ensureSchoolDir } = require('./utils');

/**
 * @typedef {Object} Student
 * @property {string} id - 學生ID
 * @property {number} chest - 胸圍尺寸
 * @property {number} waist - 腰圍尺寸
 * @property {number} pantsLength - 褲長
 * @property {string} shirtSize - 上衣尺寸
 * @property {string} pantsSize - 短褲尺寸
 * @property {string} longPantsSize - 長褲尺寸
 * @property {number} [requiredQuantity] - 需求件數，可選
 */

/**
 * @typedef {Object} InventoryItem
 * @property {string} id - 庫存項目ID
 * @property {string} size - 尺寸編號
 * @property {number} quantity - 剩餘數量
 */

/**
 * @typedef {Object} Inventory
 * @property {InventoryItem[]} shirts - 上衣庫存
 * @property {InventoryItem[]} pants - 短褲庫存
 * @property {InventoryItem[]} longPants - 長褲庫存
 */

/**
 * @typedef {Object} AllocationResult
 * @property {string} studentId - 學生ID
 * @property {string} [shirtSize] - 分配的上衣尺寸
 * @property {string} [pantsSize] - 分配的短褲尺寸
 * @property {string} [longPantsSize] - 分配的長褲尺寸
 * @property {boolean} allocated - 是否成功分配
 * @property {string} [shirtId] - 分配的上衣ID
 * @property {string} [pantsId] - 分配的短褲ID
 * @property {string} [longPantsId] - 分配的長褲ID
 */

/**
 * 驗證相關功能
 */
const validationService = {
  /**
   * 驗證學生數據
   * @param {Student[]} students - 學生數據
   * @returns {boolean|string} - 成功返回true，失敗返回錯誤訊息
   */
  validateStudents: (students) => {
    if (!Array.isArray(students) || students.length === 0) {
      return '學生數據無效或為空';
    }
    
    // 檢查必要屬性
    for (const student of students) {
      if (!student.id || 
          typeof student.chest !== 'number' || 
          typeof student.waist !== 'number' ||
          typeof student.pantsLength !== 'number') {
        return `學生ID: ${student.id || '未知'} 的數據不完整`;
      }
    }
    
    return true;
  },
  
  /**
   * 驗證庫存數據
   * @param {Inventory} inventory - 庫存數據
   * @returns {boolean|string} - 成功返回true，失敗返回錯誤訊息
   */
  validateInventory: (inventory) => {
    if (!inventory) return '庫存數據為空';
    
    const inventoryTypes = ['shirts', 'pants', 'longPants'];
    for (const type of inventoryTypes) {
      if (!Array.isArray(inventory[type])) {
        return `${type} 庫存數據無效`;
      }
      
      // 檢查每個庫存項目
      for (const item of inventory[type]) {
        if (!item.id || !item.size || typeof item.quantity !== 'number') {
          return `${type} 庫存項目數據不完整: ${JSON.stringify(item)}`;
        }
      }
    }
    
    return true;
  },
  
  /**
   * 驗證分配請求
   * @param {object} data - 請求數據
   * @returns {object} - { isValid: boolean, error: string }
   */
  validateAllocationRequest: (data) => {
    const { studentData, inventoryData, results } = data;
    
    if (!studentData || !inventoryData || !results) {
      return { isValid: false, error: '缺少必要數據' };
    }
    
    const studentsValidation = validationService.validateStudents(
      Array.isArray(studentData.students) ? studentData.students : studentData
    );
    
    if (studentsValidation !== true) {
      return { isValid: false, error: studentsValidation };
    }
    
    const inventoryValidation = validationService.validateInventory(inventoryData);
    if (inventoryValidation !== true) {
      return { isValid: false, error: inventoryValidation };
    }
    
    return { isValid: true, error: null };
  }
};

/**
 * 數據存儲和檔案操作相關功能
 */
const storageService = {
  /**
   * 創建並保存分配歷史記錄
   * @param {string} schoolId - 學校ID
   * @param {string} historyDir - 歷史記錄目錄路徑
   * @param {object} data - 分配數據
   * @returns {string} - 時間戳記錄ID
   * @throws {Error} - 如果寫入文件失敗
   */
  createHistoryRecord: (schoolId, historyDir, data) => {
    const { studentData, inventoryData, adjustmentData, results } = data;
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    
    try {
      // 確保歷史記錄目錄存在
      if (!fs.existsSync(historyDir)) {
        fs.mkdirSync(historyDir, { recursive: true });
      }
      
      // 創建分配歷史記錄
      const historyFile = path.join(historyDir, `allocation-${timestamp}.json`);
      const studentCount = storageService._getStudentCount(studentData);
      
      const history = {
        schoolId,
        allocationDate: new Date().toISOString(),
        studentCount,
        inventorySnapshot: inventoryData,
        studentSnapshot: studentData,
        adjustmentSnapshot: adjustmentData || {},
        results,
      };
      
      fs.writeFileSync(historyFile, JSON.stringify(history, null, 2));
      return timestamp;
    } catch (error) {
      console.error('創建歷史記錄失敗:', error);
      throw new Error(`創建歷史記錄失敗: ${error.message}`);
    }
  },
  
  /**
   * 獲取學生數量
   * @param {object|Array} studentData - 學生數據
   * @returns {number} - 學生數量
   * @private
   */
  _getStudentCount: (studentData) => {
    if (Array.isArray(studentData)) {
      return studentData.length;
    } else if (studentData && Array.isArray(studentData.students)) {
      return studentData.students.length;
    }
    return 0;
  },
  
  /**
   * 更新學校配置文件
   * @param {string} schoolDir - 學校目錄路徑
   * @param {string} schoolId - 學校ID
   * @param {object} data - 配置數據
   * @throws {Error} - 如果寫入文件失敗
   */
  updateSchoolConfigurations: (schoolDir, schoolId, data) => {
    const { studentData, inventoryData, adjustmentData } = data;
    
    try {
      const configurations = [
        { type: 'students', data: studentData },
        { type: 'inventory', data: inventoryData },
        { type: 'adjustments', data: adjustmentData || {} }
      ];
      
      configurations.forEach(config => {
        const configFile = path.join(schoolDir, `${config.type}.json`);
        const configData = {
          schoolId,
          configType: config.type,
          data: config.data,
          updatedAt: new Date().toISOString()
        };
        fs.writeFileSync(configFile, JSON.stringify(configData, null, 2));
      });
    } catch (error) {
      console.error('更新配置文件失敗:', error);
      throw new Error(`更新配置文件失敗: ${error.message}`);
    }
  },
  
  /**
   * 獲取分配歷史記錄列表
   * @param {string} historyDir - 歷史記錄目錄路徑
   * @param {number} limit - 返回記錄的最大數量
   * @returns {Array} - 歷史記錄列表
   */
  getHistoryRecords: (historyDir, limit) => {
    try {
      if (!fs.existsSync(historyDir)) {
        return [];
      }
      
      const files = fs.readdirSync(historyDir)
        .filter(file => file.startsWith('allocation-') && file.endsWith('.json'))
        .sort((a, b) => b.localeCompare(a)) // 按文件名降序排序（較新的文件在前）
        .slice(0, parseInt(limit));
      
      return files.map(file => {
        try {
          const filePath = path.join(historyDir, file);
          const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
          
          // 返回簡要信息
          return {
            id: file.replace('allocation-', '').replace('.json', ''),
            allocationDate: data.allocationDate,
            studentCount: data.studentCount
          };
        } catch (err) {
          console.warn(`讀取歷史記錄 ${file} 失敗: ${err.message}`);
          return null;
        }
      }).filter(record => record !== null);
    } catch (error) {
      console.error('獲取歷史記錄列表失敗:', error);
      return [];
    }
  },
  
  /**
   * 獲取特定分配歷史的詳細信息
   * @param {string} historyFile - 歷史記錄文件路徑
   * @returns {object} - 分配歷史詳細數據
   * @throws {Error} - 如果讀取文件失敗
   */
  getHistoryDetail: (historyFile) => {
    try {
      return JSON.parse(fs.readFileSync(historyFile, 'utf8'));
    } catch (error) {
      console.error('讀取歷史記錄詳情失敗:', error);
      throw new Error(`讀取歷史記錄詳情失敗: ${error.message}`);
    }
  }
};

/**
 * 制服分配算法相關功能
 */
const allocationService = {
  /**
   * 統一分配處理入口
   * @param {Student[]} students - 學生資料列表
   * @param {Inventory} inventory - 庫存資料
   * @returns {object} - 所有分配結果
   */
  allocateAll: (students, inventory) => {
    return {
      shirts: allocationService.allocateShirts(students, inventory),
      pants: allocationService.allocatePants(students, inventory),
      longPants: allocationService.allocateLongPants(students, inventory)
    };
  },
  
  /**
   * 分配短袖上衣
   * @param {Student[]} students - 學生資料列表
   * @param {Inventory} inventory - 庫存資料
   * @returns {AllocationResult[]} - 分配結果
   */
  allocateShirts: (students, inventory) => {
    // 複製庫存進行操作，避免修改原始數據
    const availableItems = allocationService._copyInventoryItems(inventory.shirts);
    const sortedStudents = allocationService._sortStudentsForShirts(students);
    
    // 考慮需求件數的分配
    return allocationService._allocateItemsWithQuantity(
      sortedStudents, 
      availableItems, 
      'shirt', 
      student => student.shirtSize
    );
  },
  
  /**
   * 複製庫存數據，避免修改原始數據
   * @param {InventoryItem[]} items - 庫存項目
   * @returns {InventoryItem[]} - 複製的庫存項目
   * @private
   */
  _copyInventoryItems: (items) => {
    return items.map(item => ({ ...item }));
  },
  
  /**
   * 按照上衣分配規則對學生進行排序
   * @param {Student[]} students - 學生資料列表
   * @returns {Student[]} - 排序後的學生列表
   * @private
   */
  _sortStudentsForShirts: (students) => {
    return students.map(student => {
      // 計算有效胸圍（取胸圍和腰圍的最大值）
      const effectiveChest = Math.max(student.chest, student.waist);
      return { ...student, effectiveChest };
    }).sort((a, b) => {
      // 主要排序依據：有效胸圍
      if (a.effectiveChest !== b.effectiveChest) {
        return a.effectiveChest - b.effectiveChest;
      }
      // 次要排序依據：褲長較短的優先
      return a.pantsLength - b.pantsLength;
    });
  },
  
  /**
   * 分配短褲
   * @param {Student[]} students - 學生資料列表
   * @param {Inventory} inventory - 庫存資料
   * @returns {AllocationResult[]} - 分配結果
   */
  allocatePants: (students, inventory) => {
    const availableItems = allocationService._copyInventoryItems(inventory.pants);
    const sortedStudents = allocationService._sortStudentsForPants(students);
    
    return allocationService._allocateItemsWithQuantity(
      sortedStudents, 
      availableItems, 
      'pants', 
      student => student.pantsSize
    );
  },
  
  /**
   * 按照短褲分配規則對學生進行排序
   * @param {Student[]} students - 學生資料列表
   * @returns {Student[]} - 排序後的學生列表
   * @private
   */
  _sortStudentsForPants: (students) => {
    return students.map(student => {
      return { ...student, effectiveWaist: student.waist };
    }).sort((a, b) => {
      // 主要排序依據：腰圍
      if (a.effectiveWaist !== b.effectiveWaist) {
        return a.effectiveWaist - b.effectiveWaist;
      }
      // 次要排序依據：胸圍較小的優先
      return a.chest - b.chest;
    });
  },
  
  /**
   * 分配長褲
   * @param {Student[]} students - 學生資料列表
   * @param {Inventory} inventory - 庫存資料
   * @returns {AllocationResult[]} - 分配結果
   */
  allocateLongPants: (students, inventory) => {
    const availableItems = allocationService._copyInventoryItems(inventory.longPants);
    const sortedStudents = allocationService._sortStudentsForLongPants(students);
    
    return allocationService._allocateItemsWithQuantity(
      sortedStudents, 
      availableItems, 
      'longPants', 
      student => student.longPantsSize
    );
  },
  
  /**
   * 按照長褲分配規則對學生進行排序
   * @param {Student[]} students - 學生資料列表
   * @returns {Student[]} - 排序後的學生列表
   * @private
   */
  _sortStudentsForLongPants: (students) => {
    return students.map(student => {
      return { ...student, effectiveWaist: student.waist };
    }).sort((a, b) => {
      // 主要排序依據：腰圍
      if (a.effectiveWaist !== b.effectiveWaist) {
        return a.effectiveWaist - b.effectiveWaist;
      }
      // 次要排序依據：褲長較短的優先
      return a.pantsLength - b.pantsLength;
    });
  },
  
  /**
   * 考慮需求件數的物品分配邏輯
   * @param {Student[]} sortedStudents - 已排序的學生列表
   * @param {InventoryItem[]} availableItems - 可用的庫存項目
   * @param {string} type - 物品類型 (shirt/pants/longPants)
   * @param {Function} getSizeFunc - 獲取學生尺寸的函數
   * @returns {AllocationResult[]} - 分配結果
   * @private
   */
  _allocateItemsWithQuantity: (sortedStudents, availableItems, type, getSizeFunc) => {
    const sizeKey = type + 'Size';
    const idKey = type + 'Id';
    const results = [];
    
    // 為每個學生分配物品
    for (const student of sortedStudents) {
      const size = getSizeFunc(student);
      const requiredQuantity = student.requiredQuantity || 1; // 默認需求1件
      
      // 查找匹配尺寸的庫存
      const itemIndex = availableItems.findIndex(item => item.size === size);
      
      if (itemIndex === -1 || availableItems[itemIndex].quantity === 0) {
        // 如果沒有找到匹配的庫存或庫存已用完
        results.push({
          studentId: student.id,
          [sizeKey]: size,
          allocated: false,
          [idKey]: null,
          requiredQuantity,
          allocatedQuantity: 0
        });
        continue;
      }
      
      // 計算可以分配的數量
      const item = availableItems[itemIndex];
      const allocatedQuantity = Math.min(requiredQuantity, item.quantity);
      
      // 更新庫存
      item.quantity -= allocatedQuantity;
      
      // 記錄分配結果
      results.push({
        studentId: student.id,
        [sizeKey]: size,
        allocated: allocatedQuantity > 0,
        [idKey]: item.id,
        requiredQuantity,
        allocatedQuantity
      });
    }
    
    return results;
  }
};

/**
 * 錯誤處理服務
 */
const errorService = {
  /**
   * 處理API錯誤
   * @param {Error} error - 錯誤對象
   * @param {string} operation - 操作名稱
   * @returns {object} - 適合API返回的錯誤對象
   */
  handleApiError: (error, operation) => {
    console.error(`${operation}失敗:`, error);
    return {
      error: `${operation}失敗`,
      message: error.message,
      code: error.code || 'UNKNOWN_ERROR'
    };
  }
};

// API 端點處理函數
/**
 * 保存分配結果
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.saveAllocation = (req, res) => {
  try {
    const { schoolId } = req.params;
    const requestData = {
      studentData: req.body.studentData,
      inventoryData: req.body.inventoryData,
      adjustmentData: req.body.adjustmentData,
      results: req.body.results
    };
    
    // 驗證數據
    const validation = validationService.validateAllocationRequest(requestData);
    if (!validation.isValid) {
      return res.status(400).json({ error: validation.error });
    }
    
    const schoolDir = ensureSchoolDir(schoolId);
    const historyDir = path.join(schoolDir, 'history');
    
    // 保存歷史記錄
    const timestamp = storageService.createHistoryRecord(
      schoolId, 
      historyDir, 
      requestData
    );
    
    // 更新最新配置
    storageService.updateSchoolConfigurations(
      schoolDir, 
      schoolId, 
      requestData
    );
    
    res.json({ 
      success: true, 
      allocationId: timestamp,
      message: '分配結果已成功保存'
    });
  } catch (err) {
    const errorResponse = errorService.handleApiError(err, '保存分配結果');
    res.status(500).json(errorResponse);
  }
};

/**
 * 獲取分配歷史
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.getAllocationHistory = (req, res) => {
  try {
    const { schoolId } = req.params;
    const { limit = 10 } = req.query;
    
    const schoolDir = ensureSchoolDir(schoolId);
    const historyDir = path.join(schoolDir, 'history');
    
    const history = storageService.getHistoryRecords(historyDir, limit);
    res.json(history);
  } catch (err) {
    const errorResponse = errorService.handleApiError(err, '獲取分配歷史');
    res.status(500).json(errorResponse);
  }
};

/**
 * 獲取特定分配的詳細信息
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.getAllocationDetail = (req, res) => {
  try {
    const { schoolId, id } = req.params;
    
    const schoolDir = ensureSchoolDir(schoolId);
    const historyFile = path.join(schoolDir, 'history', `allocation-${id}.json`);
    
    if (!fs.existsSync(historyFile)) {
      return res.status(404).json({ error: '找不到分配記錄' });
    }
    
    const allocation = storageService.getHistoryDetail(historyFile);
    res.json(allocation);
  } catch (err) {
    const errorResponse = errorService.handleApiError(err, '獲取分配詳細信息');
    res.status(500).json(errorResponse);
  }
};

/**
 * 執行分配計算
 * @param {object} req - HTTP請求對象
 * @param {object} res - HTTP回應對象
 */
exports.calculateAllocation = (req, res) => {
  try {
    const { students, inventory } = req.body;
    
    // 驗證輸入數據
    const studentsValidation = validationService.validateStudents(students);
    if (studentsValidation !== true) {
      return res.status(400).json({ error: studentsValidation });
    }
    
    const inventoryValidation = validationService.validateInventory(inventory);
    if (inventoryValidation !== true) {
      return res.status(400).json({ error: inventoryValidation });
    }
    
    // 執行分配
    const results = allocationService.allocateAll(students, inventory);
    res.json({
      success: true,
      results
    });
  } catch (err) {
    const errorResponse = errorService.handleApiError(err, '計算分配結果');
    res.status(500).json(errorResponse);
  }
};

// 導出服務，供其他模組使用
exports.allocationService = allocationService;
exports.storageService = storageService;
exports.validationService = validationService;