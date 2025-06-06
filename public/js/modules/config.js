// 尺寸常數
export const SIZES = ['XS/34', 'S/36', 'M/38', 'L/40', 'XL/42', '2L/44', '3L/46', '4L/48', '5L/50', '6L/52', '7L/54', '8L/56', '9L/58'];

// 尺寸顯示模式
export const SIZE_DISPLAY_MODES = {
    both: 'both',    // 同時顯示尺寸和尺碼 (XS/34)
    size: 'size',    // 僅顯示尺寸 (XS)
    number: 'number', // 僅顯示尺碼 (34)
    debug: 'debug'   // Debug模式（顯示詳細錯誤信息）
};

// 默認尺寸顯示模式
export let currentSizeDisplayMode = SIZE_DISPLAY_MODES.both;

// 初始化尺寸顯示模式（從localStorage載入）
export function initSizeDisplayMode() {
    const savedMode = localStorage.getItem('sizeDisplayMode');
    if (savedMode && Object.values(SIZE_DISPLAY_MODES).includes(savedMode)) {
        currentSizeDisplayMode = savedMode;
    }
}

// 分配結果頁面的個別顯示模式配置
export const ALLOCATION_DISPLAY_MODES = {
    shortSleeveShirt: SIZE_DISPLAY_MODES.both,
    shortSleevePants: SIZE_DISPLAY_MODES.both,
    longSleeveShirt: SIZE_DISPLAY_MODES.both,
    longSleevePants: SIZE_DISPLAY_MODES.both
};

// 設置尺寸顯示模式
export function setSizeDisplayMode(mode) {
    if (Object.values(SIZE_DISPLAY_MODES).includes(mode)) {
        currentSizeDisplayMode = mode;
        // 將當前顯示模式保存到 localStorage
        localStorage.setItem('sizeDisplayMode', mode);
        return true;
    }
    return false;
}

// 設置特定制服類型的顯示模式
export function setAllocationDisplayMode(uniformType, mode) {
    if (Object.keys(ALLOCATION_DISPLAY_MODES).includes(uniformType) && 
        Object.values(SIZE_DISPLAY_MODES).includes(mode)) {
        ALLOCATION_DISPLAY_MODES[uniformType] = mode;
        // 保存到localStorage
        localStorage.setItem(`allocationDisplayMode_${uniformType}`, mode);
        return true;
    }
    return false;
}

// 初始化分配顯示模式（從localStorage載入）
export function initAllocationDisplayModes() {
    Object.keys(ALLOCATION_DISPLAY_MODES).forEach(uniformType => {
        const savedMode = localStorage.getItem(`allocationDisplayMode_${uniformType}`);
        if (savedMode && Object.values(SIZE_DISPLAY_MODES).includes(savedMode)) {
            ALLOCATION_DISPLAY_MODES[uniformType] = savedMode;
        }
    });
}

// 根據指定的顯示模式格式化尺寸
export function formatSizeByMode(sizeString, mode) {
    if (!sizeString) return '';
    
    // 檢查尺寸格式，假設格式為 "XS/34"
    const parts = sizeString.split('/');
    
    const [size, number] = parts;
    
    switch (mode) {
        case SIZE_DISPLAY_MODES.size:
            return size;
        case SIZE_DISPLAY_MODES.number:
            return number;
        case SIZE_DISPLAY_MODES.both:
        default:
            return sizeString;
    }
}

// 根據當前顯示模式格式化尺寸
export function formatSize(sizeString) {
    return formatSizeByMode(sizeString, currentSizeDisplayMode);
}

// 制服類型常數
export const UNIFORM_TYPES = {
    shortSleeveShirt: '短衣',
    shortSleevePants: '短褲',
    longSleeveShirt: '長衣',
    longSleevePants: '長褲'
};

// 學校配置
export const SCHOOL_CONFIGS = {
    'dragon': {
        name: '土庫國中',
        uniformTypes: {
            shortSleeveShirt: '短衣',
            shortSleevePants: '短褲',
            longSleeveShirt: '長衣',
            longSleevePants: '長褲'
        },
        sizes: ['XS/34', 'S/36', 'M/38', 'L/40', 'XL/42', '2L/44', '3L/46', '4L/48', '5L/50', '6L/52', '7L/54', '8L/56', '9L/58'],
        femaleChestAdjustment: 0
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
        femaleChestAdjustment: 0
    }
};

// 當前選擇的學校
export let currentSchool = 'dragon';

/**
 * 獲取當前學校的配置
 * @returns {Object} 當前學校的配置信息
 */
export function getCurrentSchoolConfig() {
    return SCHOOL_CONFIGS[currentSchool] || SCHOOL_CONFIGS['dragon']; // 預設使用土庫國中配置
}

/**
 * 獲取當前學校的女生胸圍調整值
 * @returns {number} 女生胸圍調整值 (通常為負數)
 */
export function getFemaleChestAdjustment() {
    const config = getCurrentSchoolConfig();
    return config.femaleChestAdjustment || 0; // 預設值為0
}

/**
 * 設置當前學校
 * @param {string} schoolId - 學校ID
 * @returns {boolean} 設置是否成功
 */
export function setCurrentSchool(schoolId) {
    if (SCHOOL_CONFIGS[schoolId]) {
        currentSchool = schoolId;
        // 將當前學校保存到 localStorage
        localStorage.setItem('currentSchool', schoolId);
        return true;
    }
    return false;
}

// 表格ID轉換為庫存類型
export function tableIdToInventoryType(tableId) {
    switch (tableId) {
        case 'shortSleeveShirtTable': return 'shortSleeveShirt';
        case 'shortSleevePantsTable': return 'shortSleevePants';
        case 'longSleeveShirtTable': return 'longSleeveShirt';
        case 'longSleevePantsTable': return 'longSleevePants';
        default: return null;
    }
}

// 庫存類型轉換為表格ID
export function inventoryTypeToTableId(type) {
    switch (type) {
        case 'shortSleeveShirt': return 'shortSleeveShirtTable';
        case 'shortSleevePants': return 'shortSleevePantsTable';
        case 'longSleeveShirt': return 'longSleeveShirtTable';
        case 'longSleevePants': return 'longSleevePantsTable';
        default: return null;
    }
}

// 預留比例常數（10%）
export const RESERVE_RATIO = 0.1;

// 各品項的預留比例設定
export const RESERVATION_RATIOS = {
    shortSleeveShirt: { // 短袖上衣
        type: 'fixed', // 固定比例
        defaultRatio: 0.10, // 10%
        displayName: '固定預留比例'
    },
    longSleeveShirt: { // 長袖上衣
        type: 'fixed',
        defaultRatio: 0.10, // 10%
        displayName: '固定預留比例'
    },
    shortSleevePants: { // 短袖褲子
        type: 'sized', // 按尺寸比例
        displayName: '預留比例',
        ratios: {
            'XS/34': 0.02, // 2%
            'S/36': 0.03,  // 3%
            'M/38': 0.18,  // 18%
            'L/40': 0.18,  // 18%
            'XL/42': 0.23, // 23%
            '2L/44': 0.13, // 13%
            '3L/46': 0.11, // 11%
            '4L/48': 0.06, // 6%
            '5L/50': 0.00, // 0%
            '6L/52': 0.06, // 6%
            '7L/54': 0.00, // 0%
            '8L/56': 0.00, // 0%
            '9L/58': 0.00  // 0%
        }
    },
    longSleevePants: { // 長袖褲子
        type: 'sized',
        displayName: '預留比例',
        ratios: {
            'XS/34': 0.00, // 0%
            'S/36': 0.04,  // 4%
            'M/38': 0.19,  // 19%
            'L/40': 0.19,  // 19%
            'XL/42': 0.21, // 21%
            '2L/44': 0.14, // 14%
            '3L/46': 0.11, // 11%
            '4L/48': 0.06, // 6%
            '5L/50': 0.00, // 0%
            '6L/52': 0.06, // 6%
            '7L/54': 0.00, // 0%
            '8L/56': 0.00, // 0%
            '9L/58': 0.00  // 0%
        }
    }
}; 