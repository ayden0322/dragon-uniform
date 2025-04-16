// 尺寸常數
export const SIZES = ['XS/34', 'S/36', 'M/38', 'L/40', 'XL/42', '2L/44', '3L/46', '4L/48', '5L/50', '6L/52', '7L/54', '8L/56', '9L/58'];

// 尺寸顯示模式
export const SIZE_DISPLAY_MODES = {
    both: 'both',    // 同時顯示尺寸和尺碼 (XS/34)
    size: 'size',    // 僅顯示尺寸 (XS)
    number: 'number' // 僅顯示尺碼 (34)
};

// 默認尺寸顯示模式
export let currentSizeDisplayMode = SIZE_DISPLAY_MODES.both;

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

// 根據當前顯示模式格式化尺寸
export function formatSize(sizeString) {
    if (!sizeString) return '';
    
    // 檢查尺寸格式，假設格式為 "XS/34"
    const parts = sizeString.split('/');
    if (parts.length !== 2) return sizeString; // 如果不是預期的格式，直接返回原始字串
    
    const [size, number] = parts;
    
    switch (currentSizeDisplayMode) {
        case SIZE_DISPLAY_MODES.size:
            return size;
        case SIZE_DISPLAY_MODES.number:
            return number;
        case SIZE_DISPLAY_MODES.both:
        default:
            return sizeString;
    }
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
        femaleChestAdjustment: -2
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
    return config.femaleChestAdjustment || -1.5; // 預設值為-1.5cm
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