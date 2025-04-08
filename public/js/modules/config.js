// 尺寸常數
export const SIZES = ['XS/34', 'S/36', 'M/38', 'L/40', 'XL/42', '2L/44', '3L/46', '4L/48', '5L/50', '6L/52', '7L/54', '8L/56', '9L/58'];

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
export let currentSchool = 'dragon';

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