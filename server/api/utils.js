const fs = require('fs');
const path = require('path');

// 數據目錄路徑 - 優先使用環境變量
// 如果在Zeabur上，使用環境變量中的DATA_DIR
const dataDir = process.env.DATA_DIR || path.join(__dirname, '../../data');
console.log('API - 使用的數據目錄路徑:', dataDir);

// 確保學校目錄存在
function ensureSchoolDir(schoolId) {
    const schoolDir = path.join(dataDir, schoolId);
    console.log('API - 嘗試訪問學校目錄:', schoolDir);
    if (!fs.existsSync(schoolDir)) {
        console.log('API - 創建學校目錄:', schoolDir);
        fs.mkdirSync(schoolDir, { recursive: true });
    }
    return schoolDir;
}

module.exports = {
    dataDir,
    ensureSchoolDir
};
