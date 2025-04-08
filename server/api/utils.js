const fs = require('fs');
const path = require('path');

// 數據目錄路徑 - 允許通過環境變量覆蓋
// 如果在Zeabur上，數據目錄應該是掛載的Volume路徑
const dataDir = process.env.DATA_DIR || path.join(__dirname, '../../data');
console.log('使用的數據目錄路徑:', dataDir);

// 確保學校目錄存在
function ensureSchoolDir(schoolId) {
    const schoolDir = path.join(dataDir, schoolId);
    console.log('嘗試訪問學校目錄:', schoolDir);
    if (!fs.existsSync(schoolDir)) {
        console.log('創建學校目錄:', schoolDir);
        fs.mkdirSync(schoolDir, { recursive: true });
    }
    return schoolDir;
}

module.exports = {
    dataDir,
    ensureSchoolDir
};
