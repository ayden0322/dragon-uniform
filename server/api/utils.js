const fs = require('fs');
const path = require('path');

// 數據目錄路徑
const dataDir = path.join(__dirname, '../../data');

// 確保學校目錄存在
function ensureSchoolDir(schoolId) {
    const schoolDir = path.join(dataDir, schoolId);
    if (!fs.existsSync(schoolDir)) {
        fs.mkdirSync(schoolDir, { recursive: true });
    }
    return schoolDir;
}

module.exports = {
    dataDir,
    ensureSchoolDir
};
