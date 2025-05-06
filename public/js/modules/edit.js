// 修改 checkPantsLengthDeficiency 函數，禁用褲長不足檢查
const fs = require('fs');
const content = fs.readFileSync('allocation.js', 'utf8');
const newContent = content.replace(/(checkPantsLengthDeficiency\(\)[\s\S]*?)(student\.shortPantsLengthDeficiency = true;)/g, '$1// $2');
fs.writeFileSync('allocation.js', newContent);
