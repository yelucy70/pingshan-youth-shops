const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('./店铺地址汇总.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(worksheet);

const shopsJson = JSON.parse(fs.readFileSync('./shops.json', 'utf-8'));

let updatedCount = 0;
let matchedCount = 0;
let notFoundCount = 0;

console.log('=== 从店铺地址汇总.xlsx 更新地址 ===\n');

data.forEach((row, index) => {
    const name = row['店名'] || row['店铺名称'] || row['名称'] || '';
    const address = row['地址'] || row['小店地址'] || row['店铺地址'] || '';
    const rowNum = index + 2;
    
    if (!name.trim()) return;
    
    let found = false;
    
    for (let shop of shopsJson.shops) {
        if (shop.name.includes(name.trim()) || name.trim().includes(shop.name)) {
            found = true;
            matchedCount++;
            
            if (address.trim() && shop.address !== address.trim()) {
                console.log(`【更新地址】第${rowNum}行`);
                console.log(`  店铺名称: ${shop.name}`);
                console.log(`  旧地址: "${shop.address}"`);
                console.log(`  新地址: "${address.trim()}"`);
                console.log();
                shop.address = address.trim();
                updatedCount++;
            }
            break;
        }
    }
    
    if (!found) {
        console.log(`【未找到匹配】第${rowNum}行`);
        console.log(`  店铺名称: "${name.trim()}"`);
        console.log();
        notFoundCount++;
    }
});

fs.writeFileSync('./shops.json', JSON.stringify(shopsJson, null, 2));

console.log('=== 更新完成 ===');
console.log(`匹配店铺: ${matchedCount} 家`);
console.log(`更新地址: ${updatedCount} 家`);
console.log(`未找到: ${notFoundCount} 家`);
console.log(`总计店铺: ${shopsJson.shops.length} 家`);

console.log('\nshops.json 已更新！');