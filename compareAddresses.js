const XLSX = require('xlsx');
const fs = require('fs');

const workbook1 = XLSX.readFile('./店铺地址汇总.xlsx');
const workbook2 = XLSX.readFile('./坪山区青春小店介绍.xlsx');

const sheetName1 = workbook1.SheetNames[0];
const sheetName2 = workbook2.SheetNames[0];

const worksheet1 = workbook1.Sheets[sheetName1];
const worksheet2 = workbook2.Sheets[sheetName2];

const data1 = XLSX.utils.sheet_to_json(worksheet1);
const data2 = XLSX.utils.sheet_to_json(worksheet2);

const shopsJson = JSON.parse(fs.readFileSync('./shops.json', 'utf-8'));

const shopsByName = {};
shopsJson.shops.forEach(shop => {
    shopsByName[shop.name] = shop;
});

console.log('=== 店铺地址汇总.xlsx 与 shops.json 对比 ===\n');

let hasDifferences = false;

data1.forEach((row, index) => {
    const name = row['店名'] || row['店铺名称'] || row['名称'];
    const address = row['地址'] || row['小店地址'] || row['店铺地址'];
    const rowNum = index + 2;
    
    if (!name) return;
    
    const shopInJson = shopsByName[name];
    
    if (shopInJson) {
        const jsonAddress = shopInJson.address;
        
        if (address && jsonAddress && address.trim() !== jsonAddress.trim()) {
            console.log(`【地址不一致】第${rowNum}行`);
            console.log(`  店铺名称: ${name}`);
            console.log(`  店铺地址汇总.xlsx: "${address.trim()}"`);
            console.log(`  shops.json: "${jsonAddress.trim()}"`);
            console.log();
            hasDifferences = true;
        }
    } else {
        console.log(`【未找到匹配】第${rowNum}行`);
        console.log(`  店铺名称: ${name}`);
        console.log();
    }
});

console.log('=== 坪山区青春小店介绍.xlsx 与 shops.json 对比 ===\n');

data2.forEach((row, index) => {
    const name = row['店名'] || row['店铺名称'] || row['名称'];
    const address = row['地址'] || row['小店地址'] || row['店铺地址'];
    
    if (!name) return;
    
    const shopInJson = shopsByName[name];
    
    if (shopInJson) {
        const jsonAddress = shopInJson.address;
        const rowNum = index + 2;
        
        if (address && jsonAddress && address.trim() !== jsonAddress.trim()) {
            console.log(`【地址不一致】第${rowNum}行`);
            console.log(`  店铺名称: ${name}`);
            console.log(`  坪山区青春小店介绍.xlsx: "${address.trim()}"`);
            console.log(`  shops.json: "${jsonAddress.trim()}"`);
            console.log();
            hasDifferences = true;
        }
    }
});

if (!hasDifferences) {
    console.log('所有地址数据一致！');
}

console.log(`\n总计检查 ${Object.keys(shopsByName).length} 家店铺`);