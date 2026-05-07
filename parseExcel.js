const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('./坪山区青春小店介绍.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(worksheet, { header: ['序号', '店名', '分类', '介绍', '地址'] });

data.shift();

const imageDir = './店铺图片';
const images = fs.readdirSync(imageDir);

const shops = data.map((item, index) => {
    const name = item['店名'] || `店铺${index + 1}`;
    const cleanName = name.trim().replace(/\s+/g, '');
    const category = item['分类'] || '其他';
    const description = item['介绍'] || '';
    const address = item['地址'] || '';

    const shopImages = images.filter(img => {
        const imgName = img.replace(/-1\.jpg$|-\d+\.(jpg|png)$/i, '').replace(/\s+/g, '');
        return cleanName.includes(imgName) || imgName.includes(cleanName);
    });

    const imagePaths = shopImages.map(img => `/店铺图片/${img}`);

    return {
        id: String(index + 1).padStart(3, '0'),
        name: name.trim(),
        category: category.split('、').map(c => c.trim()).filter(c => c),
        images: imagePaths,
        description: description.trim(),
        address: address.trim(),
        hours: '',
        phone: '',
        coordinate: {
            lat: 22.6 + Math.random() * 0.1,
            lng: 114.2 + Math.random() * 0.1
        },
        mapPosition: {
            x: 150 + Math.random() * 500,
            y: 100 + Math.random() * 500
        }
    };
});

const result = {
    total: shops.length,
    shops: shops
};

fs.writeFileSync('./shops.json', JSON.stringify(result, null, 2));
console.log('店铺数据已生成到 shops.json');
console.log(`共处理 ${shops.length} 家店铺`);