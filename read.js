//ANALIZ
//node.js tarafından sağlanan bir kütüphaneyi kullanarak bir EXCEL dosyasindaki datalari okumak,
//ve ekte verilen datalari EXCEL dosyasinin icinde yeni bir sayfa olusturup yazdirmaK.
//YAPILANLAR
//1. read.js ve write.js dosyalarini olusturduk.
//2. npm init ile npm kutuphanesini actik. package.json ve node module klasorleri acildi.
//3. npm install xlsx komutuyla npm uzerinden excelmdosyalarini okuma ve islem yapma kutuphanesini indirdik.
//4. OrnekDosya.xlsx excel sayfasini actik.
//5. OrnekDosya sayfasini asagidaki 

const xlsxFile = require ("xlsx");

let excel = xlsxFile.readFile('OrnekDosya.xlsx');

let excelPage = excel.SheetNames;

let page1 = xlsxFile.utils.sheet_to_json(excel.Sheets[excelPage[0]]);

console.log(page1);

