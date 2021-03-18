// ilk olarak npm e excel4node u install ettik
//


let data = [{ 
    ISIM:'Cabbar',
    SOYISIM:'Mikail', 
    YAS:"22", 
    'ALDIGI MAAS':"6000", 
    CINSIYETI: 'ERKEK' 
}, 
{ 
    ISIM:'Hans',
    SOYISIM:'Joe', 
    YAS:"39", 
    'ALDIGI MAAS':"16000", 
    CINSIYETI: 'ERKEK' 
},
{ 
    ISIM:'Murtaza',
    SOYISIM:'Kaya', 
    YAS:"49", 
    'ALDIGI MAAS':"6000", 
    CINSIYETI: 'ERKEK' 
}, 
{ 
    ISIM:'Marion',
    SOYISIM:'Minna', 
    YAS:"55", 
    'ALDIGI MAAS':"9000", 
    CINSIYETI: 'KADIN' 
}, 
{ 
    ISIM:'Murat',
    SOYISIM:'Burhan', 
    YAS:"40", 
    'ALDIGI MAAS':"10000", 
    CINSIYETI: 'ERKEK' 
}, 
{ 
    ISIM:'Abdurrezzak',
    SOYISIM:'Adigüzel', 
    YAS:"22", 
    'ALDIGI MAAS':"6000", 
    CINSIYETI: 'ERKEK' 
}, 
{ 
    ISIM:'Mehmet',
    SOYISIM:'Sökmen', 
    YAS:"33", 
    'ALDIGI MAAS':"12000", 
    CINSIYETI: 'ERKEK' 
}, 

] 

const writeXlsx = require('excel4node');
const wb = new writeXlsx.Workbook();
const ws = wb.addWorksheet('Excel Worksheet');
const headingColumnNames = [
    "ISIM",
    "SOYISIM",
    "YAS",
    "ALDIGI MAAS",
    "CINSIYETI"
]
let headingColumnIndex = 1;
headingColumnNames.forEach(heading => {
    ws.cell(1, headingColumnIndex++)
        .string(heading)
});

let rowIndex = 2;
data.forEach( record => {
    let columnIndex = 1;
    Object.keys(record ).forEach(columnName =>{
        ws.cell(rowIndex,columnIndex++)
            .string(record [columnName])
    });
    rowIndex++;
});
wb.write('ExcelWorksheet.xlsx');