/**
 * При загрузке расширения высота панели подгонятся по ВЫСОТЕ контента.
 * Ограничения по габаритам панели заданы в файле CSXS/manifest.xml
 * */

try { fitPanelToContent();} catch (e) { }

/**
 * Скрипт следит за изменением цветовой схемы Иллюстратора
 * Описание скрипта в файле themeChanger.js
 * */
try { changeTheme(csInterface);} catch (e) { }

const reload_panel = document.querySelector('#reload_panel');
reload_panel.addEventListener('click', () => {
 location.reload();
});

loadJSX('json2.js');

getInkCoverage();
getXlsx();
setXlsxData();

/*function setXlsxData() {
 const testBtn = document.querySelector('#test');
 testBtn.addEventListener('click', ()=>{
 csInterface.evalScript('setXlsxData(' +JSON.stringify(o) + ')' , function (result) {
 alert(result);
 });
 })

 }*/

function setXlsxData() {
 const testBtn = document.querySelector('#test');
 testBtn.addEventListener('click', () => {

  csInterface.evalScript('setXlsxData(' + JSON.stringify(getXlsx()) + ')', function (result) {
  });
 });

}

function getXlsx() {
 const input = document.querySelector('#get_xlsx');
 const output = document.querySelector("#output");
 let outStr = '';

 input.addEventListener("change", async e => {

  try {
   return getFf(e);
   /*   const file = e.target.files[0];
    const data = await file.arrayBuffer();

    const workBook = XLSX.read(data);
    const workSheet = workBook.Sheets[workBook.SheetNames[0]];

    const dataForTable = {
    customerCompanyName: workSheet.C4?.v,
    orderNumber: workSheet.C5?.v,
    orderName: workSheet.E6?.v,

    printSideUp: workSheet.F23?.v,
    printSideDn: workSheet.F24?.v,

    formCilinder: workSheet.E40?.v,
    rapport: workSheet.L42?.v,
    streamWidth: workSheet.L40?.v,
    streamsNumber: workSheet.D6?.v,
    supports: workSheet.L44?.v,
    filmWidth: workSheet.E42?.v,

    inkNumber: workSheet.E43?.v,
    inkChange: workSheet.L37?.v,

    sensorLabelSize: workSheet.J49?.v,
    sensorLabelColor: workSheet.J51?.v,
    sensorFieldColor: workSheet.J50?.v,

    filmPrint: workSheet.F53?.v,
    filmCover: workSheet.F54?.v,

    windingSchema: workSheet.L46?.v,

    getFilmComposition() {
    try {
    if (this.filmPrint && this.filmCover) {
    return this.filmPrint + '+' + this.filmCover;
    } else if (this.filmPrint) {
    return this.filmPrint;
    } else {
    return new Error('Уточнить состав готового материала');
    }
    } catch (e) {
    return new Error('Неизвестная ошибка определения состава материала');
    }
    },

    getPrintSide() {
    try {
    if (this.printSideUp) {
    return 'Прямая';
    } else if (this.printSideDn) {
    return 'Обратная';
    } else {
    return new Error('уточнить тип печати');
    }
    } catch (e) {
    return new Error('Неизвестная ошибка определения типа печати');
    }
    },

    getMountWidth() {
    try {
    const MIN_PRINT_FIELDS = 4;
    const MIN_SUPPORTS_CUTTING_FIELDS = 4;
    const MIN_SUPPORTS_WIDTH = 4;
    const filmW = this.filmWidth;
    const printW = +this.streamWidth * +this.streamsNumber;
    const totalFields = filmW - printW - MIN_PRINT_FIELDS;
    if (!this.supports.match(/кресты/i)) return filmW;
    return filmW + 14;
    } catch (e) {
    return new Error('Неизвестная ошибка определения ширины монтажа');
    }

    }
    };

    Object.defineProperty(dataForTable, 'getPrintSide', {enumerable: false});
    Object.defineProperty(dataForTable, 'getFilmComposition', {enumerable: false});
    Object.defineProperty(dataForTable, 'getMountWidth', {enumerable: false});

    for (let key in dataForTable) {
    let val = dataForTable[key];
    outStr += val + ';\n';
    }

    output.value = outStr +
    dataForTable.getPrintSide() + '\n' + dataForTable.getFilmComposition() + '\n' + dataForTable.getMountWidth();

    return dataForTable;*/



  } catch (e) {
   output.innerHTML = e;
   return;
  }

 });
 return 'Вот блять тебе и о!';
}

`Номер C5 
Наименование продукта E6
Заказчик C4

тип печати
F23 - прямая
F24 - обратная

формный вал E40
Раппорт L42
Ширина ручья L40
Количество ручьев D6
Приводные элементы (если есть "кресты", то ширина монтажа + 14 или + 16) L44
Ширина материала (в таблице не надо) E42

Кол-во красок E43
Меняется красок L37

Размер метки J49
Цвет поля метки J50
Цвет метки J51

Материал запечатываемый F53
Матриал покрывной F54

Схема намотки L46
`;

function getInkCoverage() {
 const inkCoverageBtn = document.querySelector('#ink_coverage');
 const output = document.querySelector("#output");

 inkCoverageBtn.addEventListener('click', (e) => {
  try {
   csInterface.evalScript('getXmlStr();', function (result) {
    let xmlString = result;

    let xmlParser = new DOMParser();
    let xmlDoc = xmlParser.parseFromString(xmlString, 'text/xml');
    let outputString = '';

    let inks = xmlDoc.getElementsByTagName("Ink");

    for (let i = 0; i < inks.length; i++) {
     let currInk = inks[i];
     let currVal = currInk.getElementsByTagName('CoveragePerc')[0].getAttribute('Value');
     if (+currVal < 5) currVal = 5;
     outputString += currInk.getAttribute('Name') + ': ' + (+currVal).toFixed(2) + '%\n';
    }

    output.value = outputString.slice(0, -1);
    /* выделить, скопировать, снять выделение */
    output.select();
    document.execCommand('copy');
    output.value += ' '; // костыль
    output.value = output.value.slice(0, -1); // костыль
   });
  } catch (e) {alert(e); }
 });
}

function loadJSX(fileName) {
 var extensionRoot = csInterface.getSystemPath(SystemPath.EXTENSION) + '/assets/jsx/';
 csInterface.evalScript('$.evalFile("' + extensionRoot + fileName + '")');
}

async function getFf(e) {
 const input = document.querySelector('#get_xlsx');
 const output = document.querySelector("#output");
 let outStr = '';

 const file = e.target.files[0];
 const data = await file.arrayBuffer();

 const workBook = XLSX.read(data);
 const workSheet = workBook.Sheets[workBook.SheetNames[0]];

 const dataForTable = {
  customerCompanyName: workSheet.C4?.v,
  orderNumber: workSheet.C5?.v,
  orderName: workSheet.E6?.v,

  printSideUp: workSheet.F23?.v,
  printSideDn: workSheet.F24?.v,

  formCilinder: workSheet.E40?.v,
  rapport: workSheet.L42?.v,
  streamWidth: workSheet.L40?.v,
  streamsNumber: workSheet.D6?.v,
  supports: workSheet.L44?.v,
  filmWidth: workSheet.E42?.v,

  inkNumber: workSheet.E43?.v,
  inkChange: workSheet.L37?.v,

  sensorLabelSize: workSheet.J49?.v,
  sensorLabelColor: workSheet.J51?.v,
  sensorFieldColor: workSheet.J50?.v,

  filmPrint: workSheet.F53?.v,
  filmCover: workSheet.F54?.v,

  windingSchema: workSheet.L46?.v,

  filmComposition: '-', //this.getFilmComposition(),
  printSide: '-', // this.getPrintSide(),

  getFilmComposition() {
   try {
    if (this.filmPrint && this.filmCover) {
     return this.filmPrint + '+' + this.filmCover;
    } else if (this.filmPrint) {
     return this.filmPrint;
    } else {
     return new Error('Уточнить состав готового материала');
    }
   } catch (e) {
    return new Error('Неизвестная ошибка определения состава материала');
   }
  },

  getPrintSide() {
   try {
    if (this.printSideUp) {
     return 'Прямая';
    } else if (this.printSideDn) {
     return 'Обратная';
    } else {
     return new Error('уточнить тип печати');
    }
   } catch (e) {
    return new Error('Неизвестная ошибка определения типа печати');
   }
  },

  /*getMountWidth() {
   try {
   const MIN_PRINT_FIELDS = 4;
   const MIN_SUPPORTS_CUTTING_FIELDS = 4;
   const MIN_SUPPORTS_WIDTH = 4;
   const filmW = this.filmWidth;
   const printW = +this.streamWidth * +this.streamsNumber;
   const totalFields = filmW - printW - MIN_PRINT_FIELDS;
   if (!this.supports.match(/кресты/i)) return filmW;
   return filmW + 14;
   } catch (e) {
   return new Error('Неизвестная ошибка определения ширины монтажа');
   }

   }*/
 };

 Object.defineProperty(dataForTable, 'getPrintSide', {enumerable: false});
 Object.defineProperty(dataForTable, 'getFilmComposition', {enumerable: false});
 // Object.defineProperty(dataForTable, 'getMountWidth', {enumerable: false});

 for (let key in dataForTable) {
  let val = dataForTable[key];
  outStr += val + ';\n';
 }

 output.value = outStr /*+
  dataForTable.getPrintSide() + '\n' + dataForTable.getFilmComposition() + '\n' + dataForTable.getMountWidth()*/;

 csInterface.evalScript('setXlsxData(' + JSON.stringify(dataForTable) + ')', function (result) {
  alert(result);
 });

 return dataForTable;
}
