/**
 * При загрузке расширения высота панели подгонятся по ВЫСОТЕ контента.
 * Ограничения по габаритам панели заданы в файле CSXS/manifest.xml
 * */
let cssProfiles = document.getElementById('css_profiles');
loadCSSProfiles(cssProfiles);

try {
 fitPanelToContent();
} catch (e) {
}

/**
 * Скрипт следит за изменением цветовой схемы Иллюстратора
 * Описание скрипта в файле themeChanger.js
 * */
try {
 changeTheme(csInterface);
} catch (e) {
}

const reload_panel = document.querySelector('#reload_panel');
reload_panel.addEventListener('click', () => {
 location.reload();
});

const close_panel = document.querySelector('#close_panel');
close_panel.addEventListener('click', (e) => {
 if (!e.ctrlKey) return;
 csInterface.closeExtension();
})

loadJSX('json2.js');

const increase_vers = document.querySelector('#increase_vers');
increase_vers.addEventListener('click', (e) => {
 try {
  csInterface.evalScript('increase_vers()', function (result) {
  })
 } catch (e) {
  alert(e)
 }
})

try {
 getInkCoverage();
 getXlsx();
} catch (e) {
 alert(e);
}

cssProfiles.addEventListener('change', (e) => {
 let profileStr = e.target.value;
 csInterface.evalScript('setProfile(' + JSON.stringify(profileStr) + ')', function (result) {
 })
});

function getXlsx() {
 const input = document.querySelector('#get_xlsx');
 const output = document.querySelector("#output");
 let outStr = '';

 input.addEventListener("change", async e => {

  try {
   return getFf(e);
  } catch (e) {
   return new Error('Error in getXlsx' + e.message);
  }

 });
}

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
  } catch (e) {
   return ('Error in getInkCoverage(). ' + e);
  }
 });
}

function loadJSX(fileName) {
 const extensionRoot = csInterface.getSystemPath(SystemPath.EXTENSION) + '/assets/jsx/';
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
  // orderName: workSheet.E6?.v,
  orderName: [
   workSheet.E6?.v, workSheet.E7?.v, workSheet.E8?.v, workSheet.E9?.v, workSheet.E10?.v,
   workSheet.E11?.v, workSheet.E12?.v, workSheet.E13?.v, workSheet.E14?.v
  ],

  printSideUp: workSheet.F23?.v,
  printSideDn: workSheet.F24?.v,

  formCylinder: workSheet.E40?.v,
  rapport: workSheet.L42?.v,
  streamWidth: workSheet.L40?.v,
  // streamsNumber: workSheet.D6?.v,
  streamsNumber: [
   workSheet.D6?.v, workSheet.D7?.v, workSheet.D8?.v, workSheet.D9?.v, workSheet.D10?.v,
   workSheet.D11?.v, workSheet.D12?.v, workSheet.D13?.v, workSheet.D14?.v,
  ],
  supports: workSheet.E22?.v,
  filmWidth: workSheet.E42?.v,

  inkNumber: workSheet.E43?.v,
  inkChange: workSheet.L37?.v,

  sensorLabelSize: workSheet.J49?.v,
  sensorLabelColor: workSheet.J51?.v,
  sensorLabelFieldColor: workSheet.J50?.v,

  filmPrint: workSheet.F53?.v,
  filmCover: workSheet.F54?.v,

  windingSchema: workSheet.L46?.v,
 };

 for (let key in dataForTable) {
  let val = dataForTable[key];
  outStr += key + ': ' + val + ' (' + typeof val + ')\n\n';
 }

 output.value = outStr /*+
  dataForTable.getPrintSide() + '\n' + dataForTable.getFilmComposition() + '\n' + dataForTable.getMountWidth()*/;

 csInterface.evalScript('setXlsxData(' + JSON.stringify(dataForTable) + ')', function (result) {
  outStr += '\n\n' + 'jsx setXlsxData resutlt is: ' + result;
 });

 return dataForTable;
}

 function loadCSSProfiles(selHtmlElem) {
 const cssProfileNames = {
  "PE w": "Poligrafresursy_PE_120_X_DR_0921__C",
  "BOPP tr": "Poligrafresursy_BOPP_tr120_X_DR_0921__C",
  "BOPP mt": "Poligrafresursy_BOPP(mate)_tr120_X_DR_0921__C",
  "dn BOPP tr": "Poligrafresursy_BOPP(rev)_120_X_DR_0921__C",
  "dn BOPP mt": "Poligrafresursy_BOPP(rev)(mate)_120_X_DR_0921__C",
  "dn BOPP tr+w": "Poligrafresursy_BOPP+BOPP_120_X_DR_0921__C",
  "dn BOPP mt+w": "Poligrafresursy_BOPP(mate)+BOPP_120_X_DR_0921__C",
  "dn BOPP tr+me": "Poligrafresursy_BOPP+BOPP(met)_120_X_DR_0921__C",
  "dn BOPP mt+me": "Poligrafresursy_BOPP(mate)+BOPP(met)_120_X_DR_0921__C",
 }
 // let selHtmlElem = document.createElement('select');
 // selHtmlElem.id = 'css_profiles';
 for (let key in cssProfileNames) {
  let optHtmlElem = document.createElement('option');
  optHtmlElem.innerHTML = key;
  optHtmlElem.value = cssProfileNames[key];
  selHtmlElem.append(optHtmlElem);
 }

}