/**
 * При загрузке расширения высота панели подгонятся по ВЫСОТЕ контента.
 * Ограничения по габаритам панели заданы в файле CSXS/manifest.xml
 * */
let cssProfiles = document.getElementById('css_profiles');
let reproProfiles = document.getElementById('repro_profiles');
loadCSSProfiles(cssProfiles, reproProfiles);

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
  // alert(e)
 }
})

const prepress = document.querySelector('#prepress');
prepress.addEventListener('change', (e) => {
 try {
  let prepresser = e.target.value;
  csInterface.evalScript('setPrepresser(' + JSON.stringify(prepresser) + ')', function (result) {
  })
 } catch (e) {
  // alert(e)
 }
})

try {
 getInkCoverage();
 getXlsx();
 setInks();
} catch (e) {
 // alert(e);
}

cssProfiles.addEventListener('change', (e) => {
 let profileStr = e.target.value;
 csInterface.evalScript('setProfile(' + JSON.stringify(profileStr) + ')', function (result) {
 })
});
reproProfiles.addEventListener('change', (e) => {
 let profileStr = e.target.value;
 csInterface.evalScript('setProfile(' + JSON.stringify(profileStr) + ')', function (result) {
 })
})

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

function setInks() {
 const TOTAL_INKS = 8;
 const setInksBtn = document.querySelector('#set_inks');
 const output = document.querySelector("#output");
 let inksObj = {};

 setInksBtn.addEventListener('click', (e) => {
  try {
   csInterface.evalScript('getXmlStr();', function (result) {
    let xmlString = result;

    if (xmlString.match('Не найден xml-файл')) return;

    let xmlParser = new DOMParser();
    let xmlDoc = xmlParser.parseFromString(xmlString, 'text/xml');

    let inks = xmlDoc.getElementsByTagName("Ink");

    for (let i = 0; i < TOTAL_INKS; i++) {
     if (inks[i]) {
      let currInk = inks[i];
      let currVal = currInk.getElementsByTagName('CoveragePerc')[0].getAttribute('Value');
       /*
       if (+currVal < 5) currVal = 5;
       inksObj[i] = [currInk.getAttribute('Name'), +(+currVal).toFixed(1)];
       */
      inksObj[i] = [currInk.getAttribute('Name'), Math.ceil(+currVal)];
     } else {
      inksObj[i] = null;
     }
    }

    // const inksObjStingify = JSON.stringify(inksObj);

    // output.innerHTML = inksObjStingify;

    csInterface.evalScript('setInks(' + JSON.stringify(inksObj) + ');', function (result) {
    })

   });
  } catch (e) {
   return ('Error in setInks(). ' + e);
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
     // if (+currVal < 5) currVal = 5;
     // outputString += currInk.getAttribute('Name') + ': ' + +(+currVal).toFixed(1) + '%\n';
     outputString += currInk.getAttribute('Name').replace(/( C$)|(Process )/g, '') + ': ' + Math.ceil(+currVal) + '%\n';
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

function loadCSSProfiles(cssSelHtmlElem, reproSelHtmlElem) {
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
 const reproProfileNames = {
  "Ламинат. Обратная": {
   "60 BOPP мат + BOPP ме + белила 105": "383_060_B105.SSF114.HD.C_BOPPmt.W.BOPPme_Std7_09.2023",
   "61 BOPP мат + BOPP ме + белила 129": "383_061_B129.SSF114.HD.C_BOPPmt.W.BOPPme_Std7_09.2023",
   "62 PET пр + BOPP ме + белила 129": "383_062_B129.SSF114.HD.C_PETtr.W.BOPPme_Std7_09.2023",
   "63 PET пр + BOPP ме + белила 105": "383_063_B105.SSF114.HD.C_PETtr.W.BOPPme_Std7_09.2023",
   "64 PET пр + BOPP пр + белила 129": "383_064_B129.SSF114.HD.C_PETtr.W.BOPPtr_Std7_09.2023",
   "65 PET пр + BOPP пр + белила 105": "383_065_B105.SSF114.HD.C_PETtr.W.BOPPtr_Std7_09.2023",
   "66 PET пр + PE пр + белила 129": "383_066_B129.SSF114.HD.C_PETtr.W.LDPEtr_Std7_09.2023",
   "67 PET пр + PE пр + белила 105": "383_067_B105.SSF114.HD.C_PETtr.W.LDPEtr_Std7_09.2023",
   "68 PET пр + BOPP бел 129": "383_068_B129.SSF114.HD.C_PETtr.N.BOPPw_Std7_09.2023",
   "69 PET пр + BOPP бел 105": "383_069_B105.SSF114.HD.C_PETtr.N.BOPPw_Std7_09.2023",
   "72 BOPP мат + BOPP бел + белила 129": "383_072_B129.SSF114.HD.C_BOPPmt.W.BOPPw_Std7_09.2023",
   "73 BOPP мат + BOPP бел + белила 105": "383_073_B105.SSF114.HD.C_BOPPmt.W.BOPPw_Std7_09.2023",
   "74 BOPP мат + PE пр + белила 129": "383_074_B129.SSF114.HD.C_BOPPmt.W.LDPEtr_Std7_09.2023",
   "75 BOPP мат + PE пр + белила 105": "383_075_B105.SSF114.HD.C_BOPPmt.W.LDPEtr_Std7_09.2023",
   "78 PET пр + BOPP жем 129": "383_078_B129.SSF114.HD.C_PETtr.N.BOPPperl_Std7_09.2023",
   "79 PET пр + BOPP жем 105": "383_079_B105.SSF114.HD.C_PETtr.N.BOPPperl_Std7_09.2023",
   "80 BOPP мат + BOPP жем 129": "383_080_B129.SSF114.HD.C_BOPPmt.N.BOPPperl_Std7_09.2023",
   "81 BOPP мат + BOPP жем 105": "383_081_B105.SSF114.HD.C_BOPPmt.N.BOPPperl_Std7_09.2023",
   "82 PET пр + BOPP бел + белила 129": "383_082_B129.SSF114.HD.C_PETtr.W.BOPPw_Std7_09.2023",
   "83 PET пр + BOPP бел + белила 105": "383_083_B105.SSF114.HD.C_PETtr.W.BOPPw_Std7_09.2023",
   "84 BOPP мат + BOPP бел 129": "383_084_B129.SSF114.HD.C_BOPPmt.N.BOPPw_Std7_09.2023",
   "85 BOPP мат + BOPP бел 105": "383_085_B105.SSF114.HD.C_BOPPmt.N.BOPPw_Std7_09.2023",
   "86 BOPP мат + PE бел 129": "383_086_B129.SSF114.HD.C_BOPPmt.N.LDPEw_Std7_09.2023",
   "87 BOPP мат + PE бел 105": "383_087_B105.SSF114.HD.C_BOPPmt.N.LDPEw_Std7_09.2023",
   "90 BOPP мат + BOPP пр + белила 129": "383_090_B129.SSF114.HD.C_BOPPmt.W.BOPPtr_Std7_09.2023",
   "91 BOPP мат + BOPP пр + белила 105": "383_091_B105.SSF114.HD.C_BOPPmt.W.BOPPtr_Std7_09.2023",
   "92 PET пр + PE бел 129": "383_092_B129.SSF114.HD.C_PETtr.N.LDPEw_Std7_09.2023",
   "93 PET пр + PE бел 105": "383_093_B105.SSF114.HD.C_PETtr.N.LDPEw_Std7_09.2023",
  },
  "Прямая": {
   "41 BOPP пр + белила (129)": "383_041_F129.SSF114.HD.C_BOPPtr.W.N_Std7_09.2023",
   "44 PE бел (129)": "383_044_F129.SSF114.HD.C_LDPEw.N.N_Std7_09.2023",
   "49 BOPP пр + белила (105)": "383_049_F105.SSF114.HD.C_BOPPtr.W.N_Std7_09.2023",
   "50 BOPP бел (105)": "383_050_F105.SSF114.HD.C_BOPPw.N.N_Std7_09.2023",
   "51 BOPP жем (105)": "383_051_F105.SSF114.HD.C_BOPPperl.N.N_Std7_09.2023",
   "52 PE бел (105)": "383_052_F105.SSF114.HD.C_LDPEw.N.N_Std7_09.2023",
   "58 BOPP жем (129)": "383_058_F129.SSF114.HD.C_BOPPperl.N.N_Std7_09.2023",
   "59 BOPP бел (129)": "383_059_F129.SSF114.HD.C_BOPPw.N.N_Std7_09.2023",
  },
  "Обратная": {
   "45 PET пр + белила 129": "383_045_B129.SSF114.HD.C_PETtr.W.N_Std7_09.2023",
   "46 PET пр 129": "383_046_B129.SSF114.HD.C_PETtr.N.N_Std7_09.2023",
   "47 BOPP мат + белила 129": "383_047_B129.SSF114.HD.C_BOPPmt.W.N_Std7_09.2023",
   "48 BOPP мат 129": "383_048_B129.SSF114.HD.C_BOPPmt.N.N_Std7_09.2023",
   "53 PET пр + белила 105": "383_053_B105.SSF114.HD.C_PETtr.W.N_Std7_09.2023",
   "54 PET пр 105": "383_054_B105.SSF114.HD.C_PETtr.N.N_Std7_09.2023",
   "55 BOPP мат + белила 105": "383_055_B105.SSF114.HD.C_BOPPmt.W.N_Std7_09.2023",
   "56 BOPP мат 105": "383_056_B105.SSF114.HD.C_BOPPmt.N.N_Std7_09.2023",
  },
  "Ламинат. Прямая": {
   "70 PE бел + PE пр 129": "383_070_F129.SSF114.HD.C_LDPEw.N.LDPEtr_Std7_09.2023",
   "71 PE бел + PE пр 105": "383_071_F105.SSF114.HD.C_LDPEw.N.LDPEtr_Std7_09.2023",
   "76 BOPP жем + BOPP пр 129": "383_076_F129.SSF114.HD.C_BOPPperl.N.BOPPtr_Std7_09.2023",
   "77 BOPP жем + BOPP пр 105": "383_077_F105.SSF114.HD.C_BOPPperl.N.BOPPtr_Std7_09.2023",
   "88 BOPP бел + BOPP пр 129": "383_088_F129.SSF114.HD.C_BOPPw.N.BOPPtr_Std7_09.2023",
   "89 BOPP бел + BOPP пр 105": "383_089_F105.SSF114.HD.C_BOPPw.N.BOPPtr_Std7_09.2023",
   "94 BOPP бел + BOPP мат 129": "383_094_F129.SSF114.HD.C_BOPPw.N.BOPPmt_Std7_09.2023",
   "95 BOPP бел + BOPP мат 105": "383_095_F105.SSF114.HD.C_BOPPw.N.BOPPmt_Std7_09.2023",
  },

 }

 for (let key in cssProfileNames) {
  let optHtmlElem = document.createElement('option');
  optHtmlElem.innerHTML = key;
  optHtmlElem.value = cssProfileNames[key];
  cssSelHtmlElem.append(optHtmlElem);
 }

 for (let key in reproProfileNames) {
  let optGr = document.createElement('optgroup');
  optGr.label = key;
  let printTypeObj = reproProfileNames[key];
  for (let j in printTypeObj) {
   let optEl = document.createElement('option');
   optEl.innerHTML = j;
   optEl.value = printTypeObj[j];
   optGr.append(optEl);
  }
  reproSelHtmlElem.append(optGr);
 }

}
