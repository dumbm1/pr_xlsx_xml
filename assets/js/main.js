/**
 * При загрузке расширения высота панели подгонятся по ВЫСОТЕ контента.
 * Ограничения по габаритам панели заданы в файле CSXS/manifest.xml
 * */
let cssProfiles = document.getElementById('css_profiles');
let reproProfiles2023 = document.getElementById('repro_profiles_2023');
let reproProfiles2025 = document.getElementById('repro_profiles_2025');
loadProfiles(cssProfiles, reproProfiles2023, reproProfiles2025);

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

const change_vers = document.querySelector('#change_vers');
change_vers.addEventListener('focus', (e) => {
 try {
  csInterface.evalScript(getVers.toString() + '; getVers()', function (result) {
   if (!result) return;
   e.target.value = +result;
  })
 } catch (e) {
  alert('getVers Error\n' + e);
 }
})
change_vers.addEventListener('change', (e) => {
 try {
  var currNumber = e.target.value;
  csInterface.evalScript(
   changeVers.toString() + '; changeVers(' + JSON.stringify(currNumber) + ')',
   function (result) {
   })
 } catch (e) {
  alert('changeVers Error\n' + e);
 }
})

const prepress = document.querySelector('#prepress');
prepress.addEventListener('change', (e) => {
 try {
  let prepresser = e.target.value;
  csInterface.evalScript(setPrepresser.toString() + ';setPrepresser(' + JSON.stringify(prepresser) + ')', function (result) {
  })
 } catch (e) {
  // alert(e)
 }
})

try {
 getInkCoverage();
 getXlsx();
 setInksFromXml();
} catch (e) {
 // alert(e);
}

cssProfiles.addEventListener('change', (e) => {
 let profileStr = e.target.value;
 csInterface.evalScript(setProfile.toString() + ';setProfile(' + JSON.stringify(profileStr) + ')', function (result) {
 })
});
reproProfiles2023.addEventListener('change', (e) => {
 let profileStr = e.target.value;
 csInterface.evalScript(setProfile.toString() + ';setProfile(' + JSON.stringify(profileStr) + ')', function (result) {
 })
})
reproProfiles2025.addEventListener('change', (e) => {
 let profileStr = e.target.value;
 csInterface.evalScript(setProfile.toString() + ';setProfile(' + JSON.stringify(profileStr) + ')', function (result) {
 })
})


/**
 * ECMA6 Script Lib
 * */
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

function setInksFromXml() {
 const TOTAL_INKS = 8;
 const setInksBtn = document.querySelector('#set_inks');
 const output = document.querySelector("#output");
 let inksObj = {};

 setInksBtn.addEventListener('click', (e) => {
  try {
   csInterface.evalScript(getXmlStr.toString() + ';getXmlStr();', function (result) {
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

    csInterface.evalScript(setInks.toString() + ';setInks(' + JSON.stringify(inksObj) + ');', function (result) {
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
   csInterface.evalScript(getXmlStr.toString() + ';getXmlStr();', function (result) {
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

 csInterface.evalScript(setXlsxData.toString() + ';setXlsxData(' + JSON.stringify(dataForTable) + ')', function (result) {
  outStr += '\n\n' + 'jsx setXlsxData resutlt is: ' + result;
 });

 return dataForTable;
}

function loadProfiles(cssSelHtmlElem, repro2023SelHtmlElem, repro2025SelHtmlElem) {
 const cssProfileNames = {
  "PE w": "PE_120_X_DR_0921__C",
  "BOPP tr": "BOPP_tr120_X_DR_0921__C",
  "BOPP mt": "BOPP(mate)_tr120_X_DR_0921__C",
  "dn BOPP tr": "BOPP(rev)_120_X_DR_0921__C",
  "dn BOPP mt": "BOPP(rev)(mate)_120_X_DR_0921__C",
  "dn BOPP tr+w": "BOPP+BOPP_120_X_DR_0921__C",
  "dn BOPP mt+w": "BOPP(mate)+BOPP_120_X_DR_0921__C",
  "dn BOPP tr+me": "BOPP+BOPP(met)_120_X_DR_0921__C",
  "dn BOPP mt+me": "BOPP(mate)+BOPP(met)_120_X_DR_0921__C",
 }
 const reproProfileNames2023 = {
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

 const reproProfileNames2025 = {
  "Ламинат. Обратная": {
   "114 BOPPtr + W + BOPPme 129": "114_B129.SSF114.HD.C_BOPPtr.W.BOPPme_Std8_06.25",
   "115 BOPPtr + W + BOPPtr 129": "115_B129.SSF114.HD.C_BOPPtr.W.BOPPtr_Std8_06.25",
   "116 BOPPtr + W + PEtr 129": "116_B129.SSF114.HD.C_BOPPtr.W.PEtr_Std8_06.25",
   "117 BOPPtr + BOPPw 129": "117_B129.SSF114.HD.C_BOPPtr.N.BOPPw_Std8_06.25",
   "118 BOPPtr + BOPPperl 129": "118_B129.SSF114.HD.C_BOPPtr.N.BOPPperl_Std8_06.25",
   "119 BOPPmt + W + BOPPtr 129": "119_B129.SSF114.HD.C_BOPPmt.W.BOPPtr_Std8_06.25",
   "120 BOPPmt + W + PEtr 129": "120_B129.SSF114.HD.C_BOPPmt.W.PEtr_Std8_06.25",
   "121 BOPPmt + W + BOPPme 129": "121_B129.SSF114.HD.C_BOPPmt.W.BOPPme_Std8_06.25",
   "122 BOPPmt + BOPPw 129": "122_B129.SSF114.HD.C_BOPPmt.N.BOPPw_Std8_06.25",
   "129 BOPPtr + LDPEw 129": "129_B129.SSF114.HD.C_BOPPtr.N.LDPEw_Std8_06.25",
   "130 BOPPtr + W + BOPPme 105": "130_B105.SSF114.HD.C_BOPPtr.W.BOPPme_Std8_06.25",
   "131 BOPPtr + W + BOPPtr 105": "131_B105.SSF114.HD.C_BOPPtr.W.BOPPtr_Std8_06.25",
   "132 BOPPtr + W + PEtr 105": "132_B105.SSF114.HD.C_BOPPtr.W.PEtr_Std8_06.25",
   "133 BOPPtr + BOPPw 105": "133_B105.SSF114.HD.C_BOPPtr.N.BOPPw_Std8_06.25",
   "134 BOPPtr + BOPPperl 105": "134_B105.SSF114.HD.C_BOPPtr.N.BOPPperl_Std8_06.25",
   "135 BOPPmt + W + BOPPtr 105": "135_B105.SSF114.HD.C_BOPPmt.W.BOPPtr_Std8_06.25",
   "136 BOPPmt + W + PEtr 105": "136_B105.SSF114.HD.C_BOPPmt.W.PEtr_Std8_06.25",
   "137 BOPPmt + W + BOPPme 105": "137_B105.SSF114.HD.C_BOPPmt.W.BOPPme_Std8_06.25",
   "138 BOPPmt + BOPPw 105": "138_B105.SSF114.HD.C_BOPPmt.N.BOPPw_Std8_06.25",
   "145 BOPPtr + LDPEw 105": "145_B105.SSF114.HD.C_BOPPtr.N.LDPEw_Std8_06.25",
  },
  "Прямая": {
   "096 BOPPme + W 129": "096_F129.SSF114.HD.C_BOPPme.W.N_Std8_06.25",
   "097 BOPPtr + W 129": "097_F129.SSF114.HD.C_BOPPtr.W.N_Std8_06.25",
   "098 BOPPw 129": "098_F129.SSF114.HD.C_BOPPw.N.N_Std8_06.25",
   "099 BOPPperl 129": "099_F129.SSF114.HD.C_BOPPperl.N.N_Std8_06.25",
   "100 LDPEw 129": "100_F129.SSF114.HD.C_LDPEw.N.N_Std8_06.25",
   "101 BOPPtr + W 129": "101_B129.SSF114.HD.C_BOPPtr.W.N_Std8_06.25",
   "102 BOPPtr 129": "102_B129.SSF114.HD.C_BOPPtr.N.N_Std8_06.25",
   "103 BOPPmt + W 129": "103_B129.SSF114.HD.C_BOPPmt.W.N_Std8_06.25",
   "104 BOPPmt 129": "104_B129.SSF114.HD.C_BOPPmt.N.N_Std8_06.25",
   "105 BOPPme + W 105": "105_F105.SSF114.HD.C_BOPPme.W.N_Std8_06.25",
   "106 BOPPtr + W 105": "106_F105.SSF114.HD.C_BOPPtr.W.N_Std8_06.25",
   "107 BOPPw 105": "107_F105.SSF114.HD.C_BOPPw.N.N_Std8_06.25",
   "108 BOPPperl 105": "108_F105.SSF114.HD.C_BOPPperl.N.N_Std8_06.25",
   "109 LDPEw 105": "109_F105.SSF114.HD.C_LDPEw.N.N_Std8_06.25",
  },
  "Обратная": {
   "101 BOPPtr + W 129": "101_B129.SSF114.HD.C_BOPPtr.W.N_Std8_06.25",
   "102 BOPPtr 129": "102_B129.SSF114.HD.C_BOPPtr.N.N_Std8_06.25",
   "103 BOPPmt + W 129": "103_B129.SSF114.HD.C_BOPPmt.W.N_Std8_06.25",
   "104 BOPPmt 129": "104_B129.SSF114.HD.C_BOPPmt.N.N_Std8_06.25",
   "105 BOPPme + W 105": "105_F105.SSF114.HD.C_BOPPme.W.N_Std8_06.25",
   "106 BOPPtr + W 105": "106_F105.SSF114.HD.C_BOPPtr.W.N_Std8_06.25",
   "107 BOPPw 105": "107_F105.SSF114.HD.C_BOPPw.N.N_Std8_06.25",
   "108 BOPPperl 105": "108_F105.SSF114.HD.C_BOPPperl.N.N_Std8_06.25",
   "109 LDPEw 105": "109_F105.SSF114.HD.C_LDPEw.N.N_Std8_06.25",
   "110 BOPPtr + W 105": "110_B105.SSF114.HD.C_BOPPtr.W.N_Std8_06.25",
   "111 BOPPtr 105": "111_B105.SSF114.HD.C_BOPPtr.N.N_Std8_06.25",
   "112 BOPPmt + W 105": "112_B105.SSF114.HD.C_BOPPmt.W.N_Std8_06.25",
   "113 BOPPmt 105": "113_B105.SSF114.HD.C_BOPPmt.N.N_Std8_06.25",
  },
  "Ламинат. Прямая": {
   "123 BOPPw + BOPPtr 129": "123_F129.SSF114.HD.C_BOPPw.N.BOPPtr_Std8_06.25",
   "124 BOPPw + BOPPmt 129": "124_F129.SSF114.HD.C_BOPPw.N.BOPPmt_Std8_06.25",
   "125 LDPEw + BOPPtr 129": "125_F129.SSF114.HD.C_LDPEw.N.BOPPtr_Std8_06.25",
   "126 LDPEw + BOPPmt 129": "126_F129.SSF114.HD.C_LDPEw.N.BOPPmt_Std8_06.25",
   "127 BOPPperl + BOPPtr 129": "127_F129.SSF114.HD.C_BOPPperl.N.BOPPtr_Std8_06.25",
   "128 BOPPperl + BOPPmt 129": "128_F129.SSF114.HD.C_BOPPperl.N.BOPPmt_Std8_06.25",
   "139 BOPPw + BOPPtr 105": "139_F105.SSF114.HD.C_BOPPw.N.BOPPtr_Std8_06.25",
   "140 BOPPw + BOPPmt 105": "140_F105.SSF114.HD.C_BOPPw.N.BOPPmt_Std8_06.25",
   "141 LDPEw + BOPPtr 105": "141_F105.SSF114.HD.C_LDPEw.N.BOPPtr_Std8_06.25",
   "142 LDPEw + BOPPmt 105": "142_F105.SSF114.HD.C_LDPEw.N.BOPPmt_Std8_06.25",
   "143 BOPPperl + BOPPtr 105": "143_F105.SSF114.HD.C_BOPPperl.N.BOPPtr_Std8_06.25",
   "144 BOPPperl + BOPPmt 105": "144_F105.SSF114.HD.C_BOPPperl.N.BOPPmt_Std8_06.25",
  },

 }

 for (let key in cssProfileNames) {
  let optHtmlElem = document.createElement('option');
  optHtmlElem.innerHTML = key;
  optHtmlElem.value = cssProfileNames[key];
  cssSelHtmlElem.append(optHtmlElem);
 }

 for (let key in reproProfileNames2023) {
  let optGr = document.createElement('optgroup');
  optGr.label = key;
  let printTypeObj = reproProfileNames2023[key];
  for (let j in printTypeObj) {
   let optEl = document.createElement('option');
   optEl.innerHTML = j;
   optEl.value = printTypeObj[j];
   optGr.append(optEl);
  }
  repro2023SelHtmlElem.append(optGr);
 }

 for (let key in reproProfileNames2025) {
  let optGr = document.createElement('optgroup');
  optGr.label = key;
  let printTypeObj = reproProfileNames2025[key];
  for (let j in printTypeObj) {
   let optEl = document.createElement('option');
   optEl.innerHTML = j;
   optEl.value = printTypeObj[j];
   optGr.append(optEl);
  }
  repro2025SelHtmlElem.append(optGr);
 }

}

/**
 * Extend Script Lib
 * */

/**
 * Get version of the design-layout in PR stamp
 * @return {Text} TextFrame[s] contents
 * */
function getVers() {
 var ad = activeDocument;
 var versElNames = ['__pr-stamp__versionNumber_mainTable__', '__pr-stamp__versionNumber_txtTable__'];
 for (var i = 0; i < versElNames.length; i++) {
  try {
   var currVersFrame = ad.textFrames.getByName(versElNames[i]);
   return currVersFrame.contents;
  } catch (e) {
   throw new Error('getVers Error' + ', ' + e.message);
  }
 }
}

/**
 * Change version of the design-layout in PR stamp
 * @param {Number} num - current version of the design-layout
 * */
function changeVers(num) {
 var ad = activeDocument;
 var num = +num;

 var versElNames = ['__pr-stamp__versionNumber_mainTable__', '__pr-stamp__versionNumber_txtTable__'];
 for (var i = 0; i < versElNames.length; i++) {
  _change(versElNames[i], num);
 }

  function _change(elName, num) {
  try {
   var versTextFrame = ad.textFrames.getByName(elName);

   var currNumb = versTextFrame.contents;
   // if (currNumb == '—') currNumb = '00';
   if (+currNumb < 0) currNumb = '00';

   var newVers = num;
   versTextFrame.contents = ('' + '0' + newVers).slice(-2);
  } catch (e) {
   throw new Error(e.line + ', ' + e.message);
  }
 }
}

/**
 * Set profile in the PR-stamp
 * @param {String} profileStr - profile name from the drop-down list
 * */
function setProfile(profileStr) {
 try {
  var profileFrame = activeDocument.textFrames.getByName('__pr-stamp__icmProfile__');
  profileFrame.contents = profileStr;
 } catch (e) {
  // alert(e.line + '. ' + e.message);
 }
}

/**
 * Set prepress-engineer in the PR-stamp
 * @param {Text} prepress-engineer name from the drop-down list
 * */
function setPrepresser(name) {
 try {
  var prepresser = activeDocument.textFrames.getByName('__pr-stamp__prepresser__');
  prepresser.contents = name;
 } catch (e) {
  // alert(e.line + '. ' + e.message);
 }
}

function getXmlStr() {

 if (!documents.length) {
  var errNoDocs = new Error('Нет активных документов!');
  alert(errNoDocs);
  return errNoDocs;
 }

 var a = ('' + activeDocument.fullName).slice(0, -3);
 var x = new File(a + '.xml');

 if (!x.exists) {
  var errNoXml = new Error('Не найден xml-файл!');
  alert(errNoXml);
  return errNoXml;
 }

 var s = '';
 x.open('r');
 s = x.read();
 x.close();

 return s;
}

function setXlsxData(o) {
 var ad = activeDocument;

 var stampFieldNames = [
  '__pr-stamp__orderNumber_mainTable__',
  '__pr-stamp__customerCompanyName_mainTable__',
  '__pr-stamp__orderName_mainTable__',

  '__pr-stamp__orderNumber_txtTable__',
  '__pr-stamp__customerCompanyName_txtTable__',
  '__pr-stamp__orderName_txtTable__',

  '__pr-stamp__sensorLabelSize__',
  '__pr-stamp__sensorLabelColor__',
  '__pr-stamp__sensorLabelFieldColor__',

  '__pr-stamp__inkNumber__',
  '__pr-stamp__inkChange__',

  // '__pr-stamp__icmProfile__', // профайл выбирается из списка вручную
  '__pr-stamp__printSide__',
  '__pr-stamp__filmComposition__',

  '__pr-stamp__formCylinder__',
  '__pr-stamp__printWidth__',
  '__pr-stamp__layoutWidth__',

  '__pr-stamp__windingSchema__',
 ];

 var stampWindingSchemaCodes = [
  '__pr-stamp__windingSchema_A__',
  '__pr-stamp__windingSchema_B__',
  '__pr-stamp__windingSchema_C__',
  '__pr-stamp__windingSchema_D__',
  '__pr-stamp__windingSchema_E__',
  '__pr-stamp__windingSchema_F__',
  '__pr-stamp__windingSchema_G__',
  '__pr-stamp__windingSchema_H__',
  '__pr-stamp__windingSchema_I__',
  '__pr-stamp__windingSchema_J__',
  '__pr-stamp__windingSchema_1__',
  '__pr-stamp__windingSchema_2__',
  '__pr-stamp__windingSchema_3__',
  '__pr-stamp__windingSchema_4__',
  '__pr-stamp__windingSchema_5__'
 ];
 setWindingSchema(stampWindingSchemaCodes, o);
 return push2obj(stampFieldNames, o);

 function setWindingSchema(stampWindingSchemaCodes, o) {
  var ad = activeDocument;
  // hide elements
  for (var i = 0; i <= stampWindingSchemaCodes.length; i++) {
   try {
    var elem = ad.pageItems.getByName(stampWindingSchemaCodes[i]);
    elem.opacity = 0;
   } catch (e) {
   }
  }

  var schemaLetter = o.windingSchema.match(/[A-J]/);
  var schemaNum = o.windingSchema.match(/[1-5]/);
  try {
   ad.pageItems.getByName('__pr-stamp__windingSchema_' + schemaLetter + '__').opacity = 100;
  } catch (e) {
  }

  switch (+schemaNum) {
   case 1:
   case 2:
   case 3:
    ad.pageItems.getByName('__pr-stamp__windingSchema_' + schemaNum + '__').opacity = 100;
    break;
   case 4:
    ad.pageItems.getByName('__pr-stamp__windingSchema_1__').opacity = 100;
    ad.pageItems.getByName('__pr-stamp__windingSchema_3__').opacity = 100;
    break;
   case 5:
   default:
    break;
  }

 }

 function push2obj(stampFieldNames, o) {
  var errors = 'Host error-list:\n\n';
  var ad = activeDocument;
  var blankField = '—';

  for (var i = 0; i <= stampFieldNames.length; i++) {
   var fieldName = stampFieldNames[i];

   switch (fieldName) {
    case '__pr-stamp__orderNumber_mainTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
     break;
    case '__pr-stamp__customerCompanyName_mainTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
     break;
    case  '__pr-stamp__orderName_mainTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = __getStreamNames(o[fieldName.slice(12, -12)]) || blankField;
     break;
     continue;

    case  '__pr-stamp__orderNumber_txtTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -11)] || blankField;
     break;
    case  '__pr-stamp__customerCompanyName_txtTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -11)] || blankField;
     break;
    case  '__pr-stamp__orderName_txtTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = __getStreamNames(o[fieldName.slice(12, -11)]) || blankField;
     break;

    case  '__pr-stamp__sensorLabelSize__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__sensorLabelColor__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__sensorLabelFieldColor__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;

    case  '__pr-stamp__inkNumber__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case   '__pr-stamp__inkChange__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    /** // профайл выбирается вручную из выпадающего списка
     case  '__pr-stamp__icmProfile__':
     try {
     var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
     alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;*/
    case  '__pr-stamp__printSide__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = blankField;
     if (o.printSideUp !== undefined) fieldElem.contents = 'Прямая';
     if (o.printSideDn !== undefined) fieldElem.contents = 'Обратная';
     break;
    case  '__pr-stamp__filmComposition__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     var filmPrint = '',
      filmCover = '',
      filmSeparator = '';
     if (o.filmPrint !== undefined) filmPrint = o.filmPrint;
     if (o.filmCover !== undefined) {
      filmCover = o.filmCover;
      filmSeparator = '+';
     }
     fieldElem.contents = (filmPrint + filmSeparator + filmCover) || blankField;
     break;

    case  '__pr-stamp__formCylinder__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__printWidth__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }

     fieldElem.contents = o.streamWidth * __getStreamNumb(o.streamsNumber) || blankField;
     break;
    case  '__pr-stamp__layoutWidth__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     if (o.supports !== undefined) {
      fieldElem.contents = (o.streamWidth * __getStreamNumb(o.streamsNumber) + 14) || blankField;
     } else {
      fieldElem.contents = o.streamWidth * o.streamsNumber || blankField;
     }
     break;

    case  '__pr-stamp__windingSchema__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      // alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    default:
     break;
   }
  }

  function __getStreamNames(arr) {
   var arr2 = [];
   for (var i = 0; i < arr.length; i++) {
    if (!arr[i]) continue;
    arr2.push(arr[i]);
   }
   return arr2.join(' + ');
  }

  function __getStreamNumb(arr) {
   var streamNumb = 0;
   for (var i = 0; i < arr.length; i++) {
    if (!arr[i]) continue;
    streamNumb += +arr[i];
   }
   return streamNumb;
  }

  /* for (var i = 0; i <= stampFieldNames.length; i++) {
   var fieldName = stampFieldNames[i];

   if (fieldName.match('mainTable')) {
   try {
   var fieldElem = ad.textFrames.getByName(fieldName);
   if(fieldName.match('orderName')) {
   fieldElem.contents = o[fieldName.slice(12, -12)].join('___') || blankField;
   alert(o[fieldName.slice(12, -12)].join('___'));
   } else {
   fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
   }

   } catch (e) {
   errors += e.line + '. ' + e.message + '\n';
   }
   } else if (fieldName.match('txtTable')) {
   try {
   var fieldElem = ad.textFrames.getByName(fieldName);
   if(fieldName.match('orderName')) {
   fieldElem.contents = o[fieldName.slice(12, -12)].join('___') || blankField;
   alert(o[fieldName.slice(12, -12)].join('___'));
   } else {
   fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
   }
   } catch (e) {
   errors += e.line + '. ' + e.message + '\n';
   }
   } else {
   try {
   var fieldElem = ad.textFrames.getByName(fieldName);

   switch (fieldName) {
   case '__pr-stamp__printSide__':
   fieldElem.contents = blankField;
   if (o.printSideUp !== undefined) fieldElem.contents = 'Прямая';
   if (o.printSideDn !== undefined) fieldElem.contents = 'Обратная';
   /!*       if (o.printSideUp === undefined) {
   fieldElem.contents = 'Прямая';
   } else if (o.printSideDn) {
   fieldElem.contents = 'Обратная';
   } else {
   fieldElem.contents = blankField;
   }*!/
   break;
   case '__pr-stamp__filmComposition__':
   var filmPrint = '',
   filmCover = '',
   filmSeparator = '';
   if (o.filmPrint !== undefined) filmPrint = o.filmPrint;
   if (o.filmCover !== undefined) {
   filmCover = o.filmCover;
   filmSeparator = '+';
   }
   fieldElem.contents = (filmPrint + filmSeparator + filmCover) || blankField;
   break;
   case '__pr-stamp__printWidth__':
   fieldElem.contents = o.streamWidth * o.streamsNumber || blankField;
   break;
   case '__pr-stamp__layoutWidth__':
   if (o.supports.match(/кресты/i)) {
   fieldElem.contents = (o.streamWidth * o.streamsNumber + 14) || blankField;
   } else {
   fieldElem.contents = o.streamWidth * o.streamsNumber || blankField;
   }
   break;
   default:
   fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
   break;
   }

   } catch (e) {
   errors += e.line + '. ' + e.message + '\n';
   }
   }

   }*/
  // alert(errors);
  return errors;
 }
}

function setInks(inksObj) {
 var TOTAL_INKS = 8;
 var ad = activeDocument;
 var prTableInks;

 try {
  prTableInks = ad.groupItems.getByName('__pr-stamp__inks__');

  for (var i = 0; i < TOTAL_INKS; i++) {
   var prTableInkGr = prTableInks.groupItems[i];
   var currInkData = inksObj[i];
   if (currInkData !== null) {
    prTableInkGr.opacity = 100;
    __setInkGr(prTableInkGr, currInkData, i);
   } else {
    prTableInkGr.opacity = 0;
   }
  }
 } catch (e) {
  // alert(e.line + '. ' + e.message);
  return;
 }

 function __setInkGr(gr, inkData, i) {
  try {
   var inkName = gr.pageItems.getByName('__ink-name__'),
    inkColor = gr.pageItems.getByName('__ink-color__'),
    inkPercent = gr.pageItems.getByName('__ink-percent__'),
    inkLpi = gr.pageItems.getByName('__ink-lpi__'),
    inkAngle = gr.pageItems.getByName('__ink-angle__'),
    inkNumber = gr.pageItems.getByName('__ink-number__');
   if (inkData[0] == 'PANTONE Process Blue C') {
    inkName.contents = 'PANTONE Process Blue';
   } else {
    inkName.contents = inkData[0].replace(/( C$)|(Process )/g, '');
   }

   inkPercent.contents = inkData[1];
   inkNumber.contents = i + 1;
   __setInkColor(inkColor, inkName, inkData);
  } catch (e) {
   // alert(e.line + '. ' + e.message);
   // return;
  }

  function __setInkColor(inkColor, inkName, inkData) {
   if (inkData[0].match('Process') && !inkData[0].match('PANTONE')) {
    var cmykColor = new CMYKColor();
    switch (inkData[0]) {
     case 'Process Cyan':
      cmykColor.cyan = 100;
      inkColor.fillColor = cmykColor;
      break;
     case 'Process Magenta':
      cmykColor.magenta = 100;
      inkColor.fillColor = cmykColor;
      break;
     case 'Process Yellow':
      cmykColor.yellow = 100;
      inkColor.fillColor = cmykColor;
      break;
     case 'Process Black':
      cmykColor.black = 100;
      inkColor.fillColor = cmykColor;
      break;
     default:
      throw new Error('Покраска квадратика в процесс-цвета: что-то пошло не так :_(');
    }
   } else {
    var col = activeDocument.swatches.getByName(inkData[0]);
    inkColor.fillColor = col.color;
   }
  }
 }
}