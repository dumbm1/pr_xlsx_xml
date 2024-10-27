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

  '__pr-stamp__icmProfile__',
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
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
     break;
    case '__pr-stamp__customerCompanyName_mainTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
     break;
    case  '__pr-stamp__orderName_mainTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = __getStreamNames(o[fieldName.slice(12, -12)]) || blankField;
     break;
     continue;

    case  '__pr-stamp__orderNumber_txtTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -11)] || blankField;
     break;
    case  '__pr-stamp__customerCompanyName_txtTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -11)] || blankField;
     break;
    case  '__pr-stamp__orderName_txtTable__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = __getStreamNames(o[fieldName.slice(12, -11)]) || blankField;
     break;

    case  '__pr-stamp__sensorLabelSize__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__sensorLabelColor__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__sensorLabelFieldColor__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;

    case  '__pr-stamp__inkNumber__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case   '__pr-stamp__inkChange__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;

    case  '__pr-stamp__icmProfile__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__printSide__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = blankField;
     if (o.printSideUp !== undefined) fieldElem.contents = 'Прямая';
     if (o.printSideDn !== undefined) fieldElem.contents = 'Обратная';
     break;
    case  '__pr-stamp__filmComposition__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
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
      alert(e.line + '. ' + e.message);
     }
     fieldElem.contents = o[fieldName.slice(12, -2)] || blankField;
     break;
    case  '__pr-stamp__printWidth__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
     }

     fieldElem.contents = o.streamWidth * __getStreamNumb(o.streamsNumber) || blankField;
     break;
    case  '__pr-stamp__layoutWidth__':
     try {
      var fieldElem = ad.textFrames.getByName(fieldName);
     } catch (e) {
      alert(e.line + '. ' + e.message);
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
      alert(e.line + '. ' + e.message);
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

function increase_vers() {
 var ad = activeDocument;
 var versNumbMainStr = '__pr-stamp__versionNumber_mainTable__';
 var versNumbTxtStr = '__pr-stamp__versionNumber_txtTable__';

 try {
  var versNumbMainFrame = ad.textFrames.getByName(versNumbMainStr);
  var versNumbTxtFrame = ad.textFrames.getByName(versNumbTxtStr);
  var currNumb = versNumbMainFrame.contents;
  var changeVers = confirm('Увеличить версию?');
  if (changeVers) {
   var newVers = +currNumb + 1;
   versNumbMainFrame.contents = ('' + '0' + newVers).slice(-2);
   versNumbTxtFrame.contents = ('' + '0' + newVers).slice(-2);
  }

 } catch (e) {
  alert(e.line + ', ' + e.message);
 }
}

function setProfile(profileStr) {
 try {
  var profileFrame = activeDocument.textFrames.getByName('__pr-stamp__icmProfile__');
  profileFrame.contents = profileStr;
 } catch (e) {
  alert(e.line + '. ' + e.message);
 }
}

function setInks(inksObj) {
 var TOTAL_INKS = 8;
 var ad = activeDocument;
 try {
  var prTableInks = ad.groupItems.getByName('__pr-stamp__inks__');
  for (var i = 0; i < TOTAL_INKS; i++) {
   var prTableInkGr = prTableInks.groupItems[i];
   var currInkData = inksObj[i];
   if (currInkData !== null) {
    prTableInkGr.hidden = false;
    __setInkGr(prTableInkGr, currInkData, i);
   } else {
    prTableInkGr.hidden = true;
   }
  }
 } catch (e) {
  alert(e.line + '. ' + e.message);
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
   inkName.contents = inkData[0];
   inkPercent.contents = inkData[1];
   inkNumber.contents = i + 1;
__setInkColor(inkColor, inkName, inkData);
  } catch (e) {
   alert (e.line + '. ' + e.message);
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