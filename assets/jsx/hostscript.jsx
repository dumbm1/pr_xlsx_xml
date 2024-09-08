function getXmlStr() {

 if (!documents.length) {
  var errNoDocs = new Error('Где, блять, документы???!!!');
  alert(errNoDocs);
  return errNoDocs;
 }

 var a = ('' + activeDocument.fullName).slice(0, -3);
 var x = new File(a + '.xml');

 if (!x.exists) {
  var errNoXml = new Error('Где, блять, xml-файл???!!!');
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
 ]
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

   if (fieldName.match('_mainTable__')) {
    try {
     var fieldElem = ad.textFrames.getByName(fieldName);
     fieldElem.contents = o[fieldName.slice(12, -12)] || blankField;
    } catch (e) {
     errors += e.line + '. ' + e.message + '\n';
    }
   } else if (fieldName.match('_txtTable__')) {
    try {
     var fieldElem = ad.textFrames.getByName(fieldName);
     fieldElem.contents = o[fieldName.slice(12, -11)] || blankField;
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
       /*       if (o.printSideUp === undefined) {
               fieldElem.contents = 'Прямая';
              } else if (o.printSideDn) {
               fieldElem.contents = 'Обратная';
              } else {
               fieldElem.contents = blankField;
              }*/
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

  }
  // alert(errors);
  return errors;
 }
}