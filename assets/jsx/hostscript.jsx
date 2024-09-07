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
  '__pr-stamp__windingSchema__'
 ];

 return push2obj(stampFieldNames, o);

 function push2obj(stampFieldNames, o) {
  var errors = 'Host error-list:\n\n';
  var ad = activeDocument;

  for (var i = 0; i <= stampFieldNames.length; i++) {
   var fieldName = stampFieldNames[i];

   if (fieldName.match('_mainTable__')) {
    try {
     var fieldElem = ad.textFrames.getByName(fieldName);
     fieldElem.contents = o[fieldName.slice(12, -12)];
    } catch (e) {
     errors += e.line + '. ' + e.message + '\n';
    }
   } else if (fieldName.match('_txtTable__')) {
    try {
     var fieldElem = ad.textFrames.getByName(fieldName);
     fieldElem.contents = o[fieldName.slice(12, -11)];
    } catch (e) {
     errors += e.line + '. ' + e.message + '\n';
    }
   } else {
    try {
     var fieldElem = ad.textFrames.getByName(fieldName);

     switch (fieldName) {
      case '__pr-stamp__printSide__':
       if (o.printSideUp) {
        fieldElem.contents = 'Прямая бля!!!';
       } else if (o.printSideDn) {
        fieldElem.contents = 'Ебать, Обратная!!!';
       }
       break;
      case '__pr-stamp__filmComposition__':
       fieldElem.contents = o.filmPrint + '+' + o.filmCover || '';
       break;
      default:
       fieldElem.contents = o[fieldName.slice(12, -2)] || '—';
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