//@include json2.js

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
 var prStampNames = [
  '__pr-stamp__orderNumber__',
  '__pr-stamp__customerCompanyName__',
  '__pr-stamp__orderName__',
  '__pr-stamp__sensorLabelSize__',
  '__pr-stamp__sensorLabelColor__',
  '__pr-stamp__sensorLabelFieldColor__',
  '__pr-stamp__icmProfile__',
  '__pr-stamp__printSide__',
  '__pr-stamp__filmComposition__',
  '__pr-stamp__incNumber__',
  '__pr-stamp__incChange__',
  '__pr-stamp__windingSchema__'
 ];

 function push2obj(obj) {
  var errors = '';
  var ad = activeDocument;
  for (var i = 0; i < prStampNames.length; i++) {
   var currStampName = prStampNames[i];
   try {
    var currStampElem = ad.TextFrames.getByName(currStampName);
    currStampElem.contents = o[currStampName.slice(12, -2)];
   } catch (e) {
    errors += '\n';
   }
  }
  alert(errors);
 }
}

`
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
`;



