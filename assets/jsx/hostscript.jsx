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

function setXlsxData(){
 return 'Hello, no data there. Coming soon...'
}
