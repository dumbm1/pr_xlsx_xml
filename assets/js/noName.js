function noName() {
 var ad = activeDocument;
 var arr = [];
 var currOpacity;
 var newOpacity;

 try {
  var prLogoGr = ad.groupItems.getByName('__company-logo__');
  arr.push(prLogoGr);
 } catch (e) {
 }

 try {
  var prLogoTtGr = ad.groupItems.getByName('__company-logo-tt__');
  arr.push(prLogoTtGr);
 } catch (e) {
 }

 try {
  var clientNameGr = ad.groupItems.getByName('__client-name__');
  arr.push(clientNameGr);
 } catch (e) {
 }

 try {
  var clientNameTtGr = ad.groupItems.getByName('__client-name-tt__');
  arr.push(clientNameTtGr);
 } catch (e) {
 }

 if (arr.length == 0) return;

 currOpacity = arr[0].opacity;
 if (currOpacity === 0) {
  newOpacity = 100;
 } else {
  newOpacity = 0;
 }

 for (var i = 0; i < arr.length; i++) {
  var elem = arr[i];
  elem.opacity = newOpacity;
 }
}