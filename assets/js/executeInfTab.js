executeInfTab();

function executeInfTab() {
 let stampKUFileName = "stamp_ku_v0-1-offset.ai";
 let stampKUFileNameFlex = "stamp_ku_v0-1-flex.ai";
 let stampKfFileName = "stamp_kf_v0-1-offset.ai";
 let stampKfFileNameFlex = "stamp_kf_v0-1-flex.ai";
 const pasteInfTabKUBtn = document.getElementById("paste_inf_tab_ku_btn");
 const pasteInfTabKUBtn_flex = document.getElementById("paste_inf_tab_ku_btn_flex");
 const pasteInfTabKfBtn = document.getElementById("paste_inf_tab_kf_btn");
 const pasteInfTabKfBtn_flex = document.getElementById("paste_inf_tab_kf_btn_flex");
 const extPath = csInterface.getSystemPath(SystemPath.EXTENSION);

 pasteInfTabKUBtn.addEventListener("click", function () {
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKUFileName + '")', function (result) {
  });
 });
 pasteInfTabKUBtn_flex.addEventListener("click", function () {
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKUFileNameFlex + '")', function (result) {
  });
 });
 pasteInfTabKfBtn.addEventListener("click", function () {
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKfFileName + '")', function (result) {
  });
 });
 pasteInfTabKfBtn_flex.addEventListener("click", function () {
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKfFileNameFlex + '")', function (result) {
  });
 });

 function jsx_executeAiStampFile(extensionPath, fileName) {
  var aiStampFile = new File(extensionPath + '/assets/files/' + fileName);
  var storePath = Folder.userData + '/LocalStore/kf_xlsx_xml/';
  var store = new Folder(storePath);

  if (!store.exists) store.create();

  try {
   var aiStampFileCopy = _copyFile(aiStampFile, store);
   aiStampFileCopy.execute();
   return aiStampFileCopy;
  } catch (e) {
   return new Error(e);
  }

  function _copyFile(file, folder) {
   if (arguments.length != 2) return new Error('Invalid input-number of arguments');
   if (!file.exists) return new Error("Moving file doesn't exists");
   if (!folder.getFiles) return new Error("Folder doesn't passed");

   if (!folder.exists) folder.create();
   if (!folder.exists) return new Error("Folder cannot be created");

   var fileCopy, fileBinStr, files;

   fileCopy = new File(folder.fullName + '/' + (file.name).slice(0, -3) + '_copy@' + (new Date().getTime().toString()).slice(9) + '.ai');

   file.encoding = 'BINARY';
   fileCopy.encoding = 'BINARY';

   file.open('r');
   fileBinStr = file.read();
   file.close();

   fileCopy.open('e');
   fileCopy.write(fileBinStr);
   fileCopy.close();

   return fileCopy;
  }
 }

}
