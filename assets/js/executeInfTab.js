executeInfTab();

function executeInfTab() {
 let stampPrFileName = "stamp_pr_v3-1.ai";
 let stampKfFileName = "stamp_kf_v0-1.ai";
 const pasteInfTabPrBtn = document.getElementById("paste_inf_tab_pr_btn");
 const pasteInfTabKfBtn = document.getElementById("paste_inf_tab_kf_btn");
 const extPath = csInterface.getSystemPath(SystemPath.EXTENSION);

 pasteInfTabPrBtn.addEventListener("click", function () {
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampPrFileName + '")', function (result) {
  });
 });
 pasteInfTabKfBtn.addEventListener("click", function () {
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKfFileName + '")', function (result) {
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
