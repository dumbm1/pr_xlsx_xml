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

 let def = false;

 pasteInfTabKUBtn.addEventListener("click", function (e) {
  if (e.ctrlKey || e.metaKey) def = true;
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKUFileName + '","' + def + '")', function (result) {
   def = false;
  });
 });
 pasteInfTabKUBtn_flex.addEventListener("click", function (e) {
  if (e.ctrlKey || e.metaKey) def = true;
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKUFileNameFlex + '","' + def + '")', function (result) {
   def = false;
  });
 });
 pasteInfTabKfBtn.addEventListener("click", function (e) {
  if (e.ctrlKey || e.metaKey) def = true;
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKfFileName + '","' + def + '")', function (result) {
   def = false;
  });
 });
 pasteInfTabKfBtn_flex.addEventListener("click", function (e) {
  if (e.ctrlKey || e.metaKey) def = true;
  csInterface.evalScript(jsx_executeAiStampFile.toString() + ';jsx_executeAiStampFile("' + extPath + '","' + stampKfFileNameFlex + '","' + def + '")', function (result) {
   def = false;
  });
 });

 function jsx_executeAiStampFile(extensionPath, fileName, def) {
  var tmplFileStamp = new File(extensionPath + '/assets/files/' + fileName),
   targFileStamp;

  var targDefPath,
   targDefFolder,
   targDefFile;

  var targFile;

  if (def == 'true') {
   try {
    targDefPath = Folder.userData + '/LocalStore/kf_xlsx_xml/';
    targDefFolder = new Folder(targDefPath);

    if (!targDefFolder.exists) targDefFolder.create();

    targDefFile = new File(
     targDefFolder.fullName + '/' +
     (fileName).slice(0, -3) + '_copy@' +
     (new Date().getTime().toString()).slice(9) + '.ai');

    targFileStamp = _copyAiFile(tmplFileStamp, targDefFile);
    targFileStamp.execute();
    return targFileStamp;
   } catch (e) {
    return e;
   }
  }

  if (!documents.length ||
   (documents.length && !(new File(activeDocument.fullName).exists))) {
   try {
    targFile = File.openDialog();
    targFileStamp = _copyAiFile(tmplFileStamp, targFile);
    targFileStamp.execute();
    return targFileStamp;
   } catch (e) {
    return e;
   }
  }

  if (documents.length && new File(activeDocument.fullName).exists) {
   try {
    targFile = new File(activeDocument.fullName).openDlg();
    targFileStamp = _copyAiFile(tmplFileStamp, targFile);
    targFileStamp.execute();
    return targFileStamp;
   } catch (e) {
    return e;
   }
  }

  function _copyAiFile(tmplFile, targFile) {
   /**
    * соблюдать порядок обработки
    * */
   if (arguments.length != 2) throw new Error("_copyAiFile: ожидается два аргумента");
   if (!tmplFile.exists) throw new Error("_copyAiFile: tmplFile не существует");
   if (targFile.exists) throw new Error("_copyAiFile: targFile существует, есть опастность повредить существующий файл");
   if (targFile.lineFeed !== 'windows' && targFile.lineFeed !== 'unix' && targFile.lineFeed !== 'macintosh') {
    throw new Erorr("_copyAiFile: targFile не является объектом File"); // утиная типизация
   }
   if (tmplFile.name.slice(-3) != '.ai' || targFile.name.slice(-3) != '.ai') {
    throw new Error("_copyAiFile: поддерживается только формат Adobe Illustrator (.ai)");
   }

   var fileBinStr;

   try {
    tmplFile.encoding = 'BINARY';
    targFile.encoding = 'BINARY';

    tmplFile.open('r');
    fileBinStr = tmplFile.read();
    tmplFile.close();

    targFile.open('e');
    targFile.write(fileBinStr);
    targFile.close();

    return targFile;

   } catch (e) {
    if (e.message.match("_copyAiFile:")) return e;
    else throw (e);
   }

  }
 }

 /* function jsx_executeAiStampFile(extensionPath, fileName) {
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
  }*/

}
