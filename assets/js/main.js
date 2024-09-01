/**
 * При загрузке расширения высота панели подгонятся по ВЫСОТЕ контента.
 * Ограничения по габаритам панели заданы в файле CSXS/manifest.xml
 * */
try { fitPanelToContent();} catch (e) { }

/**
 * Скрипт следит за изменением цветовой схемы Иллюстратора
 * Описание скрипта в файле themeChanger.js
 * */
try { changeTheme(csInterface);} catch (e) { }

const reload_panel = document.querySelector('#reload_panel');
reload_panel.addEventListener('click', () => {
 location.reload();
});

getInkCoverage();
getXlsx();

function setXlsxData() {
 csInterface.evalScript('setXlsxData()', function (result) {
  alert(result);
 });
}

function getXlsx() {
 const input = document.querySelector('#get_xlsx');
 const output = document.querySelector("#output");
 let outStr = '';

 input.addEventListener("change", async e => {

  try {
   const file = e.target.files[0];
   const data = await file.arrayBuffer();

   const workBook = XLSX.read(data);
   const workSheet = workBook.Sheets[workBook.SheetNames[0]];

   const dataForTable = {
    customerCompanyName: workSheet.C4.v,
    orderNumber: workSheet.C5.v,
    orderName: workSheet.E6.v,
   };

   for (let key in dataForTable) {
    let val = dataForTable[key];
    outStr += val + '\n\n';
   }

   output.value = outStr;

   return dataForTable;

  } catch (e) {
   output.innerHTML = e;
  }

 });

}

function getInkCoverage() {
 const inkCoverageBtn = document.querySelector('#ink_coverage');
 const output = document.querySelector("#output");

 inkCoverageBtn.addEventListener('click', (e) => {
  try {
   csInterface.evalScript('getXmlStr();', function (result) {
    let xmlString = result;

    let xmlParser = new DOMParser();
    let xmlDoc = xmlParser.parseFromString(xmlString, 'text/xml');
    let outputString = '';

    let inks = xmlDoc.getElementsByTagName("Ink");

    for (let i = 0; i < inks.length; i++) {
     let currInk = inks[i];
     let currVal = currInk.getElementsByTagName('CoveragePerc')[0].getAttribute('Value');
     if (+currVal < 5) currVal = 5;
     outputString += currInk.getAttribute('Name') + ': ' + (+currVal).toFixed(2) + '%\n';
    }

    output.value = outputString.slice(0, -1);
    /* выделить, скопировать, снять выделение */
    output.select();
    document.execCommand('copy');
    output.value += ' '; // костыль
    output.value = output.value.slice(0, -1); // костыль
   });
  } catch (e) {alert(e); }
 });
}

`__pr-table__order-number__
__pr-table__customer-company-name__
__pr-table__order-name__

__pr-table__sensor-label-size__
__pr-table__sensor-label-color__
__pr-table__sensor-label-field-color__

__pr-table__icm-profile__
__pr-table__print-side__
__pr-table__material-composition__

__current_date_and_time__`;
`Номер
C5
Наименование продукта
E6
Заказчик
C4

тип печати
F23 - прямая
F24 - обратная

формный вал
E40
Раппорт
L42
Ширина ручья
L40
Количество ручьев
D6
Приводные элементы (если есть "кресты", то ширина монтажа + 14 или + 16)
L44
Ширина материала (в таблице не надо)
E42

Кол-во красок
E43
Меняется красок
L37

Размер метки
J49
Цвет поля метки
J50
Цвет метки
J51

Материал запечатываемый
F53
Матриал покрывной
F54

Схема намотки
L46
`;
