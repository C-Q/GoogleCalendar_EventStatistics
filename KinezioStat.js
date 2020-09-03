
/* task:
  Скрипт получает массив объектов (событий календаря), обходит его в цикле, отфильтровывает события по
  заданным ключевым словам. Затем инкрементирует счетчики при изменении следующих свойств событий:
  год(.getFullYear), месяц(.getMonth), день(.getDate). Считает общее количество событий в каждом году,
  количество событий в каждом месяце года, среднее количество событий в каждом месяце (событийВмесяце / днейСсобытиями),
  и выводит результаты в таблицу.
done! */

function StartFunc() {
  
  var startSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('StartPage');
  var staffCalendarsID = startSheet.getRange(3,1,20,2).getValues(); // массив с фамилиями сотрудников и ID их календарей
  var filterKeywords2Dem = startSheet.getRange(3,3,20,1).getValues(); // массив массивов с ключевыми словами для фильтрации событий
  var filterKeywords = [];
  
  for (var i=0; i<filterKeywords2Dem.length; i++) { // метод .indexOf() работает только с одномерным массивом. дадим ему его:
    filterKeywords.push(filterKeywords2Dem[i][0])
  }
  
  for (var i=0; i<staffCalendarsID.length; i++) { // создаем вкладку на каждого сотрудника. если таковая уже есть, обновляем ее
    var sheetName = staffCalendarsID[i][0];
    var employeeCalID = staffCalendarsID[i][1];
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName));
    }
    if (sheetName == '') { break } // диапазон для staffCalendarsID был взят с запасом (костыль для отсечения лишних элементов массива)
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    StatisticsForEachYear(sheetName, employeeCalID, filterKeywords);
  }
}

function StatisticsForEachYear(sheetName, employeeCalID, filterKeywords) { 
  
  var cal = CalendarApp.getCalendarById(employeeCalID);
  var events = cal.getEvents(new Date("January 1, 2014 00:00:00"), new Date());
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.clearContents();
  sheet.setColumnWidths(1, 13, 85);
  
  var this_YearEvents = [[],[]]; // первый элемент - значение года, второй - количество событий в этом году
  var each_MonthEventAver = [[0,0,0,0,0,0,0,0,0,0,0,0,],[0,0,0,0,0,0,0,0,0,0,0,0,]];  // первый элемент - количество событий в каждом месяце, второй - среднее кол-во событий в каждом месяце
  
  // стартовые значения счетчиков:
  var yearEvents = 0;
  var monthEvents = 0;
  var monthDays = 1;
  var monthAverage = 0;
  var wasFiltered = 0
  
  // стартовые значения текущего года, месяца и дня: 
  var thisYear = events[0].getStartTime().getFullYear();
  var thisMonth = events[0].getStartTime().getMonth();
  var thisDay = events[0].getStartTime().getDate();
  
  // стартовые значения вертикальных отступов для цикличного изменения диапазонов:
  var indent_YearEvents = 2;
  var indent_MonthEventAver = 5;
  var indent_titlesYearPatients = 2;
  var indent_titlesMonths = 4;
  var indent_titlesData = 4;
  var indent_devideString = 1;
  var indent_emptySells = 2;
  
  // значения для ячеек оформления:
  var titlesYearPatients = [["Год",""],["Пациентов в году",""]];
  var titlesMonths = [['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь',]];
  var titlesData = [['Месяц'],['Всего_в_мес'],['Сред_в_мес']];
  
  // допПраметры для фильтрации событий-неПациентов
  var triggeredKeywords = [];
  var debugArr = [];
  
  // и поехали:
  for (i=0; i<events.length; i++) { 
    
    // фильтруем события, заголовки которых содержат ключевые слова:
    var evTitle = events[i].getTitle().toLowerCase(); // 
    var firstWordTitle = evTitle.split(" ")[0];
    if ( filterKeywords.indexOf(firstWordTitle) != -1 ) {
      wasFiltered ++;
      debugArr.push(evTitle);
      if (triggeredKeywords.indexOf(firstWordTitle) == -1) {
        triggeredKeywords.push(filterKeywords[filterKeywords.indexOf(firstWordTitle)]);
      }
      continue;
    }

    var nextYear = events[i].getStartTime().getFullYear();
    if (thisYear == nextYear) {
      yearEvents ++;
      
      var nextMonth = events[i].getStartTime().getMonth();
      if (thisMonth == nextMonth) {
        monthEvents ++;
        
        var nextDay = events[i].getStartTime().getDate();
        if (thisDay != nextDay) {
          monthDays ++;
          thisDay = events[i].getStartTime().getDate(); // обновляем значение дня
        }
      
      } else {
        monthAverage = (monthEvents/monthDays).toFixed(2);
        each_MonthEventAver[0][thisMonth] = monthEvents.toFixed();
        each_MonthEventAver[1][thisMonth] = monthAverage;
        thisMonth = events[i].getStartTime().getMonth(); // обновляем значение месяца
        monthEvents = 1;
        monthDays = 1;
      }
      
    } else {
      TableBuild(sheet,this_YearEvents,thisYear,yearEvents,monthAverage,monthEvents,monthDays,each_MonthEventAver,thisMonth,indent_devideString,indent_emptySells,indent_titlesYearPatients,indent_titlesMonths,indent_titlesData,indent_YearEvents,indent_MonthEventAver,titlesYearPatients,titlesMonths,titlesData);
      
      // инкрементируем вертикальные отступы для оформления таблицы следующего года:
      indent_devideString += 6;
      indent_emptySells += 6;
      indent_titlesYearPatients += 6;
      indent_titlesMonths += 6;
      indent_titlesData += 6;
      indent_YearEvents += 6;
      indent_MonthEventAver += 6;
      
      // готовим переменные цикла для прохода по следующему году:
      yearEvents = 1;
      monthEvents = 1;
      monthDays = 1;
      each_MonthEventAver = [[0,0,0,0,0,0,0,0,0,0,0,0,],[0,0,0,0,0,0,0,0,0,0,0,0,]];
      thisDay = events[i].getStartTime().getDate();
      thisMonth = events[i].getStartTime().getMonth();
      thisYear = events[i].getStartTime().getFullYear();
    }
  }
  TableBuild(sheet,this_YearEvents,thisYear,yearEvents,monthAverage,monthEvents,monthDays,each_MonthEventAver,thisMonth,indent_devideString,indent_emptySells,indent_titlesYearPatients,indent_titlesMonths,indent_titlesData,indent_YearEvents,indent_MonthEventAver,titlesYearPatients,titlesMonths,titlesData);
  var triggKeys = [];
  for (var i=0; i<triggeredKeywords.length; i++){
    triggKeys.push(triggeredKeywords[i]+'\n');
  }
  
Logger.log("Отфильтровано: " + wasFiltered.toFixed() + " событий");
Logger.log("По ключевым словам: " + triggKeys);
Logger.log(debugArr);
  
  // добавляем в таблицу количество отфильтрованных событий и ключевые слова по которым сработал фильтр:
  sheet.getRange(indent_MonthEventAver+4,1,1,2).merge();
  sheet.getRange(indent_MonthEventAver+4,1).setValues([['Отфильтровано событий:']]);
  sheet.getRange(indent_MonthEventAver+4,1,1,2).setHorizontalAlignments([['right','right']]);
  sheet.getRange(indent_MonthEventAver+4,3).setValues([[wasFiltered.toFixed()]]);
  sheet.getRange(indent_MonthEventAver+4,3).setHorizontalAlignments([['center']]);
  
  sheet.getRange(indent_MonthEventAver+5,1,1,2).merge();
  sheet.getRange(indent_MonthEventAver+5,1).setValues([['По ключам:']]);
  sheet.getRange(indent_MonthEventAver+5,1,1,2).setHorizontalAlignments([['right','right']]);
  sheet.getRange(indent_MonthEventAver+5,1,1,2).setVerticalAlignments([['top','top']]);
  sheet.getRange(indent_MonthEventAver+5,3).setValues([[triggKeys.join('')]]);
}


function TableBuild(sheet,this_YearEvents,thisYear,yearEvents,monthAverage,monthEvents,monthDays,each_MonthEventAver,thisMonth,indent_devideString,indent_emptySells,indent_titlesYearPatients,indent_titlesMonths,indent_titlesData,indent_YearEvents,indent_MonthEventAver,titlesYearPatients,titlesMonths,titlesData) {
  
  // вычисляем среднее и заполняем массив для вставки в диапазон результатов:
  this_YearEvents = [[thisYear.toFixed(),""],[yearEvents.toFixed(),""]];
  monthAverage = (monthEvents/monthDays);
  each_MonthEventAver[0][thisMonth] = monthEvents.toFixed();
  each_MonthEventAver[1][thisMonth] = monthAverage.toFixed(2);
  
  // Оформление таблицы:
    // получаем диапазоны:
  var range_YearBlock = sheet.getRange(indent_devideString,1,6,13); // блок года
  var range_devideString = sheet.getRange(indent_devideString,1,1,13); // разделительная строка
  var range_emptySells = sheet.getRange(indent_emptySells,5,2,9); // пустые ячейки
  var range_titlesYearPatients = sheet.getRange(indent_titlesYearPatients,1,2,2); // заголовок год и событий в этом году
  var range_titlesMonths = sheet.getRange(indent_titlesMonths,2,1,12); // названия месяцев
  var range_titlesData = sheet.getRange(indent_titlesData,1,3,1); // Месяц; Всего_в_мес: Сред_в_мес:
  var range_YearEvents = sheet.getRange(indent_YearEvents,3,2,2); // значение года и кол-во событий в этом году
  var range_MonthEventAver = sheet.getRange(indent_MonthEventAver,2,2,12); // собсна результаты расчетов
     // задаем им форматирование:
  range_YearBlock.setBorder(true,true,true,true,true,true); // разлиновываем всю таблицу и танцуем дальше:
  range_devideString.merge();
  range_devideString.setBackgrounds([['#999999','#999999','#999999','#999999','#999999','#999999','#999999','#999999','#999999','#999999','#999999','#999999','#999999']]);
  range_emptySells.merge();
  range_emptySells.setBackgrounds([['#efefef','#efefef','#efefef','#efefef','#efefef','#efefef','#efefef','#efefef','#efefef'],['#efefef','#efefef','#efefef','#efefef','#efefef','#efefef','#efefef','#efefef','#efefef']]);
  range_titlesYearPatients.mergeAcross();
  range_titlesYearPatients.setBackgrounds([['#cccccc','#cccccc'],['#cccccc','#cccccc']]);
  range_titlesYearPatients.setFontSizes([[11,11],[11,11]]);
  range_titlesYearPatients.setHorizontalAlignments([['right','right'],['right','right']]);
  range_titlesYearPatients.setValues(titlesYearPatients);
  range_titlesMonths.setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center']]);
  range_titlesMonths.setFontSizes([[11,11,11,11,11,11,11,11,11,11,11,11]]);
  range_titlesMonths.setBackgrounds([['#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc','#cccccc']]);
  range_titlesMonths.setValues(titlesMonths);
  range_titlesData.setHorizontalAlignments([['right'],['right'],['right']]);
  range_titlesData.setValues(titlesData);
  range_titlesData.setBackgrounds([['#cccccc'],['#cccccc'],['#cccccc']]);
//  range_titlesData.setFontSizes([[11],[11],[11]]);
  range_YearEvents.setHorizontalAlignments([['center','center'],['center','center']]);
  range_YearEvents.mergeAcross();
  range_YearEvents.setFontWeights([['bold','bold'],['bold','bold']]);
  range_MonthEventAver.setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center'],['center','center','center','center','center','center','center','center','center','center','center','center']]);
  range_MonthEventAver.setFontWeights([['bold','bold','bold','bold','bold','bold','bold','bold','bold','bold','bold','bold'],['bold','bold','bold','bold','bold','bold','bold','bold','bold','bold','bold','bold']]);
    // и вставляем результаты:
  range_YearEvents.setValues(this_YearEvents);
  range_MonthEventAver.setValues(each_MonthEventAver);
  
}
