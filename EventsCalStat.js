
  /* Скрипт получает массив объектов (событий календаря), обходит его в цикле, инкрементируя 
  счетчики при изменении следующих свойств событий: год(.getFullYear), месяц(.getMonth), день(.getDate).
  Считает общее количество событий в каждом году, количество событий в каждом месяце года, 
  событий в среднем в каждом месяце (событийВмесяце / днейСсобытиями), и выводит результаты в таблицу. */

function AddEmployees() {
  
  var staffCalendarsID = {
    Лях: "kging7ubddkgvkgu6k12iuohi4@group.calendar.google.com",
//    Величко: "ruvajop3dg216vioupdmqmqh3o@group.calendar.google.com",
//    Валуева: "uej0ulbenqucdp2b1joc8vavh0@group.calendar.google.com",
  }
  
  for (var employee in staffCalendarsID) {
    var id = staffCalendarsID[employee];
    StatisticsForEachYear(employee, id);
  }
}

function StatisticsForEachYear(employee, id) { 
  
  var cal = CalendarApp.getCalendarById(id);
  var events = cal.getEvents(new Date("January 1, 2016 00:00:00"), new Date());
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(employee);
  sheet.clearContents();
  sheet.setColumnWidths(1, 13, 85);
  
  var this_YearEvents = [[],[]]; // первый элемент - значение года, второй - количество событий в этом году
  var each_MonthEventAver = [[0,0,0,0,0,0,0,0,0,0,0,0,],[0,0,0,0,0,0,0,0,0,0,0,0,]];  // первый элемент - количество событий в мес (each_MonthEventAver[0] = each_MonthEventAver[0]), второй - среднее кол-во событий в мес (each_MonthEventAver[1] = each_MonthEventAver[1]).
  
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
  
  // параметры для фильтрации событий-неПациентов:
  var filterKeywords = ['занят', 'уехал']; // !!! wasFiltered наматывает все события, но сцуко реально фильтруются только по последнему элементу массива
  var triggeredKeywords = [];
  
  
  // и поехали:
  for (i=0; i<events.length; i++) {
    
    // фильтруем события заголовки которых содержат ключевые слова:
    var evTitle =( events[i].getTitle() ).toLowerCase();
    for (var f=0; f < filterKeywords.length; f++) {
      if ( evTitle.indexOf(filterKeywords[f]) != -1 ) {
        wasFiltered ++;
        var stopIter = true;
        if (triggeredKeywords.indexOf(filterKeywords[f]) == -1) {
          triggeredKeywords.push(filterKeywords[f]);
        }
      } else {stopIter = false;}
    }
    if (stopIter == true) continue;

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
  
Logger.log(wasFiltered.toFixed());
Logger.log(triggeredKeywords);
}


function TableBuild(sheet,this_YearEvents,thisYear,yearEvents,monthAverage,monthEvents,monthDays,each_MonthEventAver,thisMonth,indent_devideString,indent_emptySells,indent_titlesYearPatients,indent_titlesMonths,indent_titlesData,indent_YearEvents,indent_MonthEventAver,titlesYearPatients,titlesMonths,titlesData) {
  
  // вычисляем среднее и заполняем массив для вставки в диапазон результатов:
  this_YearEvents = [[thisYear.toFixed(),""],[yearEvents.toFixed(),""]];
  monthAverage = (monthEvents/monthDays);
  each_MonthEventAver[0][thisMonth] = monthEvents.toFixed();
  each_MonthEventAver[1][thisMonth] = monthAverage.toFixed(2);
  
  // Оформление таблицы:
    // получаем диапазоны:
  var range_YearBlock = sheet.getRange(indent_devideString,1,6,13);
  var range_devideString = sheet.getRange(indent_devideString,1,1,13);
  var range_emptySells = sheet.getRange(indent_emptySells,5,2,9);
  var range_titlesYearPatients = sheet.getRange(indent_titlesYearPatients,1,2,2);
  var range_titlesMonths = sheet.getRange(indent_titlesMonths,2,1,12);
  var range_titlesData = sheet.getRange(indent_titlesData,1,3,1);
  var range_YearEvents = sheet.getRange(indent_YearEvents,3,2,2);
  var range_MonthEventAver = sheet.getRange(indent_MonthEventAver,2,2,12);
     // задаем им форматирование:
  range_YearBlock.setBorder(true,true,true,true,true,true); // разлиновываем всю таблицу, затем:
  range_devideString.merge();
  range_emptySells.merge();
  range_titlesYearPatients.mergeAcross();
  range_titlesYearPatients.setHorizontalAlignments([['right','right'],['right','right']]);
  range_titlesYearPatients.setValues(titlesYearPatients);
  range_titlesMonths.setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center']]);
  range_titlesMonths.setValues(titlesMonths);
  range_titlesData.setHorizontalAlignments([['right'],['right'],['right']]);
  range_titlesData.setValues(titlesData);
  range_YearEvents.setHorizontalAlignments([['center','center'],['center','center']]);
  range_YearEvents.mergeAcross();
  range_MonthEventAver.setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center'],['center','center','center','center','center','center','center','center','center','center','center','center']]);
    // и вставляем результаты:
  range_YearEvents.setValues(this_YearEvents);
  range_MonthEventAver.setValues(each_MonthEventAver);
  
}