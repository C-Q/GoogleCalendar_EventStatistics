function StatisticsForEachYear() {
  // Скрипт получает массив объектов (событий календаря), обходит его в цикле, инкрементируя счетчики по таким свойствам событий: год (.getFullYear), месяц (.getMonth), день (.getDate)
  
  var cal = CalendarApp.getCalendarById("kging7ubddkgvkgu6k12iuohi4@group.calendar.google.com");
  var events = cal.getEvents(new Date("February 1, 2017 00:00:00"), new Date());
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet4");
  sheet.clearContents();
  var sheetSourse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet3"); // шаблон оформления
  
  var range_YearEvents = sheet.getRange(2,3,2,1);                // получаем диапазоны для вывода значения года и общего количества событий в этом году,
  var range_MonthEventAver = sheet.getRange(5,2,2,12);           // и для вывода общего и среднего(в зависимости от количества дней с событиями) количества событий в каждом месяце
  
  var value_titlesYearPatients = sheetSourse.getRange(2,1,2,2).getValues(); // Заголовки "год" и "пациентов в году" array[[],[]];
  var value_titlesMonths = sheetSourse.getRange(4,2,1,12).getValues();      // названия месяцев array[[]];
  var value_titlesData = sheetSourse.getRange(4,1,3,1).getValues();         // заголовки строк "всего событий в месяце" и "в среднем событий в месяце" array[[],[]];

  
  var arrToPrint = [];
  var eachYearAll = [];
  var this_YearEvents = [[],[]]; // первый элемент - значение года, второй - количество событий в этом году
  var each_MonthEventAver = [[,,,,,,,,,,,,],[,,,,,,,,,,,,]];  // первый элемент - количество событий в мес (each_MonthEventAver[0] = each_MonthEventAver[0]), второй - среднее кол-во событий в мес (each_MonthEventAver[1] = each_MonthEventAver[1]).
  
  
  // стартовые значения счетчиков:
  var yearEvents = 0;
  var monthEvents = 0;
  var monthDays = 1;
  var monthAverage = 0;
    // стартовые значения вертикальных отступов для цикличного изменения диапазонов:
  var indent_YearEvents = 2;
  var indent_MonthEventAver = 5;
  var indent_titlesYearPatients = 2;
  var indent_titlesMonths = 4;
  var indent_titlesData = 4;
  var indent_devideString = 1;
  var indent_emptySells = 2;
  
  var thisYear = events[0].getStartTime().getFullYear();
  var thisMonth = events[0].getStartTime().getMonth();
  var thisDay = events[0].getStartTime().getDate();
  
  for (i=0; i<events.length; i++) {

    var nextYear = events[i].getStartTime().getFullYear();
    if (thisYear == nextYear) {
      yearEvents ++;
      
      var nextMonth = events[i].getStartTime().getMonth();
      if (thisMonth == nextMonth) {
        monthEvents ++;
        
        var nextDay = events[i].getStartTime().getDate();
        if (thisDay != nextDay) {
          monthDays ++;
          thisDay = events[i].getStartTime().getDate();
        }
      
      } else {
        monthAverage = (monthEvents/monthDays).toFixed(2);
        each_MonthEventAver[0][thisMonth] = monthEvents.toFixed();
        each_MonthEventAver[1][thisMonth] = monthAverage;
        monthEvents = 1;
        monthDays = 1;
        thisMonth = events[i].getStartTime().getMonth();
      }
      
    } else {
      this_YearEvents = [[thisYear.toFixed()],[yearEvents.toFixed()]];
      monthAverage = (monthEvents/monthDays).toFixed(2);
      each_MonthEventAver[0][thisMonth] = monthEvents.toFixed();
      each_MonthEventAver[1][thisMonth] = monthAverage;
      
      // оформление таблицы:
      sheet.getRange(indent_devideString,1,6,13).setBorder(true,true,true,true,true,true); // разлиновываем всю таблицу
      sheet.getRange(indent_devideString,1,1,13).mergeAcross();
      sheet.getRange(indent_titlesYearPatients,1,2,2).mergeAcross();
      sheet.getRange(indent_emptySells,5,2,9).merge();
      sheet.getRange(indent_titlesYearPatients,1,2,2).setHorizontalAlignments([['right','right'],['right','right']]);
      sheet.getRange(indent_YearEvents,3,2,2).mergeAcross();
      sheet.getRange(indent_YearEvents,3,2,2).setHorizontalAlignments([['center','center'],['center','center']]);
      sheet.getRange(indent_titlesYearPatients,1,2,2).setValues(value_titlesYearPatients);
      sheet.getRange(indent_titlesMonths,2,1,12).setValues(value_titlesMonths);
      sheet.getRange(indent_titlesMonths,2,1,12).setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center']]);
      sheet.getRange(indent_titlesData,1,3,1).setValues(value_titlesData);
      sheet.getRange(indent_titlesData,1,3,1).setHorizontalAlignments([['right'],['right'],['right']]);
      
      indent_devideString += 6;
      indent_emptySells += 6;
      indent_titlesYearPatients += 6;
      indent_titlesMonths += 6;
      indent_titlesData += 6;
      
      eachYearAll[0] = this_YearEvents;
      eachYearAll[1] = each_MonthEventAver;
      range_YearEvents = sheet.getRange(indent_YearEvents,3,2,1);
      range_MonthEventAver = sheet.getRange(indent_MonthEventAver,2,2,12);
      range_YearEvents.setValues(this_YearEvents);
      range_MonthEventAver.setValues(each_MonthEventAver);
      range_MonthEventAver.setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center'],['center','center','center','center','center','center','center','center','center','center','center','center']]);
      indent_YearEvents += 6;
      indent_MonthEventAver += 6;
      
      arrToPrint.push(eachYearAll);
      yearEvents = 1;
      monthEvents = 1;
      monthDays = 1;
      eachYearAll = [];
      each_MonthEventAver = [[,,,,,,,,,,,,],[,,,,,,,,,,,,]];
      thisDay = events[i].getStartTime().getDate();
      thisMonth = events[i].getStartTime().getMonth();
      thisYear = events[i].getStartTime().getFullYear();
    }
  }
  monthAverage = (monthEvents/monthDays).toFixed(2);
  each_MonthEventAver[0][thisMonth] = monthEvents;
  each_MonthEventAver[1][thisMonth] = monthAverage;
  this_YearEvents = [[thisYear],[yearEvents]];
  
  eachYearAll[0] = this_YearEvents;
  eachYearAll[1] = each_MonthEventAver;
//  indent_YearEvents -= 1;      // костыль:((
  range_YearEvents = sheet.getRange(indent_YearEvents,3,2,1);
  range_MonthEventAver = sheet.getRange(indent_MonthEventAver,2,2,12);
  range_YearEvents.clearContent();
  range_MonthEventAver.clearContent();
  range_YearEvents.setValues(this_YearEvents);
  range_MonthEventAver.setValues(each_MonthEventAver);
  
  // оформление таблицы:
  sheet.getRange(indent_devideString,1,6,13).setBorder(true,true,true,true,true,true);
  sheet.getRange(indent_devideString,1,1,13).mergeAcross();
  sheet.getRange(indent_titlesYearPatients,1,2,2).mergeAcross();
  sheet.getRange(indent_YearEvents,3,2,2).mergeAcross();
  sheet.getRange(indent_emptySells,5,2,9).merge();
  sheet.getRange(indent_YearEvents,3,2,2).setHorizontalAlignments([['center','center'],['center','center']]);
  sheet.getRange(indent_titlesYearPatients,1,2,2).setValues(value_titlesYearPatients);
  sheet.getRange(indent_titlesYearPatients,1,2,2).setHorizontalAlignments([['right','right'],['right','right']]);
  sheet.getRange(indent_titlesMonths,2,1,12).setValues(value_titlesMonths);
  sheet.getRange(indent_titlesMonths,2,1,12).setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center']]);
  sheet.getRange(indent_titlesData,1,3,1).setValues(value_titlesData);
  sheet.getRange(indent_titlesData,1,3,1).setHorizontalAlignments([['right'],['right'],['right']]);
  range_MonthEventAver.setHorizontalAlignments([['center','center','center','center','center','center','center','center','center','center','center','center'],['center','center','center','center','center','center','center','center','center','center','center','center']]);
  

  arrToPrint.push(eachYearAll);
Logger.log(arrToPrint);
  
}

