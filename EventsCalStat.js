function StatisticsForEachYear() {
  // добавляем запись результатов в диапазоны таблицы
  
  var cal = CalendarApp.getCalendarById("kging7ubddkgvkgu6k12iuohi4@group.calendar.google.com");
  var events = cal.getEvents(new Date("February 1, 2017 00:00:00"), new Date());
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet4");
  
  var range_YearEvents = sheet.getRange(2,3,2,1);                // получаем диапазоны для вывода значения года и общего количества событий в этом году,
  var range_MonthEventAver = sheet.getRange(5,2,2,12);           // и для вывода общего и среднего(в зависимости от количества дней с событиями) количества событий в каждом месяце
  
  var arrToPrint = [];
  var eachYearAll = [];
  var this_YearEvents = [[],[]]; // первый элемент - значение года, второй - количество событий в этом году
  var each_MonthEventAver = [[,,,,,,,,,,,,],[,,,,,,,,,,,,]];  // первый элемент - количество событий в мес (each_MonthEventAver[0] = each_MonthEventAver[0]), второй - среднее кол-во событий в мес (each_MonthEventAver[1] = each_MonthEventAver[1]).
  
//  var enevtDescription = [];                        // for debuggin
  
  var yearEvents = 0;
  var monthEvents = 0;
  var monthDays = 1;
  var monthAverage = 0;
  var indent_YearEvents = 2;
  var indent_MonthEventAver = 5;
  
  var thisYear = events[0].getStartTime().getFullYear();
  var thisMonth = events[0].getStartTime().getMonth();
  var thisDay = events[0].getStartTime().getDate();
  
  for (i=0; i<events.length; i++) {
//    enevtDescription.push( events[i].getTitle() + ' | Date-' + events[i].getStartTime().getDate() + '; startTime-' +  events[i].getStartTime().getHours() + ':00 ' );
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
        each_MonthEventAver[0][thisMonth] = monthEvents;
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
      
      eachYearAll[0] = this_YearEvents;
      eachYearAll[1] = each_MonthEventAver;
      range_YearEvents = sheet.getRange(indent_YearEvents,3,2,1);
      range_MonthEventAver = sheet.getRange(indent_MonthEventAver,2,2,12);
      range_YearEvents.setValues(this_YearEvents);
      range_MonthEventAver.setValues(each_MonthEventAver);
      indent_YearEvents += 6;
      indent_MonthEventAver += 5;
      
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
  range_YearEvents = sheet.getRange(indent_YearEvents,3,2,1);
  range_MonthEventAver = sheet.getRange(indent_MonthEventAver,2,2,12);
  range_YearEvents.setValues(this_YearEvents);
  range_MonthEventAver.setValues(each_MonthEventAver);
  indent_YearEvents += 6;
  indent_MonthEventAver += 5;

  arrToPrint.push(eachYearAll);
Logger.log(arrToPrint);
  
}

