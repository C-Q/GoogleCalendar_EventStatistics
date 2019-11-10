function myFunction() {
  // это хуй пойми шо, сохранено для истории )))))
  var cal = CalendarApp.getCalendarById("kging7ubddkgvkgu6k12iuohi4@group.calendar.google.com");
  var events = cal.getEvents(new Date("December 28, 2018 14:00:00"), new Date("January 3, 2019 13:00:00"));
  
//  var titles = [];
//  titles.push(events[0].getTitle()); 
//  titles.push(events[events.length-1].getTitle());
//  Logger.log('First and last events titles: ' + '[ ' + titles + ' ]');
//  var startTime = (events[events.length-1].getStartTime().getHours());
//  var endTime = (events[events.length-1].getEndTime().getHours());
//  Logger.log('Start/end time last event in the range: ' + startTime + ' ' + endTime);
  
  var daysEachMonth = [];                   // Work days in the each month
  var eventsEachMonth = [[]];               // Events in the each month
  var averageEachMonth =[[]];               // Average amount per month
  var year_totalEveryYear = [[],[]];        // box for this yearValue[0], and amount of all events in this year[1]
  var arrYear = [];                         // result array for writing in the sheet
  var enevtDescr = [];                        // for debugging
  
  
  var earlyDateVal = events[0].getStartTime().getDate();
  var earlyMonthVal = events[0].getStartTime().getMonth();
  var earlyYearVal = events[0].getStartTime().getFullYear();
  var daysInMonth = 1;
  var eventsInMonth = 0;
  var eventsInYear = 0;
  
  for (var i=0; i<events.length; i++) {
    enevtDescr.push( events[i].getTitle() + ' | Date-' + events[i].getStartTime().getDate() + '; startTime-' +  events[i].getStartTime().getHours() + ':00 ' );
    var nextYearVal = events[i].getStartTime().getFullYear();
    if (nextYearVal == earlyYearVal){
      eventsInYear ++;
    
      var nextMonthVal = events[i].getStartTime().getMonth();
      if (nextMonthVal == earlyMonthVal){
        eventsInMonth ++;
        
        var nextDateVal = events[i].getStartTime().getDate();
        if(nextDateVal != earlyDateVal){
          daysInMonth ++;
          earlyDateVal = events[i].getStartTime().getDate();
        };
      } else {
        daysEachMonth.push(daysInMonth);
        daysInMonth = 1;
//        eventsEachMonth[0].push(eventsInMonth);
        eventsInMonth = 1;
        earlyMonthVal = events[i].getStartTime().getMonth();
      };
    } else {
      year_totalEveryYear[0][0] = earlyYearVal;
      year_totalEveryYear[1][0] = eventsInYear;
      eventsInYear = 0;
      eventsEachMonth[0].push(eventsInMonth);
      daysEachMonth.push(daysInMonth);
      earlyYearVal = events[i].getStartTime().getFullYear();
      daysInMonth = 1;
      eventsInMonth = 1;
      for (var j=0; j<daysEachMonth.length; j++){
        var averageInMonth = (eventsEachMonth[0][j] / daysEachMonth[j]).toFixed(2);
        averageEachMonth[0].push(averageInMonth);
      };
      arrYear.push([year_totalEveryYear, eventsEachMonth, averageEachMonth]);
      daysEachMonth = [];
      eventsEachMonth = [[]];
      year_totalEveryYear = [[],[]];
      
    };
  };
  daysEachMonth.push(daysInMonth);
  eventsEachMonth[0].push(eventsInMonth);
  year_totalEveryYear[0][0] = nextYearVal;
  year_totalEveryYear[1][0] = eventsInYear;
  eventsInYear = 0;
  for (var j=0; j<daysEachMonth.length; j++){
    var averageInMonth = (eventsEachMonth[0][j] / daysEachMonth[j]).toFixed(2);
    averageEachMonth[0].push(averageInMonth);
  };
  arrYear.push([year_totalEveryYear, eventsEachMonth, averageEachMonth]);
//  daysEachMonth = [];
//  eventsEachMonth = [[]];
//  year_totalEveryYear = [[],[]];
  Logger.log('Result array:  ' + arrYear[0]);
  Logger.log('Result array:  ' + arrYear[1]);
  Logger.log('Events title:  ' + titleArr);
}
// Расставить точки останова и отдебажить пособытийно

