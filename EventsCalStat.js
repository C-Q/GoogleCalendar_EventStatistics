function myFunction() {
  
  // This version gets some range event objects, and returns arrays with:
  //   Work days in the each month
  //   Events in the each month
  //   Average amount per month
  
  var cal = CalendarApp.getCalendarById("kging7ubddkgvkgu6k12iuohi4@group.calendar.google.com");
  var events = cal.getEvents(new Date("October 1, 2018 00:00:00 CST"), new Date("December 28, 2018 00:00:00 CST"));
  
//  var titles = [];
//  titles.push(events[0].getTitle()); 
//  titles.push(events[events.length-1].getTitle());
//  Logger.log('First and last events titles: ' + '[ ' + titles + ' ]');
//  var startTime = (events[events.length-1].getStartTime().getHours());
//  var endTime = (events[events.length-1].getEndTime().getHours());
//  Logger.log('Start/end time last event in the range: ' + startTime + ' ' + endTime);
  
  var daysEachMonth = [];                 // Work days in the each month:  21, 23, 20
  var eventsEachMonth = [];               // Events in the each month:  85, 139, 82
  var averageEachMonth =[];               // Average amount per month:  4.05 ,6.04, 4.10
  var earlyDateVal = events[0].getStartTime().getDate();
  var earlyMonthVal = events[0].getStartTime().getMonth();
  var daysInMonth = 1;
  var eventsInMonth = 0;
  
  for (var i=0; i<events.length; i++) {
    var nextMonthVal = events[i].getStartTime().getMonth();
    if (nextMonthVal == earlyMonthVal){
      eventsInMonth ++;
      
      var nextDateVal = events[i].getStartTime().getDate();
      if(nextDateVal != earlyDateVal){
        daysInMonth ++;
        earlyDateVal = events[i].getStartTime().getDate();
      }
      
    } else {
      daysEachMonth.push(daysInMonth);
      daysInMonth = 1;
      eventsEachMonth.push(eventsInMonth);
      eventsInMonth = 1;
      earlyMonthVal = events[i].getStartTime().getMonth();
    }
  }
  daysEachMonth.push(daysInMonth);
  eventsEachMonth.push(eventsInMonth);
  
  for (var j=0; j<eventsEachMonth.length; j++){
    var averageInMonth = (eventsEachMonth[j] / daysEachMonth[j]).toFixed(2);
    averageEachMonth.push(averageInMonth);
  }
  
  Logger.log('Work days in the each month:  ' + daysEachMonth);
  Logger.log('Events in the each month:  ' + eventsEachMonth);
  Logger.log('Average amount per month:  ' + averageEachMonth);
}


