The the script was made using the Google Apps Script (GAS).
It gets an array of objects (calendar events) bypasses it in a loop, filters event by user-defined keywords, it then increments the counters when changing the following properties of objects: year (.getFullYear), month (.getMonth), day (.getDate).
Counts the total number of events in each year, the number of events in each month of the year, the average number of events in each month (eventsMonth / daysEvents), and displays the results in a table.
