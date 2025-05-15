function updateCapacityFromCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() + 1); 
  const includedCalendars = ["Accomodation", "COURSE", "EVENT", "EXAM!!!!!", "Exercise", "JOB", "Transport", "Travel"];
  const allCalendars = CalendarApp.getAllCalendars();

  const calendars = allCalendars.filter(cal => includedCalendars.includes(cal.getName()));

  const defaultCapacity = 5600;
  const penaltyPerHour = 700;

  sheet.getRange("A2:C1000").clearContent(); // clear old content from row 2 onwards

  for (let i = 0; i < 14; i++) {
    const currentDate = new Date(startDate);
    currentDate.setDate(currentDate.getDate() + i);

    const dayStart = new Date(currentDate);
    dayStart.setHours(0, 0, 0, 0);
    const dayEnd = new Date(currentDate);
    dayEnd.setHours(23, 59, 59, 999);

    let totalEventHours = 0;

    calendars.forEach(calendar => {
      const events = calendar.getEvents(dayStart, dayEnd);
      events.forEach(event => {
        const durationInHours = (event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60);
        totalEventHours += durationInHours;
      });
    });

    const capacity = Math.max(defaultCapacity - Math.floor(totalEventHours * penaltyPerHour), 0);
    const status = capacity > 0 ? "Available" : "Unavailable";

    const row = i + 2;
    sheet.getRange(row, 1).setValue(currentDate).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(row, 2).setValue(capacity);
    sheet.getRange(row, 3).setValue(status);

    if (status === "Unavailable") {
      sheet.getRange(row, 1, 1, 3).setBackground("#d9d9d9"); // Light Grey 1
    } else {
      sheet.getRange(row, 1, 1, 3).setBackground(null); // Clear background if not Unavailable
    }
  }
}
