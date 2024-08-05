function run() {
    const startDate = new Date('Mon Apr 8 2024 00:00:00 GMT+0000')
    progressiveEmailAudit(startDate, 30)
}


function progressiveEmailAudit(startDate, weeksCount){


  console.log('staring from ', startDate)

  const sheet = SpreadsheetApp.getActiveSheet(); // Or get the specific sheet you want

  for (let i = 0; i < weeksCount; i++) { 

    const weekStart = new Date(startDate);
    weekStart.setDate(startDate.getDate() + (i * 7));

    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 7);


    console.log(`from  ${weekStart} , to ${weekEnd}`)

    const threads = GmailApp.search(`after:${weekStart.getFullYear()}/${weekStart.getMonth() + 1}/${weekStart.getDate()} before:${weekEnd.getFullYear()}/${weekEnd.getMonth() + 1}/${weekEnd.getDate()}`);

    let receivedCount = 0;
    let repliedCount = 0;

    threads.forEach(thread => {
      receivedCount++;
      if (haveReplied(thread)) { 
        repliedCount++;
      }
    });

    sheet.appendRow([
      weekStart.toDateString(), 
      receivedCount, 
      repliedCount,
      receivedCount - repliedCount
    ]);
  }

}


function haveReplied(thread){

  const currentUser = Session.getActiveUser().getEmail()

  const messages = thread.getMessages()
  for (let i = 0; i < messages.length; i++) {
    sender = parseEmail(messages[i].getFrom())

    if (sender == currentUser){
      return true
    }
    
  }

  return false
}


function parseEmail(str) {
  // Use a regular expression to match the email pattern
  const emailRegex = /<([^>]+)>/; 
  const match = str.match(emailRegex);

  if (match) {
    return match[1]; // Return the captured email address
  } else {
    return null; // Or handle the case where no email is found
  }
}