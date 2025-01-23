
// This constant is written in column G for rows for which an email
// has been sent successfully.
let EMAIL_SENT = 'Sent';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendNonDuplicateEmails() {
  try{
    // Get the active sheet in spreadsheet
    const sheet = SpreadsheetApp.getActiveSheet();
    let startRow = 3; // First row of data to process
    let numRows = 23; // Number of rows to process
    // Fetch the range of cells C2:E18
    const dataRange = sheet.getRange(startRow, 2, numRows, 7);
    // Fetch values for each row in the Range.
    const data = dataRange.getValues();
    for (let i = 0; i < data.length; ++i) {
      const row = data[i];
      const emailAddress = row[1]; // Fourth Column in our selection
      console.log(row)
      var email_body = 
      `Greetings,${row[0]},
      &nbsp;<p>Hope this email finds you well.</p>
      <p>Thank you for your great efforts during the past days. ❤ ❤ </p> 
      Kindly find your grades in linear algebra.
      
      
      <br/>
      <ul>
        <li><strong>Linear Algebra:</strong> ${row[2]}</li>
      </ul>
      &nbsp;&nbsp;&nbsp;&nbsp; 
      <br/> 
      <br/>
      <p>It seems impossible until it is done </p>
    Best of luck ^^ <br/>
    Yours,<br/>
     Noha Shehab <br/>`
      const emailSent = row[7]; 
      if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
        let subject = 'AI45-Linear Algebra Grade';
        MailApp.sendEmail( {to:emailAddress, name:'Noha Shehab', subject:subject, body:email_body,htmlBody:email_body});
        sheet.getRange(startRow + i, 8).setValue(EMAIL_SENT);
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      }
    }
  }
  catch(err){
    Logger.log(err)
  }
}
