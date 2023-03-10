function onOpen() {
  SpreadsheetApp
                .getUi()
                .createMenu("Mail Merge")
                .addItem("Start", "tableClassQuickstart")
                .addItem("Reset labels", "clear")
                .addToUi();
}

function clear () {
  let temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("people").getRange("F2:F");
  temp.clearContent();
  temp.clearFormat();
}

function tableClassQuickstart() {
    let sheetName = 'people';
    let headerRow = 1;
    let table = Sheetfu.getTable(sheetName, headerRow);       
    let item = table.items;
    let first_name, last_name, body, email, is_sent, subject;
    
    for (let i = 0; i < item.length; i++) {
        first_name = item[i].getFieldValue("first_name");
        last_name = item[i].getFieldValue("last_name");
        subject = item[i].getFieldValue("subject");
        body = item[i].getFieldValue("message");
        email = item[i].getFieldValue("email");
        is_sent = item[i].getFieldValue("is_sent");

        if (is_sent == "") {
            GmailApp.createDraft(email, subject, body);
            item[i].setFieldValue("is_sent", "done").commit(); 
            item[i].setFieldBackground("is_sent", "green").commit();
        } else if (is_sent == "done"){
            continue;
        } else {
            continue;
        } 
    }
}
