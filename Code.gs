function remove_triggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
function make_trigger() {
  remove_triggers();
  ScriptApp.newTrigger('main')
    .timeBased()
    .atHour(16)
    .everyDays(1)
    .inTimezone('Asia/Singapore')
    .create();
}

function get_emails() {
  var ss = SpreadsheetApp.openById('SPREADSHEET_ID');
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange("B:B");
  var emails = [];
  for (var i = 2; range.getCell(i, 1).getValue() !== ''; i++) {
    emails.push(range.getCell(i, 1).getValue());
  }
  return emails;
}

function main() {
  var query = 'has:document subject:"HS Daily Bulletin"';
  query += ' subject:"'+(new Date().getFullYear())+'"';
  var threads = GmailApp.search(query);
  var thread = threads[0];
  var doc = thread.getMessages()[0];
  var url = doc.getPlainBody().split('\n')[3];
  
  var id = url.split('/')[5];
  Logger.log(id);
  var file = DriveApp.getFileById(id);
  var new_file = file.makeCopy('bot: '+file.getName(), DriveApp.getFolderById('FOLDER_ID'));
  var new_file_id = new_file.getId();
  var new_doc = DocumentApp.openById(new_file_id);

  Logger.log(new_doc.getName());
  Logger.log(new_doc.getBody().getImages());


  var body = new_doc.getBody();
  var numElems = body.getNumChildren();

  for (var childIndex=0; childIndex<numElems; childIndex++) {
    var child = body.getChild(childIndex);
    switch ( child.getType() ) {
      case DocumentApp.ElementType.PARAGRAPH:
        var container = child.asParagraph();
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        var container = child.asListItem();
        break;

      default:
        continue;
    }

    var imagesHere = container.getPositionedImages();
    for (var i = 0; i < imagesHere.length; i++) {
      if (imagesHere[i].getHeight() <= 20) {
        container.removePositionedImage(imagesHere[i].getId());
      }
    }
  }

  var emails = get_emails();
  for (var i = 0; i < emails.length; i++) {
    new_file.addViewer(emails[i]);
  }
}

