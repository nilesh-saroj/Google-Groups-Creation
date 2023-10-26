function removedMembersNotInShet() {
  var sheet = SpreadsheetApp.openById('1yxKOdJF5HqaPUF2Ut_LB3HgTQ7YN7UnUBnkwhLGCVUU').getSheetByName("Port Groups");
  var data = sheet.getDataRange().getValues();
  

  var removedMembersSheet = SpreadsheetApp.openById("1yxKOdJF5HqaPUF2Ut_LB3HgTQ7YN7UnUBnkwhLGCVUU").getSheetByName("Removed Members");

  var lastDataRow = sheet.getLastRow();

  var columnAValues = sheet.getRange(2, 1, lastDataRow - 1, 1).getValues();
  var lastRowWithData = 1;
  for (var i = columnAValues.length - 1; i >= 0; i--) {
    if (columnAValues[i][0] !== "") {
      lastRowWithData = i + 2;
      break;
    }
  }

  for (var i = 1; i < lastRowWithData; i++) {
    var portname = data[i][1]; // Port name is in the second column of the sheet.
    var groupEmail = data[i][3]; // group email is in the fourth column of the sheet.
    var memberEmail = data[i][2]; // member email is in the third column of the sheet


    // Member Removal
    var group = GroupsApp.getGroupByEmail(groupEmail);
    var groupMembers = group.getUsers();
    var sheetData = sheet.getDataRange().getValues();
    
    for (var k = 0; k < groupMembers.length; k++) {
      var memberEmail = groupMembers[k].getEmail();
      var isMemberIntheSheet = false;
      
      // Check if member is in the sheet
      for (var j = 0; j < sheetData.length; j++) {
        if (memberEmail === sheetData[j][2] && group.getEmail() === sheetData[j][3]) {
          isMemberIntheSheet = true;
          break;
        }
      }
      
      // If member is not in the sheet, remove them from the group
      if (!isMemberIntheSheet) {
        removeMemberFromGroup(groupEmail, memberEmail);
        removedMembersSheet.appendRow([portname, groupEmail, memberEmail, 'Member removed']);
        Logger.log('Removed member: ' + memberEmail);
      }
    }
  }
}


function removeMemberFromGroup(groupEmail, memberEmail) {
  var url = `https://www.googleapis.com/admin/directory/v1/groups/${encodeURIComponent(groupEmail)}/members/${encodeURIComponent(memberEmail)}`;
  var options = {
    method: 'delete',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(`Removed member ${memberEmail} from group ${groupEmail}.`);
  Logger.log(response.getContentText());
}
