function synchronizeGroupMembers() {
  var sheet = SpreadsheetApp.openById('1yxKOdJF5HqaPUF2Ut_LB3HgTQ7YN7UnUBnkwhLGCVUU').getSheetByName("Port Groups");
  var data = sheet.getDataRange().getValues();
  var statusColumnIndex = 5; // Column index for Member Status (5th column)

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
  
    // Get current members of the group
    var currentMembers = getCurrentGroupMembers(groupEmail);

    // Check if member is in the sheet but not in the group, add to the group
    if (currentMembers.indexOf(memberEmail) === -1) {
      addMemberToGroup(groupEmail, memberEmail);  
      sheet.getRange(i + 1, statusColumnIndex).setValue('Member added Successfully');
    }
    // Check if member is in the group but not in the sheet, remove from the group
    else if (currentMembers.indexOf(memberEmail) !== -1 && !isMemberInSheet(memberEmail, data)) {
      removeMemberFromGroup(groupEmail, memberEmail);
      Logger.log("Member not in the sheet");
      
      //sheet.getRange(i + 1, statusColumnIndex).setValue('Member removed');
      }
    // Check if member is already in the group and the sheet
    else if (currentMembers.indexOf(memberEmail) !== -1 && isMemberInSheet(memberEmail, data)) {
      sheet.getRange(i + 1, statusColumnIndex).setValue('Member already in the group');
    }
  }
} 

function isMemberInSheet(memberEmail, data) {
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === memberEmail) {
      Logger.log("Member :" + data[i][2] + data[i][2] + "===" + memberEmail);
      return true;
    }
  }
  return false;
}

function getCurrentGroupMembers(groupEmail) {
  var url = `https://www.googleapis.com/admin/directory/v1/groups/${encodeURIComponent(groupEmail)}/members`;
  var options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  var members = JSON.parse(response.getContentText()).members || [];
  return members.map(function(member) {
    return member.email;
  });
}

function addMemberToGroup(groupEmail, memberEmail) {
  var url = `https://www.googleapis.com/admin/directory/v1/groups/${encodeURIComponent(groupEmail)}/members`;
  var member = {
    email: memberEmail,
    role: 'MEMBER' // Use 'OWNER' or 'MANAGER' or 'MEMBER' based on the member's role
  };

  var options = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(member)
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
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
