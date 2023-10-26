function createGoogleGroup() {
    var group = {
      email: 'test-group@rawalwasia.in', // Replace with desired group email address
      name: 'Test Group', // Replace with desired group name
      description: 'Description of the new group' // Replace with desired group description
    };
    
    var url = 'https://www.googleapis.com/admin/directory/v1/groups';
    var options = {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(group)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    Logger.log(response.getContentText());


}
