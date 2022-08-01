function bingSearch() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getCurrentCell().getRow()

  const subscriptionKey = 'Add your key here'
  let searchString = 'test'
  var query = encodeURIComponent(searchString)
  hostname = 'https://api.bing.microsoft.com/v7.0/search?q=' + query
  //mkt = 'en-US'
  //params = {'q' : query, 'mkt' : mkt }
  var options = {
    headers: {'Ocp-Apim-Subscription-Key': subscriptionKey}
  }
  let response = UrlFetchApp.fetch(hostname, options);
  
  //debug
  //let response = UrlFetchApp.fetch('')
  //
  let json = response.getContentText();
  data = JSON.parse(json);

  //Example: get the snippets of the the search results and the relevant images
  
  const searchOutput = []
  for (const webResult of data.webPages.value){
    searchOutput.push(webResult.snippet);
  }
  const imageOutput = []
  for (const imgResult of data.images.value){
    var thmbnail = 'image("' + imgResult.contentUrl + '",4,30,30)';
    var hyperlink = '=hyperlink("' + imgResult.contentUrl + '", ' + thmbnail + ')';
    imageOutput.push(hyperlink);
  }
  // Print the information in the spreadsheet (2 is the first column where we write results: row, 2 and searchOutput.length + 2)
  
  sheet.getRange(row,2,1,(searchOutput.length)).setValues([searchOutput]);
  sheet.getRange(row,(searchOutput.length + 2),1,(imageOutput.length)).setValues([imageOutput]);
}
