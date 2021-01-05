//Paperpile tools for Tufts manuscripts v1.002 10/27/20
//Copyright Gregory Hongyuan Fan 2020
//102720 Added move by URL function
//Updated comma function to not delete URLs and to only insert spaces after commas that are in between parentheses
//To install: 
//1. Open the "Tools" menu button in a googledocs manuscript.
//2. Open the script editor.
//3. Delete everything on the screen.
//4. Paste the entirety of this text in the blank area.
//5. Save the project. A suggested name is "NDLR Paperpile Tools"
//6. Refresh the googledocs manuscript.
//7. Click on the "Add-ons" menu to access the tools in this script.

function onInstall() {
  onOpen();
}
function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
  .addItem('Combine In-Text Citations', 'citation_combine')
  .addItem('Abbreviate Journal Names', 'abbreviate_journals')
  .addItem('Remove doi from references', 'doi_remove')
  .addItem('Add spaces between commas', 'comma_addspace')
  .addItem('Move File by URL', 'move_byurl')
  .addToUi()
}

//function onOpen() {
//  var ui = DocumentApp.getUi();
//  ui.createMenu('Scripts')
//      .addItem('Combine Citations', 'citation_combine')
//      .addToUi() 
//}

function citation_combine() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var textObj = body.editAsText();
  var text = body.getText();
  var ui = DocumentApp.getUi();
  
  //Set first result to nothing
  var search_result = null;
  var original_url = null;
  var readout = 'Intra-parenthetical citation compilation complete. URLs within parentheses have been combined into single URLs for proper Paperpile formatting. Use Paperpile to reformat the in-text citations into the proper format for publication submission.\n';
  
  do {
    search_result = body.findText("\\(.+?\\)", search_result);  //Get initial search result.
    if (search_result != null) {
      var search_result_element = search_result.getElement();  //Get element of search result.
      var startoffset = search_result.getStartOffset();  //Get starting offset
      var endoffset = search_result.getEndOffsetInclusive();  //Get ending offset
      var findtext = search_result_element.getText().substring(startoffset, endoffset);  //Establish final string to be replaced, if necessary. Not currently used.
      //var combined_url = "";  //reset combined url. Deprecated with the find unique function
      readout = readout + '\n' + search_result_element.getText() + '\n';  //Diagnostic code
      var url_array = [];  //reset url array
      for (var i = startoffset; i < endoffset; i++) {
        var url = search_result_element.asText().getLinkUrl(i);  //get url by character in body region between offsets
        if (url != null) {  //only activate if the url is not empty as there are nonURL characters captured
          if (url_array[0] == null) {  //url array is set to null 
            original_url = url;  //set the original url to the new url
            //combined_url = url;  //grab the whole url for the first element of the url array. Deprecated with the find unique function.
            url_array.push(url);  //url_array[0] should always be the full html link.
          } else if (original_url != url) {  //skip if the new url is the same as the original url.
            var new_url = url.match(/[^\/]+?$/g);  //grab only the end of the link
            url_array.push(new_url[0]);  //add the end of new link to the url array.
          }
          readout = readout + '\nURL "' + url + '" from "' + search_result_element.getText().substring(i - 2, i + 1) + '"'; //Diagnostic Code
        }
      }
      var combined_url = url_array[0];  //set the root of the combined url to the big element 0 of the array
      //Find unique function
      for (var j = 1; j < url_array.length; j++) {
        var num_appear = 0;  //reset number of appearances
        for (var k = 0; k <= j; k++) {
          if (url_array[j] == url_array[k]) {  //Activate only if the new url is found to repeat. Should expect 1 occurance per variable.
            num_appear = num_appear + 1;  //Count up how many times a variable is found in the array up to the current array number. If the array is repeated then increase the appearance number by 1.
          }
        }
        if (num_appear == 1) {
          combined_url = combined_url.concat("+", url_array[j]);  //Add array member to final url only if the number of occurances is equal to 1.
        }
      }
      search_result_element.asText().setLinkUrl(startoffset + 1, endoffset - 1, combined_url);  //Payload: change all characters within the search result to the combined url.
      readout = readout + '\nhave been combined into: "' + combined_url + '"\n';  //Diagnostic Code
      original_url = null;  //Reset original url.
    }
  }
  while (search_result != null);
  ui.alert(readout); //Produce readout
  
//  Code for diagnostics
//  var replacetext = search_result_element.asText().setLinkUrl(startoffset, endoffset, combined_url);
//          body.insertParagraph(1,combined_url);
//    
//  body.insertParagraph(1,search_result_element.asText().getLinkUrl(295));
//  if (search_result != null) {
//    body.replaceText("asdfg",search_result.getElement().getText());
//  } else {
//    body.replaceText("asdfg","null");
//  }
}

function abbreviate_journals() {
  var doc = DocumentApp.getActiveDocument();
  var sheet = SpreadsheetApp.openById("1VYXovvPOLPcGGgaRkwDw00J6Vc89lxvQjJLIy8jH_p8"); //Get spreadsheet with journal names and abbreviations
  var title_array = sheet.getRangeByName("title_full").getValues(); //Populate array with full names from spreadsheet
  var abbrev_array = sheet.getRangeByName("title_abbreviation").getValues(); //Populate array with abbreviations from spreadsheet
  var body = doc.getBody();
  var ui = DocumentApp.getUi();
  
  //Reset search result and alert dialog text
  var search_result = null;
  var readout = 'Journal abbreviation is complete. Single word journal titles were unchanged.\n';
  
  do {
    //search_result = body.findText("(^\\d+\\.\\s\\t.*?\\..*?\\.\\s*)(.*?)(\\s*\\.)",search_result); //Initial find; returns full reference and not just title
    search_result = body.findText("(^\\d+\\.\\s\\t)",search_result); //Initial find; returns full reference and not just title
    //ui.alert(search_result.getElement().getText());
    if (search_result == null){
    } else {
      readout = readout + '\n' + search_result.getElement().getText();
    }
    if (search_result != null) {
      var search_result_element = search_result.getElement();  //Turn search result into an element
      var search_result_start = search_result_element.findText("(^\\d+\\.\\s\\t.*?\\..*?\\.\\s*)");  //Get the starting point for the italic search area from the end offset of this regex
      var search_result_end = search_result_element.findText("(^\\d+\\.\\s\\t.*?\\..*?\\.\\s*)(.*?)(\\s*\\..*)");  //Get the ending point for the italic search area from the end offset of this regex
      var startoffset = search_result_start.getEndOffsetInclusive();  //Get starting offset of italic search area
      var endoffset = search_result_end.getEndOffsetInclusive();  //Get ending offset of italic search area
      //readout = readout + " " + startoffset + " " + endoffset;  //Diagnostic code
      var italic_array = [];  //Initialize array of where italic characters are going to be found
      var digit = search_result_element.getText().match(/^\d+/);  //diagnostic entry numbering
      for (var i = startoffset; i < endoffset; i++) {
        //readout = readout + '\nnew i here ' + i + ' match here ' + search_result_element.getText().substring(i, i + 1).match(/\./);  //Diagnostic code
        if (search_result_element.isItalic(i) == true)
        {
          italic_array.push(i); //Only populate the array with character offsets that are italic positive
          //readout = readout + ' italic array length ' + italic_array.length;
        }
        if (italic_array.length > 1 && search_result_element.getText().substring(i, i + 1).match(/\./) != null) {
          //readout = readout + ' exit here ' + digit + ' ' + search_result_element.getText().substring(i, i + 1).match(/\./);  //Diagnostic code
          break;
        }
      }
      if (italic_array[0] == null) {
        readout = readout + '\nFeedback: No italics found.\n';
      }
      if (italic_array[0] != null) {  //Only proceed if the italic array is not empty
        var jname_start = Math.min.apply(Math, italic_array);  //Get the smallest number in the italic array
        var jname_end = Math.max.apply(Math, italic_array);  //Get the largest number in the italic array
        var jname_original = search_result_element.getText().substring(jname_start, jname_end + 1).trim();  //Get the text from the smallest and largest offsets in the italic array
        
        //Reset variables
        var jname_replace = "";
        var jname_is_abbrev = false;
        var jname_is_single = false;
        
        if (/\s/g.test(jname_original) == false) {  //Select for journal names that do not have spaces in them; replace single word titles in order to trim any extra spaces from each title
          jname_is_single = true;
          jname_replace = jname_original.toString();
          readout = readout + '\nFeedback: "' + jname_original + '"' + ' is a single word title and was not replaced.\n';  //Update readout 
        } else { // proceed if (jname_is_single == false)
          for (var k = 0; k < abbrev_array.length; k++) {
            if (abbrev_array[k].toString().toLowerCase() == jname_original.toLowerCase()) {  //Select for titles that already match an entry in the abbreviation array; replace text in order to trim extra spaces from the title
              jname_is_abbrev = true;
              jname_replace = abbrev_array[k].toString();
              readout = readout + '\nFeedback: "' + abbrev_array[k].toString() + '"' + ' does not need replacing.\n';
              break;
            }
          }
          if (jname_is_abbrev == false && jname_is_single == false) {  //Select for titles that do not match the abbreviation array and that also contain spaces as these correct titles may be matches for longer titles and the title will end up incorrect.
            for (var j = 0; j < title_array.length; j++ ) {
              var test_title = title_array[j].toString().toLowerCase().replace(/[^\w]/g,'');  //Scrunch the full length title array entry down to lowercase and no punctuation or spaces
              var test_jname = jname_original.toLowerCase().replace(/&/,'and');  //Scrunch the journal name down to lowercase and no punctuation or spaces
              var test_jname = test_jname.replace(/[^\w]/g,'');
              if (test_title.indexOf(test_jname) > -1) {  //Find the manuscript title within the title array. This allows for incomplete manuscript titles to match.
                jname_replace = abbrev_array[j].toString();
                readout = readout + '\nFeedback: "' + jname_original.toLowerCase() + '"' + ' was replaced with ' + '"' + jname_replace + '".' + '\n';  //Log replacement in readout variable
                break;
              }
            }
          }
        }
        if (jname_replace != "") {
          search_result_element.deleteText(jname_start, jname_end);  //If the replacement title is not empty then delete the old text
          search_result_element.insertText(jname_start, jname_replace);  //Insert new text at same location as deleted text
          search_result_element.setItalic(jname_start, jname_start + jname_replace.length - 1, true);  //Italicize the new text
        } else if (jname_replace == "" && jname_original != "" && jname_is_abbrev == false && jname_is_single == false) {
          readout = readout + '\nFeedback: "' + jname_original + '"' + ' was not found.\n';  //Log titles that were not found in the database
        }
      }
    }
    //var findtext = search_result_element.getText().substring(startoffset, endoffset); //Establish final string to be replaced, if necessary. Not currently used.
  }
  while (search_result != null);
  
  ui.alert(readout); //Output the readout into the alert system
  //body.insertParagraph(1,jname_replace); //Diagnostic code. Depracated by alert system
}

function doi_remove() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var search_result = null;
  var ui = DocumentApp.getUi();
  var readout = "doi removal complete.\n";
  do {
    search_result = body.findText("doi:\\d.+", search_result);
    if (search_result != null) {
      var search_result_element = search_result.getElement();
      var startoffset = search_result.getStartOffset();
      var endoffset = search_result.getEndOffsetInclusive();
      var text_original = search_result_element.getText().substring(startoffset, endoffset + 1);
      var text_replace = "";
      readout = readout + '\n' + search_result_element.getText();
      readout = readout + '\nFeedback: "' + text_original + '"' + ' was replaced with ' + '"' + text_replace + '".\n';
      search_result_element.deleteText(startoffset, endoffset);
      search_result_element.insertText(startoffset, text_replace);
    }
  }
  while (search_result != null)
  ui.alert(readout);
}

function comma_addspace() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var search_result = null;  //Starting search result is null, as the search function iterates over search results.
  var ui = DocumentApp.getUi();
  var readout = "A single space was inserted between commas that were immediately followed by non-space characters.\n";
  do {
    search_result = body.findText("\\([^\\(]*?,\\w[^\\(]*?\\)", search_result); //Find only commas followed by a letter or digit and also between parentheses.
    if (search_result != null) {
      var search_result_element = search_result.getElement();
      var startoffset = search_result.getStartOffset(); //Starting offset within the element.
      var endoffset = search_result.getEndOffsetInclusive(); //Ending offset within the element.
      var text_original = search_result_element.getText().substring(startoffset, endoffset + 1); //String version of the element text
      var text_replace = text_original.replace(/(?<=,)(\w)/,' $1');  //Regex to demonstrate the replacement for the readout.
      var n_startoffset = text_original.search(/(?<=,)(\w)/) + startoffset;  //Regex to find the index of the replacement and to add it to the element startoffset
      readout = readout + '\n' + search_result_element.getText(); //Readout the element
      readout = readout + '\nFeedback: "' + text_original + '"' + ' was replaced with ' + '"' + text_replace + '".\n'; //Readout the replacement
      //search_result_element.deleteText(startoffset, endoffset);
      search_result_element.insertText(n_startoffset, " ");  //Insert the space
    }
  }
  while (search_result != null)
  ui.alert(readout);  //Print the final readout
}

function move_byurl(){
  var doc = DocumentApp.getActiveDocument();
  var doc_id = doc.getId();
  var ui = DocumentApp.getUi();
  var targeturl = ui.prompt('Destination Directory');
  if (targeturl.getSelectedButton() == ui.Button.OK) {
    var regex = new RegExp('[^/]+?$','g');
    var targetfolder_id = regex.exec(targeturl.getResponseText());
    var destFolder = DriveApp.getFolderById(targetfolder_id);
    DriveApp.getFileById(doc_id).moveTo(destFolder);
  }
}
