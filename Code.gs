function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('ExtractUrlFmDocAndWeb', 'extractUrlFmDoc')
      .addItem('ExtractUrlFmWebBody', 'extractUrlFmWebOnly')
      .addToUi();
}

function extractUrlFmDoc(){
  var sheetSourceName = "Docs";
  var sheetResName = "Extracted URLs";
  var sheetSettings = 'Settings';
  clean(sheetResName);
  
  var urlScaner = new UrlScaner(sheetSourceName, sheetResName, sheetSettings);
  var sourceUrls = urlScaner.gerUrlsSource();
  if(sourceUrls[0] != ""){
    var onlyWeb = false;
    var res = urlScaner.scanEachUrl(sourceUrls, onlyWeb);
//    Browser.msgBox("All Done");
  }else{
    Browser.msgBox("At least one url requered in 'Docs!A2:A'");
  }
}

function extractUrlFmWebOnly(){
  var sheetSourceName = "Docs";
  var sheetResName = "All Links Extracted from Web body";
  var sheetSettings = 'Settings';
  clean(sheetResName);
  
  var urlScaner = new UrlScaner(sheetSourceName, sheetResName, sheetSettings);
  var sourceUrls = urlScaner.gerUrlsSource();
  if(sourceUrls[0] != ""){
    var onlyWeb = true;
    var res = urlScaner.scanEachUrl(sourceUrls, onlyWeb);
//    Browser.msgBox("All Done");
  }else{
    Browser.msgBox("At least one url requered in 'Docs!A2:A'");
  }
}


// === Clean Res page
function clean(sheetName){
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
  var lastRow = sheet.getLastRow();
  if( lastRow > 1 ){
    sheet.getRange(2, 1, lastRow, 5 ).clearContent();
  }
}

/*
* Class Urls Scaner
* 
* Check each Url and scan depends it source 
*/
var UrlScaner = function(sheetSourceName, sheetResName, sheetSettings){
  this.sheetSourceName = sheetSourceName;
  this.sheetResName = sheetResName;
  this.sheetSettings = sheetSettings;
 
  // === 
  this.scanEachUrl = function(sourceUrls, onlyWeb){
    if(!sourceUrls) { return Logger.log("ERROR in scanEachUrl()") };
    sourceUrls.forEach( function(url, index) {
      var docUrl = ( url.indexOf("//docs.google") > 0 );          // return true if url is GDoc link
      if(docUrl && !onlyWeb){
        // Pull data from G Doc 
        var doc = DocumentApp.openByUrl(url);
        var docParser =new DocParser(doc);
        var linkList = docParser.getLinks();
        var uniqList = docParser.getUnicLinks(linkList);
        docParser.printData(uniqList, this.sheetResName);
//        Logger.log(uniqList);
      }else{
        // Pull data from Web
        var webParser = new WebParser(url);
        if(onlyWeb){                                            // 
          var linkList = webParser.getAllLinks();
        }else{
          var tagList = this.getTags();                         // get tags from settings tab
          Logger.log(tagList);
          var linkList = webParser.getLinksByTag(tagList);
        }
        webParser.printData(linkList, this.sheetResName);
      }      
    }, this)
  }
  
   // ===
  this.gerUrlsSource = function(){
    var sheet = SpreadsheetApp.getActive().getSheetByName(this.sheetSourceName);
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - 1;
    var urlSourceList = sheet.getRange(2, 1, numRows).getValues().map(function(item){ return item[0]});
  //  Logger.log(docUrlList);
    var filtered = [], i=0;
    while( i< urlSourceList.length && urlSourceList[i] != "" ){
      filtered.push(urlSourceList[i]);
      i++;
    }
    return filtered;
  }
  
  // ===
  this.getTags = function(){
    var sheet = SpreadsheetApp.getActive().getSheetByName(this.sheetSettings);
    try{
      var lastRow = sheet.getLastRow();
      var numRows = lastRow - 1;
      var urlSourceList = sheet.getRange(2, 1, numRows).getValues().map(function(item){ return item[0]});
      var filtered = [], i=0;
      while( i< urlSourceList.length && urlSourceList[i] != "" ){
        filtered.push(urlSourceList[i]);
        i++;
      }
      return filtered;
    }catch(e){
      Logger.log("Error in getTags(): "+e);
      return null;
    }
  }
}

/*
* Class WebParser
*
* This class using Cheerio (cheeriogs) for Google Apps Script 
* Script ID: 1ReeQ6WO8kKNxoaA_O0XEQ589cIrRvEBA9qcWpNqdOP17i47u6N9M5Xh0
* Web: https://github.com/fgborges/cheeriogs
*/
var WebParser = function(url){
  this.url = url;
  
  this.getAllLinks = function(){
    var html = UrlFetchApp.fetch(this.url).getContentText();
    const $ = Cheerio.load(html);
    
    var self = this;
    var data = [];
    var unicLinks = [];
    var siteUrl = this.url;
    var tag = 'body';
    
    $('body a').each(function(i, el){
      var linkText = $(el).text().trim();
      var link = $(el).attr('href');
      var unic = ( unicLinks.indexOf(link) == -1);
      if(unic && linkText != ''){
          var asin = self.getAsin(link);
          data.push([linkText, asin, link, siteUrl, tag]);
      }
    });
//    Logger.log(data.length);
//    Logger.log(data);
    return data;
  }
  
  this.getLinksByTag = function( opt_tags  ){
    if( opt_tags && opt_tags[0] != ''){
        var tags = opt_tags;
    }else{
        var tags = ['h1', 'h2', 'h3', 'h4', 'h5', 'p'];
    }
    var html = UrlFetchApp.fetch(this.url).getContentText();
    const $ = Cheerio.load(html);
    
    var self = this;
    var data = [];
    var unicLinks = [];
    var siteUrl = this.url;
    
    // iterate each tag
    tags.forEach(function(tag, index){
//            console.log(item, index);
        $(tag+' a').each(function(i, elem){
            var linkText = $(elem)
                        .text()
                        .trim();
            var link = $(elem).attr('href');
            var unic = ( unicLinks.indexOf(link) == -1);
            if(unic && linkText != ''){
                var asin = self.getAsin(link);
                data.push([linkText, asin, link, siteUrl, tag]);
            }
        })
    }, this)
//    Logger.log(data.length);
//    Logger.log(data);
    return data;
  }

  
  // === Helper RegEx get ASIN code from link
  this.getAsin = function(link){
//    var url = "https://www.amazon.com/Roland-DB-90-BOSS-Metronome/B000ATOFS4/ref=sr_1_3?keywords=BOSS+DB-90&qid=1568907205&s=musical-instruments&sr=1-3";
    var amazon = ( link.indexOf("amazon.com") > 0 );
    if(amazon){
      var keyWord = new RegExp(".+\\/dp\\/(.+)\\/", 'i');
      var match = link.match(keyWord);
  //    Logger.log(match);
      return match ?  match[1]: "" ;    // return "" if no any Matches
    }else{ return "";}
  }
  
  this.printData = function(unicList, resSheet){
    if(!unicList){return false };
//    unicList = [unicList];
    var sheet = SpreadsheetApp.getActive().getSheetByName(resSheet);
    var lastRow = sheet.getLastRow();
//    sheet.getRange(1, 1, sheet.getLastColumn(), sheet.getLastRow()).clear();
    Logger.log(unicList.length);
    var range = sheet.getRange(lastRow+2, 1, unicList.length, unicList[0].length).setValues(unicList);
  }
} // === END Class


/*
* Class DocParser
*
* Parse Google Document
*/
var DocParser = function(doc){
  this.doc = doc;
  
  this.getUnicLinks = function(linkList){
    if(!linkList){var linkList = this.getLinks() };
    var unicUrl = []; 
    unicList = [];                                   // List with format to set data in cell range
//    Logger.log(linkList[0]);
    linkList.forEach( function( item, index ){
       item.forEach( function(el, ind){              // attr linkText, url
         var url = el.url;
//         Logger.log(el.url);
         var unic = ( unicUrl.indexOf(url) == -1 );
         var amazon = ( url.indexOf("amazon") > 0 );
         if(unic){ 
//           Logger.log(url);// if url unic - made unicList                                           
           var asin = this.getAsin(url);
           
           unicList.push([ el.linkText, asin, url, doc.getUrl() ]);
           unicUrl.push(url);           
         } 
       }, this);
    }, this)
//    Logger.log(unicList);
    return unicList
  }
    // === Helper RegEx get ASIN code from url
  this.getAsin = function(url){
//    var url = "https://www.amazon.com/Roland-DB-90-BOSS-Metronome/B000ATOFS4/ref=sr_1_3?keywords=BOSS+DB-90&qid=1568907205&s=musical-instruments&sr=1-3";
    var keyWord = new RegExp(".+\\/dp\\/(.+)\\/", 'i');
    var match = url.match(keyWord);
//    Logger.log(match);
    return match ?  match[1]: "" ;    // return "" if no any Matches
  }
  
  this.printData = function(unicList, resSheet){
    if(!unicList){var unicList = this.getUnicLinks() };
    var sheet = SpreadsheetApp.getActive().getSheetByName(resSheet);
    var lastRow = sheet.getLastRow();
    Logger.log(unicList.length);
    var range = sheet.getRange(lastRow+2, 1, unicList.length, unicList[0].length).setValues(unicList);
  }
  
  // ===
  this.getLinks = function(){
    var links = [];
    var element = this.doc.getBody();
    
    var numChildren = element.getNumChildren();              // get number of children ()
    for(var i=0; i<numChildren; i++){
      var innerElement = element.getChild(i);                // typeOf innerElement - Paragraph
      if ( innerElement.getNumChildren() > 0 ){              // if paragraph not empty
        var link = this.getLinksFromText(innerElement.getChild(0));
        if(link != null){
          links.push(link);
        }
      }  
    }
    return links;
  }
  
  // === Helper
  this.getLinksFromText = function(element){
    var links = [];
    if (element.getType() === DocumentApp.ElementType.TEXT) {            // if elem is TEXT check entire text
      var textObj = element.editAsText();
      var text = element.getText();
      var inUrl = false;
      var length = text.length;
      for (var psn=0; psn < length; psn++) {                             // getLinkUrl from each position 
        
        var url = textObj.getLinkUrl(psn);
  //      Logger.log(url);
        if (url != null) {                                               // if found link 
          if (!inUrl) {                                                      // if this is first position - save startOffset, url 
            inUrl = true;
            var curUrl = {};
            curUrl.element = element;
            curUrl.url = String( url ); // grab a copy
            curUrl.startOffset = psn;
//            Logger.log("First Pos with url"+psn);
          }
          else if ( psn+1 == text.length-1  ){                               // if NEXT is last pos - save endOffset, linkText, push data to list
            curUrl.endOffsetInclusive = psn;
            curUrl.linkText = text.substring(curUrl.startOffset, curUrl.endOffsetInclusive+2);
  //          Logger.log(curUrl.linkText);
            links.push(curUrl);  // add to links
            curUrl = {};
          }
          else if (  psn+1 < text.length-1){                                 // if NEXT is NOT last pos
            if(!textObj.getLinkUrl(psn+1)){                                      // and if NEXT is NOT a URL - save endOffset, linkText, push data to list
              curUrl.endOffsetInclusive = psn;
              curUrl.linkText = text.substring(curUrl.startOffset, curUrl.endOffsetInclusive+2);
    //          Logger.log(curUrl.linkText);
              links.push(curUrl);  // add to links
              inUrl = false;        // reset flag
              curUrl = {};
              
            }
          }
        }
      }
    }
  //  Logger.log(links);
    return links.length>0 ? links: null;
  }
} // === END Class


