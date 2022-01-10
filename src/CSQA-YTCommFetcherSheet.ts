//=====================================================================================================================
//
// Name: TweetFetcher
//
// Desc:
//
// Author:  Mune
//
// History:
//  2022-01-09 : Initial version
//
//=====================================================================================================================

// ID of Target Google Spreadsheet (Book)
//let VAL_ID_TARGET_BOOK           = '';
// ID of Google Drive to place backup Spreadsheet
//let VAL_ID_GDRIVE_FOLDER_BACKUP  = '';

//=====================================================================================================================
// DEFINES
//=====================================================================================================================
let VERSION                      = 1.0;
let TIME_LOCALE                  = "JST";
let FORMAT_DATETIME_ISO8601_DATE = "yyyy-MM-dd";
let FORMAT_DATETIME_ISO8601_TIME = "HH:mm:ss";
let FORMAT_DATETIME_DATE_NUM     = "yyyyMMdd";
let FORMAT_DATETIME              = "yyyy-MM-dd (HH:mm:ss)";
let FORMAT_TIMESTAMP             = "yyyyMMddHHmmss";
let NAME_SHEET_USAGE             = "!USAGE";
let NAME_SHEET_LOG               = "!LOG";
let NAME_SHEET_ERROR             = "!ERROR";
let SHEET_NAME_COMMON_SETTINGS   = "%Video IDs%";

//=====================================================================================================================
// DEFINES
//=====================================================================================================================
class DataRow {
  top_seq_num                    : any;
  cld_seq_num                    : any;
  published_at                   : any;
  comment                        : any;
  author                         : any;
  good_count                     : any;
  reply_count                    : any;
  constructor(
    top_seq_num                  : any,
    cld_seq_num                  : any,
    published_at                 : any,
    comment                      : any,
    author                       : any,
    good_count                   : any,
    reply_count                  : any
  ) {
    this.top_seq_num             = top_seq_num  ;
    this.cld_seq_num             = cld_seq_num  ;
    this.published_at            = published_at ;
    this.comment                 = comment      ;
    this.author                  = author       ;
    this.good_count              = good_count   ;
    this.reply_count             = reply_count  ;
  }
}
let HEADER_TITLES:DataRow = {
  top_seq_num                    : "seq #",
  cld_seq_num                    : "child #",
  published_at                   : "date time",
  comment                        : "comment",
  author                         : "author",
  good_count                     : "good count",
  reply_count                    : "reply count"
}
class HeaderInfo {
  videoId                        : string;
  videoTitle                     : string;
  rowHeader                      : number;
  headerCols                     : DataRow;
}

let MAX_ROW_SEEK_HEADER          = 20;
let MAX_COL_SEEK_HEADER          = Object.keys(HEADER_TITLES).length;
let DEFAULT_ROW_HEADER           = 4;

let OFFSET_ROW_VIDEO_LIST        = 1;

//=====================================================================================================================
// GLOBALS
//=====================================================================================================================
let g_isDebugMode                = true;
let g_isEnabledLogging           = true;
let g_isDownlodingMedia          = true;
let g_datetime                   = new Date();
let g_timestamp                  = TIME_LOCALE + ": " + Utilities.formatDate(g_datetime, TIME_LOCALE, FORMAT_DATETIME);
//let g_folderBackup               = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_BACKUP);
let g_book                       = SpreadsheetApp.getActiveSpreadsheet();

//=====================================================================================================================
// CODE for General
//=====================================================================================================================

//
// Name: gsAddLineAtLast
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function gsAddLineAtBottom(sheetName, text) {
  try {
    let sheet = g_book.getSheetByName(sheetName);
    if (!sheet) {
      sheet = g_book.insertSheet(sheetName, g_book.getNumSheets());
    }
    let range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2);
    if (range) {
      let valsRng = range.getValues();
      let row = valsRng[0];
      row[0] = g_timestamp;
      row[1] = String(text);
      range.setValues(valsRng);
    }
  }
  catch (e) {
    Logger.log("EXCEPTION: gsAddLineAtBottom: " + e.message);
  }
}

//
// Name: logOut
// Desc:
//
function logOut(text) {
  text = g_timestamp + "\t" + text;
  if (!g_isEnabledLogging) {
    return;
  }
  gsAddLineAtBottom(NAME_SHEET_LOG, text);
}
//
// Name: errOut
// Desc:
//
function errOut(text) {
  text = g_timestamp + "\t" + text;
  gsAddLineAtBottom(NAME_SHEET_ERROR, text);
}

//=====================================================================================================================
// CODE
//=====================================================================================================================

//
// Name: generateHeader
// Desc:
//
function generateHeader(sheet, headerTitles:DataRow, videoId:string):HeaderInfo {
  if (sheet.getMaxRows() > 1) {
    sheet.deleteRows(2, sheet.getMaxRows() - 1);
  }
  if (sheet.getMaxColumns() > Object.values(headerTitles).length) {
    sheet.deleteColumns(Object.values(headerTitles).length + 1, sheet.getMaxColumns() - Object.values(headerTitles).length );
  }
  let range = sheet.getRange(1, 1, (DEFAULT_ROW_HEADER + 1), MAX_COL_SEEK_HEADER);
  if (range == null) {
    throw new Error("generateHeader: range wasn't able to acquired.");
  }
  let valsRng = range.getValues();
  valsRng[0][0] = videoId;
  let headerCols = new DataRow(null, null, null, null, null, null, null);
  let objRow = valsRng[DEFAULT_ROW_HEADER];
  for (let c = 0; c < Object.values(headerTitles).length; c++) {
    objRow[c] = Object.values(headerTitles)[c];
    headerCols[Object.keys(headerTitles)[c]] = c;
  }
  range.setValues(valsRng);
  return { videoId: videoId, videoTitle: null, rowHeader: DEFAULT_ROW_HEADER, headerCols: headerCols };
}

//
// Name: generateHeader
// Desc: get comments at the both top and replies recursively
//
function getComments( videoId:string, parentId:string, pageToken:string, seqTop:number, seqCld:number, listCommData:DataRow[] ):DataRow[] {
  if (typeof pageToken == 'undefined') {
    return listCommData;
  }

  let listComment = null;
  if ( videoId ) {
    listComment = YouTube.CommentThreads.list('id, replies, snippet', {
      videoId: videoId,
      maxResults: 100,
      order: 'time',
      textFormat: 'plaintext',
      pageToken: pageToken,
    } );
  } else {
    listComment = YouTube.Comments.list('id, snippet', {
      maxResults: 100,
      parentId : parentId,
      textFormat: 'plaintext',
      pageToken: pageToken,
    });
  }

  listComment.items.forEach( item => {
      let countReply = (videoId)? item.snippet.totalReplyCount : '' ;
      let snippet = (videoId)? item.snippet.topLevelComment.snippet : item.snippet ;
      listCommData.push(new DataRow(++seqTop, (videoId)? '': ++seqCld, snippet.publishedAt, snippet.textDisplay, snippet.authorDisplayName, snippet.likeCount, countReply));

      if ( videoId && countReply > 0 ) {
        // all replies are covered by this page token
        if(countReply == item.replies.comments.length) {
          for (let i = 0; i < item.replies.comments.length; i++) {
            let snippetCld = item.replies.comments[i].snippet;
            listCommData.push(new DataRow(seqTop, i+1, snippetCld.publishedAt, snippetCld.textDisplay, snippetCld.authorDisplayName, snippetCld.likeCount, ''));
          }
        }
        // replies are covered by multiple pages
        else {
          let parentId = item.snippet.topLevelComment.id;
          listCommData = getComments(null, parentId, '', seqTop, 0, listCommData );
        }
      }
  });
  return getComments( videoId, parentId, listComment.nextPageToken, ++seqTop, 0, listCommData );
}

//
// Name: getData
// Desc:
//
function getData(headerInfo:HeaderInfo ):DataRow[] {
  let video = YouTube.Videos.list('id,snippet, statistics', { id: headerInfo.videoId, });
  headerInfo.videoTitle = video.items[0].snippet.title;
  return getComments( headerInfo.videoId, null, '', 0, 0, [] );
}

//
// Name: updateSheet
// Desc:
//
function updateSheet(sheet, videoId:string):string {
  let headerInfo = generateHeader(sheet, HEADER_TITLES, videoId);
  sheet.setName(headerInfo.videoId);
  let listCommData = getData(headerInfo);
  if ( !headerInfo.videoTitle ) {
    return null;
  }
  sheet.getRange(2,1,1,1).setValue('=HYPERLINK("https://www.youtube.com/watch?v=' + videoId + '", "' + headerInfo.videoTitle + '")' );
  if ( listCommData.length ){
    let range = sheet.getRange(headerInfo.rowHeader+2, 1, listCommData.length, sheet.getLastColumn());
    let valsRng = range.getValues();
    for (let i=0; i<listCommData.length; i++ ) {
      valsRng[i][headerInfo.headerCols.top_seq_num]  = listCommData[i].top_seq_num;
      valsRng[i][headerInfo.headerCols.cld_seq_num]  = listCommData[i].cld_seq_num;
      valsRng[i][headerInfo.headerCols.published_at] = listCommData[i].published_at;
      valsRng[i][headerInfo.headerCols.comment]      = listCommData[i].comment;
      valsRng[i][headerInfo.headerCols.author]       = listCommData[i].author;
      valsRng[i][headerInfo.headerCols.good_count]   = listCommData[i].good_count;
      valsRng[i][headerInfo.headerCols.reply_count]  = listCommData[i].reply_count;
    }
    range.setValues(valsRng);
  }
  return headerInfo.videoTitle;
}

//
// Name: getVideoIds
// Desc:
//
function getVideoIds( sheet ):string[] {
  let range = sheet.getRange(1 + OFFSET_ROW_VIDEO_LIST, 1, sheet.getLastRow()-OFFSET_ROW_VIDEO_LIST, 1); // video Ids need to be placed at the first column
  let valsRng = range.getValues();
  let listVideoIds:string[] = [];
  for ( let i=0; i<valsRng.length; i++) {
    if (valsRng[i][0]) {
      if ( (valsRng[i][0]).trim().toLowerCase().match(/youtube/) ) {
        listVideoIds.push( (valsRng[i][0].trim()).replace(/.+v=([^?&]+).*/, '$1'));
      } else {
        listVideoIds.push((valsRng[i][0]).trim());
      }
    } else {
      listVideoIds.push(null);
    }
  }
  return listVideoIds;
}

//
// Name: main
// Desc: entry point of this proguram
//
function main() {
  try {
    let sheetConfig = g_book.getSheetByName(SHEET_NAME_COMMON_SETTINGS);
    if (!sheetConfig) {
      throw new Error("The sheet \"" + SHEET_NAME_COMMON_SETTINGS + "\" was not found.");
    }
    let videoIds: string[] = getVideoIds(sheetConfig);
    for (let i = 0; i < videoIds.length; i++) {
      if (!videoIds[i]) {
        continue;
      }
      let sheet = g_book.getSheetByName(videoIds[i]);
      if (!sheet) {
        sheet = g_book.insertSheet(videoIds[i], g_book.getNumSheets());
      }
      let videoTitle:string = updateSheet(sheet, videoIds[i]);
      if (videoTitle) {
        sheetConfig.getRange(i + 1 + OFFSET_ROW_VIDEO_LIST, 2, 1, 3).setValues(
          [['=HYPERLINK("https://www.youtube.com/watch?v='+ videoIds[i] + '", "' + videoTitle +'")'
          , "OK"
          , g_timestamp]]);
      } else {
        sheetConfig.getRange(i + 1 + OFFSET_ROW_VIDEO_LIST, 2, 1, 3).setValues([['' , "OK", g_timestamp]]);
      }
    }
  }
  catch (ex) {
    errOut(ex.message);
  }
}