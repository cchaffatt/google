/***************************************************************************************************************************\
 * Authors: Carl and Josh 
 * Programmer: Carl
 * Script: canvas_calls.gs
 * Original Source: Nick Nelson - Custom Functions and Macros for google spreadsheet.
 * Other Sources: canvas api:                 https://canvas.instructure.com/doc/api/assignments.html
 *                google script api:          https://developers.google.com/apps-script/reference/spreadsheet/
 *                InstructureCon:             http://www.youtube.com/watch?v=rN4rDdNByWE
 *                Regexr:                     http://www.regexr.com/
 *                HTTPResponse:               http://www.tinkeredge.com/blog/2012/04/check-on-page-for-broken-links-with-google-docs/     (Cheok Lu)
 *
\**************************************************************************************************************************/
  

/* Global variables *
 ********************/
 
// SPREADSHEET
var ss = SpreadsheetApp.getActiveSpreadsheet();


// SHEETS
var readme_sheet = ss.getSheetByName('README');
var credentials_sheet = ss.getSheetByName('Credentials');
var selection_sheet = ss.getSheetByName('SELECTION');
var course_sheet = ss.getSheetByName('Course');
var assignments_sheet = ss.getSheetByName('Assignments');
var pages_sheet = ss.getSheetByName('Pages');
var discussions_sheet = ss.getSheetByName('Discussions');
var announcements_sheet = ss.getSheetByName('Announcements');
var modules_sheet = ss.getSheetByName('Modules');
var folders_sheet = ss.getSheetByName('Folders');
var files_sheet = ss.getSheetByName('Files');
var dates_sheet = ss.getSheetByName('Dates');
var lists_sheet = ss.getSheetByName('Lists');



// TOKEN: The user's distinct token generated from canvas's website  
var token = credentials_sheet.getRange('B2').getValue();                   /*****************************************************************************************************************************\
                                                                            *
                                                                            * ATTENTION: DO NOT SHARE YOUR TOKEN PUBLICLY, IT WILL GIVE ACCESS TO ALL THE INFORMATION IN YOUR ACCOUNT. SECURITY ISSUE-- 
                                                                            *
                                                                           \*****************************************************************************************************************************/        
// ITEMS: Items requested per page
var ITEMS_PER_PAGE = '?per_page=100';  


// SPREADSHEET STRIPES
var ZEBRA = [ '#efefef', '#FAEBFF', '#E6F5F5'];                            // Randomly choose between: light gray, light purple, light green
var STRIPE_COLOR = ZEBRA[ Math.floor( Math.random() * ZEBRA.length ) ];


/* HTTP options and credentials *
 ********************************/
var URL = 'https://champlain.instructure.com/api/v1/courses/';
var courseID = selection_sheet.getRange('C5').getValue();

if( typeof courseID == 'number' )
  URL = 'https://champlain.instructure.com/api/v1/courses/' + courseID;

var headers = {
    'Authorization': 'Bearer ' + token
};
var options = {
    'headers': headers
};






/**************************************************************************************************
 * Function: onOpen
 * Purpose: Trigger script to fill google spreadsheet to enable user api calls.
 **************************************************************************************************/

function onOpen() 
{
   // Create buttons 
   var repopulateEntries = [];                    
   var wipeEntry = [];
   var check = [];
  
   // refreshes data entries: Renew, Reload, Refresh, Update, Restore, Repopulate
   repopulateEntries.push({ name: "All Sheets", functionName: "updateAll" },
                          { name: "SELECTION Sheet", functionName: "updateSelection" },
                          { name: "Course Sheet", functionName: "updateCourse" },                      
                          { name: "Assignments Sheet", functionName: "updateAssignments" },
                          { name: "Pages Sheet", functionName: "updatePages" },
                          { name: "Discussions Sheet", functionName: "updateDiscussions" },
                          { name: "Announcements Sheet", functionName: "updateAnnouncements" },
                          { name: "Modules Sheet", functionName: "updateModules" },  /*
                          { name: "Folders Sheet", functionName: "updateFolders" },
                          { name: "Files Sheet", functionName: "updateFiles" },       */
                          { name: "Folders and Files Sheet", functionName: "updateFoldersAndFiles" });
                          

   // Clear data entries
   wipeEntry.push({ name: "Clear Sheet", functionName: "wipeSheet" },
                  { name: "Start from Scratch", functionName: "wipeAll" });
   
   // Fires up trigger that checks for broken links and date violations
   check.push({ name: "Broken Links", functionName: "LinkChecker" },
              { name: "Dates", functionName: "dateChecker" });               
  
  
   // Add it to the menu
   ss.addMenu( "Repopulate", repopulateEntries );
   ss.addMenu( "Wipe", wipeEntry );
   ss.addMenu( "Check", check );
   
   
   // Security
   recreateDeletedSheets();
   protectSheets();

 }



function onEdit( e ) 
{
  
  if( ss.getActiveSheet().getName() == "SELECTION" )
  {
    
    var account_cell = selection_sheet.getRange('C2');
    var term_cell = selection_sheet.getRange('C3');
    var courseName_cell = selection_sheet.getRange('C4');  
    var id_cell = selection_sheet.getRange('C5');
    var editedCell = e.range.getA1Notation();
    
    
    var account = account_cell.getValue();
    var term = term_cell.getValue();
    var courseName = courseName_cell.getValue();
    var courseID = id_cell.getValue();
    
    
    
    // Initialize Account and Term value
    if( account == '' )
    {
      account_cell.setFontStyle( 'italic' );
      account_cell.setValue( 'Select One' );
    }
    else if( account != '' && account != 'Select One' )
      account_cell.setFontStyle( 'normal' );
        
        
    if( term == '' )
    {
      term_cell.setFontStyle('italic');
      term_cell.setValue( 'Select One' );
    }
    else if( term != '' && term != 'Select One' )
      term_cell.setFontStyle( 'normal' );    
    
 
    
    // Handle Course Name Errors and formats
    if( courseName != "Repopulate" && courseName != "Pick a course" && courseName != "Processing..." && courseName != "" && courseName != 'No Courses Available' )    // if its a course name
    {
        id_cell.setFontStyle( 'normal' );
        courseName_cell.setFontStyle( 'normal' );
        
      if( id_cell.getFormula() == '' )
        id_cell.setFormula( '=vlookup(C4,Lists!G:H,2)' );
    }
    else if( courseName == "Pick a course" || courseName == "Processing..." )                  // If we course selction options clear id cell
      id_cell.setValue( '' );
    
    if( courseName == "" && typeof courseID != 'number' )                                      // If no course name is present and courseID is not a number
    {   
      id_cell.setValue( "type id or do above" );
      id_cell.setFontStyle( 'italic' );
    }    
    

    
    // Handle Course ID Errors and formats
    if( courseID == "Course ID" )
    {      
      id_cell.setValue( "Not a course" );
      id_cell.setFontStyle( 'italic' );
    }
    
    if( courseID == "#N/A" )
    {   
      id_cell.setValue( "type id or do above" );
      id_cell.setFontStyle( 'italic' );
    }  
    
    if( typeof courseID == 'number' )
      id_cell.setFontStyle( 'normal' );
      
    if( typeof courseID != 'number' && courseID != "type id or do above" && courseID != "" && courseID != "Not a course" )
      id_cell.setValue( "" );
      
      
 
    // If the user changes account or term start from scratch
    if( ( editedCell == 'C2' || editedCell == 'C3' ) && courseName != "Repopulate" ) 
    {
      courseName_cell.setDataValidation( SpreadsheetApp.newDataValidation().requireValueInList( [] ).build() );
      courseName_cell.setValue( 'Repopulate' );
      courseName_cell.setFontStyle( 'italic' );  
      
      id_cell.setValue( 'type id or do above' );
      id_cell.setFontStyle( 'italic' );
    } 
 
    
  }
      
}



/**************************************************************************************************
 * Function: recreateDeletedSheets
 * Purpose: recreate needed sheets in the event that they were deleted
 **************************************************************************************************/

function recreateDeletedSheets() 
{

  // If the sheet doesn't exist create it
  if( ! credentials_sheet )
  {
     ss.insertSheet( 2 ).setName('Credentials'); 
     
     var title_cell = ss.getRange('A1:C1');
     
     // Insert Titles
     title_cell.setValues( [["Key", "Credetial (FILL IN THIS COLUMN)", "Epires"]] );
     title_cell.setFontSize( 12 );
     title_cell.setFontWeight( 'bold' );
     title_cell.setVerticalAlignment( "middle" );
     
      
     // Set the Column Width 
     ss.setColumnWidth( 1 , 350 );
     ss.setColumnWidth( 2 , 700 );

     // Fields
     ss.getRange('A2').setValue("Secret Token (DO NOT SHARE THIS WITH ANYONE)");

  }//end-if
  
  if( ! selection_sheet )  
  {
     ss.insertSheet( 3 ).setName('SELECTION'); 
     
     var title_cell = ss.getRange("A1");
     var account_cell = ss.getRange("C2");
     var term_cell = ss.getRange("C3");
     var courseName_cell = ss.getRange("C4");
     var selection = ss.getRange('B2:B5');
 
     
     title_cell.setValue( "SELECT YOUR COURSE" );
     title_cell.setFontSize( 14 );
     title_cell.setFontWeight( 'bold' );
     title_cell.setHorizontalAlignment( "center" );
     
     
     // Insert and format Field Titles
     selection.setValues( [["Account"], ["Term"], ["Course Name"], ["Course ID"]] );
     selection.setFontSize( 10 );
     selection.setFontWeight( 'bold' );
     selection.setBackgrounds( [ ['#d9d2e9'], ['#d9d2e9'], ['#d9d2e9'], ['#d9d2e9'] ] );     // Light purple:  #d9d2e9
     
     ss.getRange('B5').setFontColor('#008000');
     
     // Format data fields
     account_cell.setBackground('#f3f3f3');                                                  // Light gray: #f3f3f3
     courseName_cell.setBackground('#f3f3f3');
     
     
     
     // Set the Column Width 
     ss.setColumnWidth( 1 , 237 );
     ss.setColumnWidth( 2 , 120 );
     ss.setColumnWidth( 3 , 140 );
     
     
     
     var accountNames = lists_sheet.getRange( "A2:A" + ss.getDataRange().getNumRows() + 2 );
     var termNames = lists_sheet.getRange( "D2:D" + ss.getDataRange().getNumRows() + 2 );
     
 
     var accountList = SpreadsheetApp.newDataValidation().requireValueInRange( accountNames ).build();       //https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder
     var termList = SpreadsheetApp.newDataValidation().requireValueInRange( termNames ).build();
 
     account_cell.setDataValidation( accountList );
     term_cell.setDataValidation( termList );  

  }//end-if
  
  if( ! course_sheet ) 
     ss.insertSheet( 4 ).setName('Course');     
  if( ! assignments_sheet ) 
     ss.insertSheet( 5 ).setName('Assignments');
  if( ! pages_sheet ) 
     ss.insertSheet( 6 ).setName('Pages');
  if( ! discussions_sheet ) 
     ss.insertSheet( 7 ).setName('Discussions');     
  if( ! announcements_sheet ) 
     ss.insertSheet( 8 ).setName('Announcements');
  if( ! modules_sheet ) 
     ss.insertSheet( 9 ).setName('Modules');
  if( ! folders_sheet ) 
     ss.insertSheet( 10 ).setName('Folders');
  if( ! files_sheet ) 
     ss.insertSheet( 11 ).setName('Files');
     
  if( ! dates_sheet )
  {
     ss.insertSheet( 12 ).setName('Dates'); 
     
     var title_cell = ss.getRange('A1:C1');
     var weeks_cell = ss.getRange('A2:A16');
     
     // Insert Titles
     title_cell.setValues( [["Week #", "Start", "End"]] );
     title_cell.setFontSize( 12 );
     title_cell.setFontWeight( 'bold' );
     title_cell.setHorizontalAlignment( "center" );
     title_cell.setVerticalAlignment( "middle" );
     
     // And week numbers
     for( var week = 1; week < 16; week++ )
       ss.getRange('A' +  ( week + 1 ) ).setValue( 'Week ' + week );
     
     // Plus side-note
     ss.getRange('E2').setValue("Populate for the current semester.  Leave off the hours.");
     
  }//end-if
  
  if( ! lists_sheet ) 
  {
     ss.insertSheet( 13 ).setName( 'Lists' );
     
     var accounts = paginatingCallToCanvas( URL.split('courses/')[0] + 'accounts', [ "name", "id" ] );
     var terms = [
                   ['2013S7A', 419],
                   ['2013SP', 445],
                   ['2013S7B', 426],
                   ['2013U7A', 605],
                   ['2013SU', 608],
                   ['2013UM1', 606],
                   ['2013U7B', 607],
                   ['2013TAF', 965],
                   ['2013F7A', 856], 
                   ['2013FM1', 861],
                   ['2013FA', 837],
                   ['2013FA9', 838],
                   ['2013F7B', 854],
                   ['2013FM2', 862],
                   ['2013FA11', 839],
// 2014                    
                   ['2014S7A', 1061],
                   ['2014SM1', 1064],
                   ['2014S7B', 1062],
                   ['2014SM2', 1065],
                   ['2014U7A', 1171],
                   ['2014SU', 1170],
                   ['2014UM1', 1173],
                   ['2014U7B', 1172],
                   ['2014TAF', 1550],
                   ['LEAD_2014FA', 1784],
                   ['2014F7A', 1521],
                   ['2014FM1', 1523],
                   ['2014FA', 1282],
                   ['2014FE1', 1577],
                   ['2014F7B', 1522],
                   ['2014FM2', 1524], 
                   ['2014FE2', 1578],
// 2015                    
                   ['2015TAS', 1906],
                   ['2015SM1', 1895],
                   ['2015SE1', 1897],
                   ['2015S7A', 1893],
                   ['2015SP', 1841],
                   ['2015SE2', 1898],
                   ['2015SM2', 1896],
                   ['2015S7B', 1894],
                   ['Default Term', 130],
                   ['PATHe', 863]
         ];
     

     
     var title_cell = ss.getRange('A1:H1');
  
     
     // Insert Titles
     title_cell.setValues( [[ "Account", "Account ID", "", "Term", "Term ID", "", "Course", "Course ID" ]] );
     
     // Format
     title_cell.setFontWeight( 'bold' );
     title_cell.setVerticalAlignment( 'middle' );
     
     ss.setRowHeight( 1, 25 );

    
     // Insertion
     ss.getActiveSheet().getRange( 2, 1, accounts.length, 2 ).setValues( accounts );    /* Accounts */
     ss.getActiveSheet().getRange( 2, 4, terms.length, 2 ).setValues( terms );          /* Terms */
      
  }
}




/**************************************************************************************************
 * Function: protectSheets
 * Purpose: Protect and hide sensitive data
 **************************************************************************************************/

function protectSheets() 
{
  var readme_permissions = readme_sheet.getSheetProtection();
  var cred_permissions = credentials_sheet.getSheetProtection();
  var user = readme_permissions.getUsers().length - 1;


  // For all members of the README sheet
  for( ;; )   
  {
      
    var editer = readme_permissions.getUsers()[ user ];
    
    if( editer != "cchaffatt@champlain.edu" && editer != "jblumberg@champlain.edu" )       // Prevent editing of README unless you are Carl or Josh 
    {
      readme_permissions.removeUser( editer );
      user = readme_permissions.getUsers().length - 1;
    }
    else if( user < 1 )
      break;

    else
       user--;
  }
  
  
  // Set Protection
  readme_permissions.setProtected( true );
  cred_permissions.setProtected( true );
  
  
  // Protect Sheets
  readme_sheet.setSheetProtection( readme_permissions );           
  credentials_sheet.setSheetProtection( cred_permissions );
  dates_sheet.setSheetProtection( cred_permissions );
  lists_sheet.setSheetProtection( cred_permissions );
  
  
  
  // Hide Sheets and rows/ columns
  credentials_sheet.hideSheet();
  credentials_sheet.hideRows( 2 );                          // Hide token row
  lists_sheet.hideSheet();
  lists_sheet.hideColumns( 7, 2 );
  

}



/**************************************************************************************************
 * Function: getSelection
 * Purpose: Gets the selection of courses using the options given on the 'SELECTION' sheet.
 * @return an array of courses
 **************************************************************************************************/

function getSelection()
{
 
  var userAccount = selection_sheet.getRange('C2').getValue();                      // The user selected account
  var userTerm = selection_sheet.getRange('C3').getValue();                         // The user selected term
    
  
  if( userAccount != "Select One" && userTerm != "Select One" && token != '' )      // If the user selected an option for account and term and has a token let's go to work
  {
    var rows = lists_sheet.getDataRange().getNumRows() ;
    
    
    selection_sheet.getRange('C4').setFontStyle( 'italic' );
    selection_sheet.getRange('C4').setValue( 'Processing...' );                     // Tell the user there will be a wait
    

    for( var account = 1; account < rows; account++ )
      if( userAccount == lists_sheet.getRange( "A" + ( account + 1 ) ).getValue() ) 
      {
        userAccount = lists_sheet.getRange( "B" + ( account + 1 ) ).getValue();
        break;
      }
        
    for( var term = 1; term < rows; term++ )
      if( userTerm == lists_sheet.getRange( "D" + ( term + 1 ) ).getValue() ) 
      {
        userTerm = lists_sheet.getRange( "E" + ( term + 1 ) ).getValue();
        break;
      }
   
    return paginatingCallToCanvas( URL.split('courses/')[0] + 'accounts' + '/' + userAccount + '/courses?enrollment_term_id=' + userTerm + '&per_page=100', [ "name", "id" ] ); 
  
  }//end-if
  else if( token == '' )
    SpreadsheetApp.getUi().alert( "Please see instructions on 'README' and supply a token to the 'Credentials' sheet." );
  
  else   // Alert, if no options were given
    SpreadsheetApp.getUi().alert( 'Please, select options for both of the following ( on \'SELECTION\' sheet ):\n - Account\n - Term' );
   
  
}



/**************************************************************************************************
 * Function: updateSelection
 * Purpose: Puts selection of courses that the user chose into the spreadsheet under the dropbox 'Course Names'
 **************************************************************************************************/

function updateSelection()
{

  try                                                                                   // See if the user satified the account and term requirement
  {
    var selection = getSelection();
    var courseNames = lists_sheet.getRange( "G2:G" + ( selection.length + 1 ) );
    var courseName_cell = selection_sheet.getRange('C4');
    var courseNameValidation = SpreadsheetApp.newDataValidation();                      // Reference:  https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder
  }
  catch( error )
  {  
   return;
  }
    
 
  // Data Validation   
  courseNameValidation.setAllowInvalid( true );
  courseName_cell.setDataValidation( courseNameValidation.requireValueInRange( courseNames ).build() ); 
  
                                                                                                                        
  // Insertion
  try                                                                                   // If term-account pair's retrieved data is empty let user know 
  {
    lists_sheet.getRange( 2, 7, selection.length, 2 ).setValues( selection );
  }
  catch( error )
  {
    // Let user know something failed
    courseName_cell.setFontStyle( 'italic' );
    courseName_cell.setValue( 'No Courses Available' );
    selection_sheet.getRange( 'C5' ).setValue( '' );
    return;
  }
  
  
  //Hide Insertion
  lists_sheet.hideColumns( 7, 2); 
  
  
  // Erase the rest of data on 'Lists' sheet
  lists_sheet.getRange( selection.length + 1, 7, lists_sheet.getMaxRows(), 2 ).clear();
  
  
  // Let user know they may proceed
  courseName_cell.setValue( 'Pick a course' );
  selection_sheet.getRange( 'C5' ).setValue( '' );
      
}





/**************************************************************************************************
 * Function: getAssignments
 * Purpose: Gets assignments for a particular course using the canvas api and returns data in an array.
 * @return an array of assignment object items
 **************************************************************************************************/

function getAssignments()
{
  var url = URL + '/assignments' + ITEMS_PER_PAGE;                       //Pass global URL to url so we can change it
  
  var apiParameters = [ "name",
                        "html_url",
                        "published",
                        "muted",
//Dates
                        "updated_at",                        
                        "unlock_at",
                        "due_at",
                        "lock_at",                       
//Turnit in
                        "turnitin_enabled",
                        "turnitin_settings",
//Groups
                        "group_category_id",
                        "grade_group_students_individually",
                        "peer_reviews",
//Grading
                        "grading_type",
                        "points_possible",
//Submissions                        
                        "has_submitted_submissions",
                        "needs_grading_count",
                        "submission_types",
                        "allowed_extensions",                        
//Rubric
                        "rubric_settings",
                        "rubric",
                        "use_rubric_for_grading"
                      ];
                      
  return paginatingCallToCanvas( url, apiParameters );
  
}




/**************************************************************************************************
 * Function: updateAssignments
 * Purpose: Using the assignment data from the canvas api to import/ fill data cells on the 'Assigments' Sheet.
 **************************************************************************************************/

function updateAssignments()
{
  var assignments = getAssignments();
  var titles = [
    ["Assignment Name",
     "Assigment Link",
     "Published?",
     "Muted?",
     "Last Update",
     "Unlock Date",
     "Due Date",
     "Lock Date",
     "Turnit in Enabled?",
     "Turnit in Settings",
     "Group Assignment?",
     "Grade Group Students Individually",
     "Peer Review Require?",
     "Grading Type",
     "Possible Points", 
     "Has Submissions?",
     "Submissions that Need Grading",
     "Submission Type",
     "Allowed Extensions",
     "Rubric Settings",
     "Rubric",
     "Uses Rubric for Grading?",
     "Broken Links"]
  ];
  
  formatCells( assignments, assignments_sheet, titles );
  
}






/**************************************************************************************************
 * Function: getPages
 * Purpose: Gets 'Pages' object from a particular course using the canvas api and returns data in an array.
 * @return an array of page object items
 **************************************************************************************************/

function getPages()
{
  var url = URL + '/pages' + ITEMS_PER_PAGE;                       // Pass global URL to url so we can change it
  
  var pagesData = [];                                              // Stores page data

  
  var apiParameters = [ "title",
                        "",
                        "url",
                        "published",
                        "front_page",
                        "hide_from_students",
//Dates
                        "created_at",
                        "updated_at",
                        "lock_at",
//Roles
                        "last_edited_by",
                        "editing_roles"
                      ];
                      
  
  var pages = paginatingCallToCanvas( url, apiParameters );
  
 
   // Swap pages list field with page html url
    for(var i = 0; i < pages.length; i++)    
    {     
      var page = URL.split( 'api/v1/' );
      pages[ i ][ 1 ] = page[ 0 ] + page[ 1 ] + '/pages/' + pages[ i ][ 2 ];            // Insert page url into empty array slot    
    }
    
    return pages;     
}




/**************************************************************************************************
 * Function: updatePages
 * Purpose: Using the pages data from the canvas api to import/ fill data cells on the 'Pages' Sheet.
 **************************************************************************************************/

function updatePages()
{
  var pages = getPages();
  
  var titles = [
    ["Page Title",
     "Page Link",
     "URL",
     "Published?",
     "Front Page?",
     "Hidden from Students?",
     "Created At",
     "Last Update",
     "Lock Date",
     "Last Edited By",
     "Editing Roles",
     "Broken Links"]
  ];
  
  formatCells( pages, pages_sheet, titles );
  
}






/**************************************************************************************************
 * Function: getDiscussions
 * Purpose: Gets discussions for a particular course using the canvas api and returns data in an array.
 * @return an array of discussion object items
 **************************************************************************************************/

function getDiscussions()
{
  var url = URL + '/discussion_topics' + ITEMS_PER_PAGE;                       // Pass global URL to url so we can change it
  
  var apiParameters = [ "title",
                        "html_url",
                        "published",
//State
                        "discussion_type",
                        "pinned",
                        "podcast_url",
                        "read_state",
                        "require_initial_post",
                        "discussion_subentry_count",                             
//Dates
                        "posted_at",
                        "delayed_post_at",
                        "locked",
                        "lock_at",
//User
                        "user_name"
                      ];
                      
  return paginatingCallToCanvas( url, apiParameters );
  
}




/**************************************************************************************************
 * Function: updateDiscussions
 * Purpose: Using the assignment data from the canvas api to import/ fill data cells on the 'Discussions' Sheet.
 **************************************************************************************************/

function updateDiscussions()
{
  var discussions = getDiscussions();
  var titles = [
    ["Title",
     "Discussion Link",
     "Published?",
     "Discussion Type", 
     "Pinned?",
     "Podcast url",
     "Read State",
     "Require Initial Post?",
     "Number of Posts?",
     "Posted At", 
     "Delayed Post At",
     "Locked?",
     "Locked At",
     "Creator",
     "Broken Links"]
  ];
  
  
  formatCells( discussions, discussions_sheet, titles );
  
}





/**************************************************************************************************
 * Function: getCourse
 * Purpose: Gets 'course' object for a particular course using the canvas api and returns data in an array.
 * @return an array of course object items
 **************************************************************************************************/

function getCourse()
{
  var apiParameters = [ "name",
                        "course_code",
                        "id",
                        "workflow_state",
                        "default_view",
                        "course_format",
                        "term",
                        "license",
                        "is_public",
                        "public_description",
                        "storage_quota_mb",
                        "hide_final_grades",
//Privileges                                                
                        "allow_student_assignment_edits",
                        "allow_wiki_comments",
                        "allow_student_forum_attachments",
//Dates
                        "start_at",
                        "end_at",
//Enrollment                 
                        "enrollments",
                        "open_enrollment",
                        "self_enrollment",
                        "restrict_enrollments_to_course_dates",
                        "apply_assignment_group_weights",                             
//Syllabus
                        "public_syllabus"
                      ];
  
  
  return singleCallToCanvas( URL + '?include[]=term&include[]=total_scores&include[]=course_progress', apiParameters ); 
  
}





/**************************************************************************************************
 * Function: updateCourse
 * Purpose: Using the course data from the canvas api to import/ fill data cells on the 'Course' Sheet.
 **************************************************************************************************/

function updateCourse()
{
  var course = getCourse();
  
  var titles = [
      "Course Title",
      "Course Code",
      "Course ID",
      "Current State", 
      "Default View",    
      "Course Format",
      "Term",
      "License",
      "Is Public",
      "Public Description",
      "Storage Quota (MB)",
      "Hide Final Grades?",     
      "Allow Student Assignment Edits",
      "Allow Wiki Comments",
      "Allow Student Forum Attachments",
      "Start Date",
      "End Date", 
      "Enrollments",
      "Open Enrollment?",
      "Self Enrollment?",
      "Restrict Enrollment to Course Dates",
      "Weighted Assignment Groups?",
      "Public Syllabus?" ,
      "Broken Syllabus Links"
  ];
  

  course_sheet.setColumnWidth(1, 220);          //Set the title row's height
 
  
  // Dimensions
  var rows = course.length + 1;
  
  var fullRange = course_sheet.getRange( 1, 1, rows, 2 );
 
  
  // 2 - D Format matrices
  var backgrounds = fullRange.getBackgrounds();
  var fonts = fullRange.getFontWeights();
  var formats = fullRange.getNumberFormats();
  var horizontals = fullRange.getHorizontalAlignments();
  var verticals = fullRange.getVerticalAlignments();
  
  
  
  var course2D = [];
  var titles2D = [];
  
  for( var j = 0; j < rows; j++ )
  {
    var title = titles[ j ];
          
    if( j < course.length )                          // Post processing for data array  
    {   
      
      var value = course[ j ];  
      
        
/* Substitutions *
 *****************/                            
         
      // Term  
      if( title == "Term" )
        course[ j ] = value[ "name" ];
           
      // Dates  
      if( title == "Start Date" ||  title == "End Date" )
      {
        course[ j ] = UTCtoCurrentTimeZone( value );
        formats[ j ][ 1 ] = 'm/d/yyyy h:mm am/pm';                               
      }
              
      // No data
      if(  /^undefined/i.test( course[ j ] ) || course[ j ] == null )
        course[ j ] = "-"; 
  
        
/* Format Cells *
 ****************/ 
 
      // Zebra Stripes
      if( ( j + 1 ) % 2 == 0 )
        backgrounds[ j ][ 1 ] = STRIPE_COLOR;
      
      // Boolean Colors 
      if( /^false/i.test( course[ j ] )  )
      {     
        backgrounds[ j ][ 1 ] = '#6fa8dc';
        fonts[ j ][ 1 ] = 'bold';
      }
      
      if( /^true/i.test( course[ j ] ) )
      {
        backgrounds[ j ][ 1 ] = 'green';
        fonts[ j ][ 1 ] = 'bold';
      }     
  
      
      course2D.push( [ course[ j ] ] );        // Make 2-D and transpose
      
    }//end-if: post data handling

                                      
    // Data
    horizontals[ j ][ 1 ] = "left";
    verticals[ j ][ 1 ] = "middle";

    
    
    // Titles
    titles2D.push( [ title ] );               // Make 2-D and transpose
    
    fonts[ j ][ 0 ] = 'bold';
    horizontals[ j ][ 0 ] = "left";
    verticals[ j ][ 0 ] = "middle"; 

    // Broken Links 
    if( j == course.length )                  // Indiscriminately color the last cell
    {
      backgrounds[ j ][ 1 ] = '#989898';
      course_sheet.getRange("B24").setValue('-');
    }
    
  }//end-for
  
  
  // Insertion 
  course_sheet.getRange( 1, 1, rows ).setValues( titles2D );                    // Titles
  course_sheet.getRange( 1, 2, rows - 1 ).setValues( course2D );                // Data 
  

  // Set All
  fullRange.setFontWeights( fonts );
  fullRange.setBackgrounds( backgrounds );
  fullRange.setNumberFormats( formats );
  fullRange.setHorizontalAlignments( horizontals );
  fullRange.setVerticalAlignments( verticals );
 

  // Hide The Rest
  course_sheet.hideColumns(  3 , ( course_sheet.getMaxColumns() + 1 ) - 3 ); 
  course_sheet.hideRows( rows + 1, ( course_sheet.getMaxRows() + 1 ) - ( rows + 1 ) );
      
}




/**************************************************************************************************
 * Function: getAnnouncements
 * Purpose: Gets Announcements for a particular course using the canvas api and returns data in an array.
 * @return array of announcement items
 **************************************************************************************************/

function getAnnouncements()
{
  var url = URL + '/discussion_topics?only_announcements=1' + ITEMS_PER_PAGE;                       //Pass global URL to url so we can change it
  
  var apiParameters = [ "title",
                        "html_url",
                        "published",
                        "discussion_type",
                        "require_initial_post",
                        "podcast_url",
                        "podcast_has_student_posts",
//State                      
                        "discussion_subentry_count",                    
                        "topic_children",
                        "attachments",                        
                        "read_state",
//Dates
                        "delayed_post_at",
                        "last_reply_at",
                        "locked",
                        "lock_at",
//Message
                        "author"
                      ];
                      
                      
  return paginatingCallToCanvas( url, apiParameters );    
  
}




/**************************************************************************************************
 * Function: updateAnnouncements
 * Purpose: Using the Announcements object data from the canvas api to import/ fill data cells on the 'Announcements' Sheet.
 **************************************************************************************************/

function updateAnnouncements()
{
  var announcements = getAnnouncements();
  var titles = [
    ["Announcement",
     "Link",
     "Published?",
     "Discussion Type",
     "Require Initial Post?",
     "Podcast url",
     "Podcast Has Student Posts",
     "Subentry Count",
     "Topic Children",
     "Attachments",
     "Read State",
     "Delayed Post At",
     "Last Reply At",
     "Locked?",
     "Locked At",
     "Creator",
     "Broken Links"]
  ];
                 
 
  formatCells( announcements, announcements_sheet, titles );
  
}






/**************************************************************************************************
 * Function: getModules
 * Purpose: Gets Modules object for a particular course using the canvas api and returns data in an array.
 * @return array of module items
 **************************************************************************************************/

function getModules()
{
  var url = URL + '/modules?include[]=items&per_page=100';                       //Pass global URL to url so we can change it
  
  var apiParameters = [ "name",
                        "",
                        "items_url",
                        "items_count",
                        "items",
                        "workflow_state",
                        "published",
                        "publish_final_grade",
                        "require_sequential_progress",                    
                        "prerequisite_module_ids",
                        "unlock_at"
                      ];
                      
 
  
  var modules = paginatingCallToCanvas( url, apiParameters );
  
 
  // Insert modules URL in empty slot in modules array 
  for(var i = 0; i < modules.length; i++)
  {               
    module = URL.split('api/v1/');
    modules[ i ][ 1 ] = module[0] + module[1] + "/modules#module_" + modules[ i ][ 2 ].split( URL + '/modules/' )[1].split('/items')[0];  
  }
 
    
  return modules;     
  
}




/**************************************************************************************************
 * Function: updateModules
 * Purpose: Using the Modules object data from the canvas api to import/ fill data cells on the 'Modules' Sheet.
 **************************************************************************************************/

function updateModules()
{
  var modules = getModules();
  var titles = [
    ["Module",
     "Module Link",
     "Items URL",
     "Items Count",
     "Items",
     "Workflow State",
     "Published",
     "Publish Final Grade Upon Completion",
     "Require Seqential Progress",
     "Pre-requisite Modules",
     "Unlock At",
     "Broken External Urls"]
  ];
  
  formatCells( modules, modules_sheet, titles );
  
}




/**************************************************************************************************
 * Function: getFolders
 * Purpose: Gets Folders object for a particular course using the canvas api and returns data in an array.
 * @return array of folder items
 **************************************************************************************************/

function getFolders()
{
  
  var allFolders = [];                                         // Store all folders
  var folder = 0;                                              // Initialiize folder index
  
  var root = urlFetchCall( URL + '/folders/by_path' )[ 1 ];       
  var url = URL.split('/api/v1');
  
  
  var apiParameters = [ "files_url",
                        "position",
                        "hidden",
                        "locked",
                        "created_at",
                        "lock_at",
                        "updated_at",
                        "folders_count",
                        "files_count"                     
                      ];   
   
  
  
  recurseFolders( root[ 0 ][ "id" ] );                          // Recursively scan (sub)folders  
  

  return allFolders;
  
  
  /**************************************************************************************************
  * Sub-function: recurseFolders
  * Purpose: Scans all subfolder and put data about each folder in allfolders.
  **************************************************************************************************/  
  
  function recurseFolders( id )
  {
    var index = 0;                                                                         
    var folders = urlFetchCall( URL.split('courses/')[ 0 ] + 'folders/' + id + '/folders' + ITEMS_PER_PAGE )[ 1 ];
 
    while( index < folders.length )                                                                 // While we have more subfolder keep working
    {
      
      recurseFolders( folders[ index ][ "id" ] );                                                   // recursive find all subfolders
      
      allFolders[ folder ] = []; 
      allFolders[ folder ].push( folders[ index ][ "full_name" ].split( 'course files/' )[ 1 ] );   // Insert path from root               
      allFolders[ folder ].push( url[0] + url[1] + '/files' );                                      // Insert Link to files section
      
      for( var key in apiParameters )     
        allFolders[ folder ].push( folders[ index ][ apiParameters[ key ] ] );                      // Push api values       
      
      index++;
      folder++;
      
    }//end-while: sub-folder loop
    
  }//end-function: recurseFolders
  
  
}




/**************************************************************************************************
 * Function: updateFolders
 * Purpose: Using the folders object data from the canvas api to import/ fill data cells on the 'Folders' Sheet.
 **************************************************************************************************/

function updateFolders()
{
  var folders = getFolders();

  var titles = [
    [  "Folder (Path)",
       "Files Link",
       "Files URL",
       "Position",
       "Hidden",   
       "Locked?",
       "Created At",
       "Locked At",
       "Updated At",
       "Folder Count",
       "Files Count",     
       "Broken Folders"
     ]
  ];

                    
  formatCells( folders, folders_sheet, titles );
  
  // Clear and hide excess
  folders_sheet.getRange( 1, titles[ 0 ].length, folders_sheet.getDataRange().getNumRows() ).clear();
  folders_sheet.hideColumns( titles[ 0 ].length );
  
}




/**************************************************************************************************
 * Function: getFiles
 * Purpose: Gets Files objects for a particular course using the canvas api and returns data in an array.
 * @return array of file items
 **************************************************************************************************/

function getFiles()
{

  var apiParameters = [ "display_name",
                        "url",
                        "filename",
                        "",
                        "size",
                        "hidden",
                        "locked",
                        "lock_at",
                        "created_at",
                        "updated_at"                        
                      ];
  
  
  var rows = folders_sheet.getDataRange().getNumRows();                         // Rows on the 'Folders' sheet
  var ff = [];                                                                  // Filed Files
  
  for( var i = 2; i <= rows; i++ )                                              // For each row on the 'Folders' sheet
  {
    var url = folders_sheet.getRange( 'C' + i ).getValue() + ITEMS_PER_PAGE;
    var files = paginatingCallToCanvas( url, apiParameters );
    
    
    for( var file in files )
    {
      files[ file ][ 3 ] = folders_sheet.getRange( 'A' + i ).getValue();        // Insert folder path
      ff.push( files[ file ] );
    }
    
  }
  
 
  var allFiles = paginatingCallToCanvas( URL + '/files' + ITEMS_PER_PAGE, apiParameters );
  
  for( var item in allFiles )
  {
    var found = 0;
    var orphanFile = allFiles[ item ];                        
    
    for( var filed in ff )
      if( orphanFile[ 4 ] == ff[ filed ][ 4 ] )                                // If we find the same file size in allFiles as ff then the files must be the same
        found++;                                                               // Mark that we found the file ( has a folder )
   
    
    if( found == 0 )                                                           // If we didn't find the file it doesn't belong to a folder
      ff.push( orphanFile )
  }
  
  return ff;
}




/**************************************************************************************************
 * Function: updateFiles
 * Purpose: Using the Files object data from the canvas api to import/ fill data cells on the 'Files' Sheet.
 **************************************************************************************************/

function updateFiles()
{
  var files = getFiles();

  var titles = [
    [  "Display Name",
       "File Link",
       "File Name",
       "Folder (Path)",
       "Size (KB)",
       "Hidden",
       "Locked?",
       "Locked At",
       "Created At",
       "Updated At",     
       "Broken Files"
     ]
  ];


  formatCells( files, files_sheet, titles );
  
}



/**************************************************************************************************
 * Function: updateFoldersAndFiles
 * Purpose: Fill cells on the 'Folders' and 'Files' Sheet using canvas api data.
 **************************************************************************************************/
function updateFoldersAndFiles()
{
  updateFolders();
  updateFiles();
}




/**************************************************************************************************
 * Function: repopulateAll
 * Purpose: Fill all sheets using canvas api data.
 **************************************************************************************************/
function updateAll()
{
  updateCourse();
  updateAssignments();
  updatePages();
  updateDiscussions();
  updateAnnouncements();
  updateModules();
  updateFoldersAndFiles();
}



/*************************************************
 * Function: urlFetchCall
 * Purpose: Retrieves data and a http response from the given url using the global options.
 * @param url the url data will be drawn from 
 *************************************************/
function urlFetchCall( url )
{
   
    // Response
    var response = UrlFetchApp.fetch( url, options );
    
    
    //Put data into the variable
    return [ response.getAllHeaders(), JSON.parse( response.getContentText() ) ];
  
}



/*************************************************
 * Function: singleCallToCanvas
 * Purpose: Retrieves a response from canvas storing the contents in an array.
 * @param url the url data will be drawn from
 * @api a list of api calls
 * @return data the state of the api querys in an array 
 *************************************************/

function singleCallToCanvas( url, api )
{
   
    var data = [];                                     //Stores all data
    
    // Response
    var response = urlFetchCall( url );
    
    //Parse response and store in the fields variable
    var fields = response[1];
                                       
      
    for( key in api )
      data.push( fields[ api[ key ] ] );                //Store all parameters for a single field                                            
     
  
    return data;
  
}




/**************************************************************************************************
 * Function: paginatingCallToCanvas
 * Purpose: Retrieves a response from canvas and checks to see if there's any other links that need data drawn from.
 * @url the url to query
 * @api an array of api parameters
 * @return data an array of api reponses
 **************************************************************************************************/

function paginatingCallToCanvas( url, api )
{ 
  
  var next = 'rel=\\\"next\\\"';                                             //Check for next page
  var linkreg = /\s*>\s*;\s*|^\s*"\s*|\s*"$\s*|\s*,\s*<\s*|\s*"\s*<\s*/g;    //Delimit paginating links
  var data = [];                                                             //Stores all data
  var j = 0;                                                                 //Counter for for-loop

  //Handling Pagination
  do
  {
  
    // Response
    var response = urlFetchCall( url );
    
    
    //Get next link
    var links = JSON.stringify( response[0]['Link'] );
    var linkArr = links.split( linkreg );               //Regex to split string into array elements 
   
    //Index of next url
    var nextUrl = linkArr.indexOf( next ) - 1;
    
    //Parse response and store in the fields variable
    var fields = response[1];
    
    
    //Make sense of data and store in an array
    for(var i = 0; i < fields.length; i++)
    {
      data[ i + j ] = [];
      
      for( key in api )
        data[ i + j ].push( fields[ i ][ api[ key ] ] );
        
    }                                              
     
    j += fields.length;                                 //Save where we left off so we can insert new items in correct location
    
    if( nextUrl > 0 )                                   //If we have more pages let's store that url
      url = linkArr[ nextUrl ];
  }
  while( nextUrl > 0 );                                 //While we have more pages let's get more items
 
  return data;
}





/**************************************************************************************************
 * Function: formatCells
 * Purpose: Inserts data and formats a sheet.
 * @data the data to insert into the cells
 * @sheet the sheet to insert data into
 * @titles the titles of the sheet
 **************************************************************************************************/

function formatCells( data, sheet, titles )
{

  sheet.setRowHeight( 1, 40 );     // Set the title row's height
  sheet.setColumnWidth( 1, 273 );  // Set the oject name's column width
  sheet.setColumnWidth( 2, 110 );  // Set the URL column's width
  sheet.setFrozenColumns( 2 );     // Freeze name and link
  sheet.setFrozenRows( 1 );        // Freeze titles

  
  // Dimensions
  var rows = data.length + 1,
      cols = titles[ 0 ].length;
      
  var fullRange = sheet.getRange( 1, 1, rows, cols );
 
  
  // 2 - D Format matrices
  var backgrounds = fullRange.getBackgrounds();
  var fonts = fullRange.getFontWeights();
  var formats = fullRange.getNumberFormats();
  var horizontals = fullRange.getHorizontalAlignments();
  var verticals = fullRange.getVerticalAlignments();
  
  

  for( var j = 0; j < rows; j++ )
  {
    for(var k = 0; k < cols; k++)
    {
      
      var title = titles[ 0 ][ k ];
      

      if( j > 0  )                          // Post processing for data array  
      {  
        var value = data[ j - 1 ][ k ];     
         
/* Substitutions *
 *****************/            
 
         
        // Assignments  
        if( title == "Group Assignment?" )
          data[ j - 1 ][ k ] = ! isNaN( parseInt( value ) );        
          
        // Announcements & Pages
        if( title == "Creator" || title == "Last Edited By" )
          try
          {
            data[ j - 1 ][ k ] = value[ "display_name" ];
          }
          catch( err ){}
          
        // Files
        if( title == "Size (KB)" )
        {
          data[ j - 1 ][ k ] = data[ j - 1 ][ k ] / Math.pow( 2, 10 );              // Convert file size to kilobytes
          formats[ j ][ k ] = '#,##0';
        }
       
       
       // Dates  
        if( title == "Due Date" ||  title == "Last Update" || title == "Last Reply At" || title == "Unlock Date" || title == "Lock Date" || 
            title == "Unlock At" || title == "Lock At" || title == "Created At" || title == "Posted At" || title == "Delayed Post At" || title == "Updated At" )
        {
          data[ j - 1 ][ k ] = UTCtoCurrentTimeZone( data[ j - 1 ][ k ] );             
          formats[ j ][ k ] = 'm/d/yyyy h:mm am/pm';    
        }
            
        // No data
        if(  /^undefined/i.test( data[ j - 1 ][ k ] ) || data[ j - 1 ][ k ] == null )
          data[ j - 1 ][ k ] = "-"; 
          
          
        
/* Format Cells *
 ****************/ 
 
        // Zebra Stripes
        if( ( j + 1 ) % 2 == 0 )
          backgrounds[ j ][ k ] = STRIPE_COLOR;
          
        // Boolean Colors 
        if( /^false/i.test( data[ j - 1 ][ k ] )  )
        {
         
          backgrounds[ j ][ k ] = '#6fa8dc';
          fonts[ j ][ k ] = 'bold';
        }
        
        if( /^true/i.test( data[ j - 1 ][ k ] ) )
        {
          backgrounds[ j ][ k ] = 'green';
          fonts[ j ][ k ] = 'bold';
        }     
        
        // Broken Links 
        if( k == cols - 1 )                  // Indiscriminately color the last cell
          backgrounds[ j ][ k ] = '#989898';  
       
       
        // Item Position
        if( ( ( title == "Submission Type" || title == "Turnit in Settings" || title == "Rubric Settings" || title == "Rubric" || title == "URL" || title == "Creator" ||
                title == "File Name" || title == "Folder (Path)" ) && data[ j - 1 ][ k ] != "-" ) || ( k == cols - 1 && data[ j - 1 ][ k ] != "-" ) )
          verticals[ j ][ k ] = "middle";
        else if( k > 1 )  
        {
          horizontals[ j ][ k ] = "center";
          verticals[ j ][ k ] = "middle";
        }
       
      
      }//end-if: data array handling

 
      else if( j == 0 )                     // Format titles
      {
        fonts[ j ][ k ] = 'bold';
        horizontals[ j ][ k ] = "center";
        verticals[ j ][ k ] = "middle";
      }    
      
    }//end-for: columns
    
  }//end-for: rows
  
  
  // Insertion 
  try
  {
    sheet.getRange( 1, 1, 1, titles[ 0 ].length ).setValues( titles );               // Titles
    sheet.getRange( 2, 1, data.length, data[ 0 ].length ).setValues( data );         // Data 
  }
  catch( err ){}
  

  // Set All
  fullRange.setFontWeights( fonts );
  fullRange.setBackgrounds( backgrounds );
  fullRange.setNumberFormats( formats );
  fullRange.setHorizontalAlignments( horizontals );
  fullRange.setVerticalAlignments( verticals );
  


  // Hide The Rest
  sheet.hideColumns( cols + 1, ( sheet.getMaxColumns() + 1 ) - ( cols + 1 ) ); 
  sheet.hideRows( rows + 1, ( sheet.getMaxRows() + 1 ) - ( rows + 1 ) );
      
     
  if( sheet.getName() == "Modules" || sheet.getName() == "Pages" || sheet.getName() == 'Folders' )
    sheet.hideColumns( 3 );    

}





/*************************************************
 * Function: wipeSheet
 * Purpose: Clear active sheet
 *************************************************/

function wipeSheet()
{

  if( ss.getActiveSheet().getName() == "Credentials" || ss.getActiveSheet().getName() == "README" || ss.getActiveSheet().getName() == "Lists" ) 
  {}   //Do nothing
  
  else if( ss.getActiveSheet().getName() == "Dates" )
     ss.getActiveSheet().getRange( 2, 2, ss.getActiveSheet().getDataRange().getNumRows(), 2 ).clear(); 
     
  else
  {
    ss.getActiveSheet().clear();                                                                                      // Else clear the whole sheet
    ss.getActiveSheet().unhideRow( ss.getActiveSheet().getRange( 1, 1, ss.getActiveSheet().getMaxRows() ) );          // Unhide everything
    ss.getActiveSheet().unhideColumn( ss.getActiveSheet().getRange( 1, 1, 1, ss.getActiveSheet().getMaxColumns() ) ); 
    ss.getActiveSheet().setFrozenColumns( 0 );                                                                        // Unfreeze everything
    ss.getActiveSheet().setFrozenRows( 0 );
  }
  
}



/*************************************************
 * Function: wipeAll
 * Purpose: Clear all sheets that draw data
 *************************************************/

function wipeAll()
{
    var sheets = [ "course_sheet", "assignments_sheet", "pages_sheet", "discussions_sheet", "announcements_sheet", "modules_sheet", "folders_sheet", "files_sheet" ];
    
    for( var sheet = 0; sheet < sheets.length; sheet++ )
    {
      eval( sheets[ sheet ] + ".clear();" );                                                                                                // clear the whole sheet
      eval( sheets[ sheet ] + ".unhideRow( " + sheets[ sheet ] + ".getRange( 1, 1, " + sheets[ sheet ] + ".getMaxRows() ) );" );            // Unhide everything
      eval( sheets[ sheet ] + ".unhideColumn( " + sheets[ sheet ] + ".getRange( 1, 1, 1, " + sheets[ sheet ] + ".getMaxColumns() ) );" );
      eval( sheets[ sheet ] + ".setFrozenColumns( 0 );" );                                                                                  // Unfreeze everything
      eval( sheets[ sheet ] + ".setFrozenRows( 0 );" );
    }
    

}





/*************************************************
 * Function: checkHTMLForBrokenLinks
 * Purpose: Checks HTML for broken links.
 * @param content html to be scanned
 *************************************************/
function checkHTMLForBrokenLinks( content )
{
  var brokenlinks = "";
  
  try
  {
    var tags = content.match(/(((href\s*=\s*("|'))|(src\s*=\s*("|')))([^\"\']*)("|'))/g);
  }
  catch( error )
  {
   return '-';
  }
  
  if ( tags )
  { 
    for( var i = 0; i < tags.length; i++)
    {
      var link = tags[i].substring( tags[i].indexOf('"') + 1, tags[i].length - 1 );
      var code = HTTPResponse( link );
      
      
      if( code >= 400 && code < 600 || code == 0 )
        brokenlinks += link + " (" + code +"), ";
      
    }//end-for 
  }   
  
  if( brokenlinks == "" )
    return brokenlinks = "-";
    
  return brokenlinks;
}




/***************************************************************************************
* Function: HTTPResponse
* Purpose: Sub-function that returns the http server status code for a particular uri.
* @param uri the uri to be checked
* @return response_code the http server status code
***************************************************************************************/

function HTTPResponse( uri )
{
  var response_code;
  
  try{
    response_code = UrlFetchApp.fetch( uri ).getResponseCode().toString() ;
  }
  
  catch( error ){
    response_code = error.toString().match( / returned code (\d\d\d)\./ )[1] ;
  }
  
  finally{
    return response_code ;
  }
} 





/******************************************************************
 * Function: LinkChecker
 * Purpose: Finds gray cells turns them white and checks for links
 ******************************************************************/
var linksChecked = 0; 
 
function LinkChecker()
{     

  var sheetName = ss.getActiveSheet().getName(),
      sheet = ss.getActiveSheet(),
      rows = sheet.getDataRange().getNumRows(),
      cols = sheet.getDataRange().getNumColumns();
   
  
  for( var row = linksChecked + 2; row <= rows; row++ )         //Scan the rows starting at 2
  {
    
    var lastCell = String.fromCharCode( ( "A".charCodeAt() - 1 ) + cols ) + row;        //last column of rows
    
    
    //Look for grey cells and loop through them
    if( /^#989898/i.test( sheet.getRange( lastCell ).getBackground() ) )
    {
    
      /* Check the Course sheet *
       **************************/
      if( sheetName == "Course" )
      {
        
        //Single call to canvas to get html, check for broken links, then insert them into the sheet
        sheet.getRange( lastCell ).setValue( checkHTMLForBrokenLinks( singleCallToCanvas( URL + '?include[]=syllabus_body', ["syllabus_body"] )[0] ) );
       
      }
      
      /* Check the Modules sheet *
       ***************************/
      else if( sheetName == "Modules" )
      {
        var items = urlFetchCall( sheet.getRange( "C" + row ).getValue() + "?per_page=100" );
        var brokenlinks = "";
        
        for( var item = 0; item < items[1].length; item++ )
        {
          try{       
            var link = items[1][ item ][ "external_url" ];     
          }
          
          catch( error ){
            continue;
          }
          
          if( link )   
          {
            code = HTTPResponse( link );
            
            if( code >= 400 && code < 600 || code == 0 )
              brokenlinks += link + " (" + code + "), ";
          }
          
        }//end-for
        
        if( brokenlinks == "" )
          brokenlinks = '-';
         
         sheet.getRange( lastCell ).setValue( brokenlinks );    
 
      }
      
      /* Check the Files sheet *
       *************************/ 
      else if( sheetName == 'Files' )                                         // Attempt to check html files in file section
      {
        var fileName = sheet.getRange( 'C' + row ).getValue();
        if( /(.*.html?$)/i.test( fileName ) )                                 // If the file name ends with htm or html select it
        {         
          var url = sheet.getRange( 'B' + row ).getValue();    
          var response = UrlFetchApp.fetch( url.split('?download')[ 0 ], {muteHttpExceptions: true} );      
        }
      }
      
      /* Check the Other sheets *
       **************************/
      else                                                                  
      {
      
        if( sheetName == "Assignments" )  // Check the Assignments sheet
          var api = 0;
      
        if( sheetName == "Pages" )        // Check the Pages sheet
          api = 1;
        
        if( sheetName == "Discussions" || sheetName == "Announcements" )    // Check the Discussions and Announcements sheet
          api = 2;     
             
  
        var link = sheet.getRange( 'B' + row ).getValue();
        var idx = link.indexOf( 'courses/' );
        
        
        //Single call to canvas to get html, check for broken links, then insert them into the sheet
        sheet.getRange( lastCell ).setValue( checkHTMLForBrokenLinks( singleCallToCanvas( link.substr( 0, idx ) + 'api/v1/' + link.substr( idx ), [ ["description"], ["body"], ["message"] ][api] ) + "" ) );
       
      }       
      

      if( sheet.getRange( lastCell ).getValue() != '-' )
        sheet.getRange( lastCell ).setHorizontalAlignment("left");
      else if( sheet.getRange( lastCell ).getValue() == '-' && sheetName != "Course" )
        sheet.getRange( lastCell ).setHorizontalAlignment("center");
        
      sheet.getRange( lastCell ).setBackground( '#ffffff' );          // Indicate that link has been checked on sheet
      linksChecked++;                                                 // Record check
    
    }//end-if: search for gray cells
    
     
  }//end-for
  
  
  if( linksChecked == rows - 1 )   // if we finish checking links
    linksChecked = 0;
 
             
}





/******************************************************************
 * Function: UTCtoCurrentTimeZone
 * Purpose: Convert string of the format "yyyy-MM-ddTHH:mm:ssZ" to a date type. Then, convert UTC time, canvas' default timezone for api dates, to the local timezone
 *          https://canvas.instructure.com/doc/api/ 
 * @return a local date type from the string date
 ******************************************************************/
 
function UTCtoCurrentTimeZone( stringDate )
{

  try
  {
    var dateTime = stringDate.split('T');
    var date = dateTime[0].split('-');
    var time = dateTime[1].split(':');
  }
  catch( error )
  {
    return "-";
  }
   
 
  return new Date( Date.UTC( date[0], date[1] - 1 , date[2], time[0], time[1] ) );     // Input a UTC date that is convert to a local time zone date

}





/***************************************************************
 * Function: dateChecker
 * Purpose: Compare assignment dates to the official starting and ending week dates of the term. Highlights the corresponding cell
 *          depending on if the assignment date violates the dates sheet's date (in days). 
 *****************************************************************/

function dateChecker()
{  
  var sheet = ss.getActiveSheet(),
      rows = sheet.getDataRange().getNumRows(),
      cols = sheet.getDataRange().getNumColumns();                                         

  var unprovidedDates = 0;
  

  for( var col = 1; col <= cols; col++ )                     
  {
    var title = ss.getActiveSheet().getRange( 1, col ).getValue();                     
    
    
    if( title == "Due Date" || title == "Unlock Date" || title == "Lock Date" )                                // Parse dates based on titles
    {
      
      for( var row = 2; row <= rows; row++ )                                             
      { 
      
        // Active sheet variables       
        var value = ss.getActiveSheet().getRange( row, col ).getValue();
        var assignment = ss.getActiveSheet().getRange( row, 1 ).getValue();
        var weekAssigned = assignment.match( /week\s*\d?\d/i );                                                 // Get the assignment name and pop off prepended week #  
        
       
       /* Assignment has preprended week # */
        if ( weekAssigned )                        
        {
          var week = parseInt( weekAssigned[0].split( /^week\s*/i )[1] );
          
          colorCodeDateRanges( row, col, week + 1, value );
        }
        
        /* Assignment doesn't have preprended week # */
        else if( ! weekAssigned )                                                                             // if the assignment doesn't have a prepended week #
        {
          var modrows = modules_sheet.getDataRange().getNumRows();
          var counter = 0;                                                                                    // A counter for seeing if the assignment was in a module                            
          
          for( var modrow = 2; modrow < modrows; modrow++ )                                                   // Loop through module items
          {
            if( modules_sheet.getRange( "E" + modrow ).getValue().indexOf( assignment ) > 0 )                 // Check if that assignment is in a module
            {
              counter++;
              
              /* Check date using prepended week # from module */
              try
              {
                var module = modules_sheet.getRange( "A" + modrow ).getValue().match(/week\s*\d?\d/i);         // Get the module name and pop off prepended week #
                var modweek = parseInt( module[0].split( /^week\s*/i )[1] );
              }
              catch( error )
              {         
                /* Check date by assuming any digit found is the week # */
                var module = modules_sheet.getRange( "A" + modrow ).getValue().match(/\d?\d/i);              // See if a digit is present
                var modweek = parseInt( module );
                
                continue;
              }            
            }
            
          }//end-for: module search
          
          if( counter > 0 &&  !isNaN( modweek ) )                                                  // If module has prepended week or includes some digit
            colorCodeDateRanges( row, col, modweek + 1, value );      
          else                                                                                     // If none of the above occurs guess: NOT FULLY TESTED might not need
          {            
            //weeks from the dates sheet
            var weeks = dates_sheet.getRange( 2, 1, dates_sheet.getDataRange().getNumRows(), 1 ).getValues();     
        
            // Arbitrarily big initialzation for date search
            var searchdiff = 7 * 24 * 60 * 60 * 1000;                                                             // Initialize to 7 days in milliseconds
            var count = 0;                                                                                        // Count violations
             
            for( var datesrow = 2; datesrow < weeks.length + 2; datesrow++ )                                      // For each week 
            { 
              try                                                                                                               // Hack: add offset of 2 hours
              {
                var startDate = new Date( dates_sheet.getRange( datesrow, 2 ).getValue().getTime() + 2 * 60 * 60 * 1000 ),      // Start date from 'Dates' sheet   
                    endDate = new Date( dates_sheet.getRange( datesrow, 3 ).getValue().getTime() + 2 * 60 * 60 * 1000 ),        // End date from 'Dates' sheet
                    date = new Date( value.getTime() + 2 * 60 * 60 * 1000 );                                                    // Date value under title
              }
              catch( error ){
                break;
              }
              
              
              if( title == "Due Date" || title == "Lock Date" )
              {
                tmp = searchdiff;                                                               // Hold old searchdiff
                searchdiff = endDate.getTime() - date.getTime();                                // store new searchdiff
                
                if( searchdiff <= 0 && searchdiff >= tmp )                                      // Grab the closest date before our date variable
                  var mostLikelyCoordinates = [ row, col, datesrow, value ];
                if( searchdiff > 0 )                                                            // Out of range
                  count++;
              }           
             
              if( title == "Unlock Date" )
              {
                tmp = searchdiff;
                searchdiff = startDate.getTime() - date.getTime();
                
                if( searchdiff >= 0 && ( searchdiff + tmp ) <= ( 2 * searchdiff ) )             // Grab the closest date that falls after our date variable
                  var mostLikelyCoordinates = [ row, col, datesrow, value ];
                if( searchdiff < 0 )                                                            // Out of range
                  count++;
              }
              
            }//end-for: week search
            
             
            if( count == 8 || count == weeks.length )                                           // Completely out of range      
              sheet.getRange( row, col ).setBackground("#f4cccc");
              
            if( mostLikelyCoordinates )
            {
              colorCodeDateRanges( mostLikelyCoordinates[0], mostLikelyCoordinates[1], mostLikelyCoordinates[2], mostLikelyCoordinates[3] );  // Pass coordinates to be colored 
              mostLikelyCoordinates = null;                                                                                                   // Reset variable
            }
            
            
          }//end-else: Guess
            
        }//end-else if: Assignment has no prepended week #
        
      }//end-for: row parse  
      
    } 
    
  }//end-for: column parse
  
  
  if( unprovidedDates > 0 )
    SpreadsheetApp.getUi().alert( "Please provide additional dates on 'Dates' sheet. Item week exceeds weeks provided." );
    
  
  /***************************************************************
   * Function: colorCodeDateRanges
   * Purpose: Color code date on active sheet depending on if it falls into or out of the range on the 'Dates' sheet.
   *****************************************************************/  
  
  function colorCodeDateRanges( row, col, datesRow, value )
  {
    try
    {
      var startDate = new Date( dates_sheet.getRange( datesRow, 2 ).getValue().getTime() + 2 * 60 * 60 * 1000 ),         // Hack: add offset of 2 hours
          endDate = new Date( dates_sheet.getRange( datesRow, 3 ).getValue().getTime() + 2 * 60 * 60 * 1000 );
    }
    catch( error )
    {
      unprovidedDates++;
      return;
    }
    
    // Various differences to determine date comparison
    var absolutediff = endDate.getTime() - startDate.getTime(),
        maxdiff = endDate.getDate() - startDate.getDate(),
        anywherediff,
        daydiff,
        monthdiff,
        yeardiff;
    
    try{
      var date = new Date( value.getTime() + 2 * 60 * 60 * 1000 );
    }     
    catch(error){         
      return;
    }
    
    if( title == "Due Date" || title == "Lock Date" )
    {
      anywherediff = endDate.getTime() - date.getTime();      //should be positive
      daydiff = endDate.getDate() - date.getDate();           
      monthdiff = endDate.getMonth() - date.getMonth();
      yeardiff = endDate.getFullYear() - date.getFullYear();
    }           
    
    if( title == "Unlock Date" )
    {
      anywherediff = startDate.getTime() - date.getTime();    //should be positive
      daydiff = startDate.getDate() - date.getDate();         
      monthdiff = startDate.getMonth() - date.getMonth();
      yeardiff = startDate.getFullYear() - date.getFullYear();
    }
    
    /* Color Code Date Ranges */ 
// Same Month
    if( daydiff > 0 && daydiff < maxdiff && monthdiff == 0 && yeardiff == 0 )                                                           
      sheet.getRange( row, col ).setBackground("#ffe599");                                                       //ok: yellow -- #ffe599    
    else if( title == "Unlock Date" && daydiff > 0 && monthdiff == 0 && yeardiff == 0 )   
      sheet.getRange( row, col ).setBackground("#ffe599");  
    else if( daydiff < 0 && monthdiff == 0 && yeardiff == 0 )                             
      sheet.getRange( row, col ).setBackground("#f4cccc");                                                       //bad: ligth red -- #f4cccc
    else if( daydiff == 0 && monthdiff == 0 && yeardiff == 0 )                            
      sheet.getRange( row, col ).setBackground("#00ffff");                                                       //good: cyan -- #00ffff
// Different Months
    else if( Math.abs( monthdiff ) > 0 && anywherediff > 0 && anywherediff < absolutediff  && yeardiff == 0 )                            
      sheet.getRange( row, col ).setBackground("#ffe599");                                                       //ok: yellow -- #ffe599 
    else if( Math.abs( monthdiff ) > 0 && title == "Unlock Date" && anywherediff > 0 && yeardiff == 0 )                            
      sheet.getRange( row, col ).setBackground("#ffe599");                                                       
    else if( Math.abs( monthdiff ) > 0 && anywherediff < 0 && yeardiff == 0 )                            
      sheet.getRange( row, col ).setBackground("#f4cccc");                                                       //bad: ligth red -- #f4cccc
    else
      sheet.getRange( row, col ).setBackground("#f4cccc");
      
  }//end-subfunction
  
  
}