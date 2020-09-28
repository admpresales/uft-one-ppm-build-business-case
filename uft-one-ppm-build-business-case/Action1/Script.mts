'===========================================================
'Function to Create a Random Number with DateTime Stamp
'===========================================================
Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

'======================== End Function =====================
End Function

Dim BrowserExecutable, Counter

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon
Set AppContext2=Browser("CreationTime:=1")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Executive Overview link
'===========================================================================================
AIUtil.FindText("Strategic Portfolio").Click

'===========================================================================================
'BP:  Click the Barbara Getty (Business Relationship Manager) link to log in as Barabara Getty
'===========================================================================================
AIUtil.FindTextBlock("Barabara Getty").Click
AIUtil.FindTextBlock("New Proposals").Exist

'===========================================================================================
'BP:  Click the Search menu item
'===========================================================================================
AIUtil.FindText("SEARCH", micFromTop, 1).Click

'===========================================================================================
'BP:  Click the Requests text
'===========================================================================================
AIUtil.FindTextBlock("Requests", micFromTop, 1).Click

'===========================================================================================
'BP:  Enter PFM - Proposal into the Request Type field
'===========================================================================================
AIUtil("text_box", "Request Type:").Type "PFM - Proposal"
AIUtil("text_box", "Assigned To").Click

'===========================================================================================
'BP:  Enter a status of "New" into the Status field
'===========================================================================================
AIUtil("text_box", "Status").Type "New"

'===========================================================================================
'BP:  Click the Search button (OCR not seeing text, use traditional OR)
'===========================================================================================
Browser("Search Requests").Page("Search Requests").Link("Search").Click

'===========================================================================================
'BP:  Click the first record returned in the search results
'===========================================================================================
DataTable.Value("dtFirstReqID") = Browser("Search Requests").Page("Request Search Results").WebTable("Req #").GetCellData(2,2)
AIUtil.FindTextBlock(DataTable.Value("dtFirstReqID")).Click

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
AIUtil.FindText("Approved").Click

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
AIUtil.FindText("Continue WorkflowAction").Click
AIUtil("text_box", "*Region:").Exist

'===========================================================================================
'BP:  Enter "US" into the Region field
'===========================================================================================
AIUtil("text_box", "*Region:").Type "US"

'===========================================================================================
'BP:  Click the Completed button
'===========================================================================================
AIUtil("button", "Completed").Click
AIUtil.FindTextBlock("Project Class").Exist

'===========================================================================================
'BP:  Select "Innovation" in the Project Class
'===========================================================================================
AIUtil("combobox", "Project Class").Select "Innovation"

'===========================================================================================
'BP:  Select "Infrastructure" in the Asset Class
'===========================================================================================
AIUtil("combobox", ":Asset Class").Select "Infrastructure"

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
AIUtil.FindText("Continue WorkflowAction").Click

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
AIUtil.FindText("Approved", micFromLeft, 1).Click

'===========================================================================================
'BP:  Enter the Expected Start Period as June 2021
'===========================================================================================
AIUtil("text_box", "Expected Stan Period").Type "June 2021"

'===========================================================================================
'BP:  Enter the Expected Finish Period as December 2021
'===========================================================================================
AIUtil("text_box", "Expected Finish Period").Type "December 2021"

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
AIUtil.FindText("Continue WorkflowAction").Click

'===========================================================================================
'BP:  Click the Completed button
'===========================================================================================
AIUtil("button", "Completed").Click

'===========================================================================================
'BP:  Click the Create button
'===========================================================================================
Browser("Search Requests").Page("Req #42957: More Information").WebElement("Create").Click
AppContext2.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext2.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext2																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Create button in the opopup window
'===========================================================================================
Browser("Create a Blank Staffing").Page("Create a Blank Staffing").WebButton("button.create").Click

'===========================================================================================
'BP:  Click the Select the Staffing Profile button
'===========================================================================================
AIUtil("button", "Select the Staffing Profile").Click

'===========================================================================================
'BP:  Enter "A/R Billing Upgrade" into the Staffing Profile field
'===========================================================================================
AIUtil("text_box", "Staffing Profile:").Type "A/R Billing Upgrade"
AIUtil.FindText("Staffing Profile:", micFromBottom, 1).Click

'===========================================================================================
'BP:  Click the Import button
'===========================================================================================
'Counter = 0
'Do
'	Counter = Counter + 1
'	wait(1)
	Browser("Create a Blank Staffing").Page("Staffing Profile").Frame("copyPositionsDialogIF").Link("Import").Click
'	If Counter >= 60 Then
'		msgbox("Something is wrong")
'	End If
'Loop While AIUtil("text_box", "Staffing Profile:").Exist

'===========================================================================================
'BP:  Click the Done text
'===========================================================================================
Do
	AIUtil.FindText("Done").Click
Loop While AIUtil.FindText("Done").Exist
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
AIUtil.FindText("Continue WorkflowAction").Click
AIUtil.FindTextBlock("Status: Finance Review").Exist

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req #42953: Details").WebElement("menuUserIcon").Click
AIUtil.FindTextBlock("Sign Out (Barbara Getty)").Click

AppContext.Close																			'Close the application at the end of your script

