'===========================================================
'20200929 - DJ Modified the dates for expected start and finish date on the project proposal
'			to calculate the current year and add 1 so that we don't have to update the script
'			when the new year happens, it will always be current year + 1
'20200929 - DJ: Added .sync statements after .click statements and cleaned up some commented code
'20200929 - DJ: Redesigned the AI steps as sometimes the PPM application freezes the browser at the same time that
'				the AI is trying to click
'20200930 - DJ: Modified reporter event error text to be more accurate
'20200930 - DJ: When executing from Jenkins, one AI statement doesn't execute properly the first time, if you
'				execute it a second time (in the loop) it executes fine.  While investigating, replace with a 
'				traditional OR statement.
'20200930 - DJ: Modified the loop condition to utilize the workflow status soas to be more reliable
'20200930 - DJ: Updated setting the Region value to be traditional OR, the validation process of PPM can cause the value
'				not to be accepted by the UI when we do the AI type
'20200930 - DJ: Added loop break for continue workflow action of last step, in case loop timing happened to get to retrying
'				the continue action when PPM finally completed the last attempt
'20200930 - DJ: Updated the click out of the Request Type to just click on the label for the status to ensure that PPM page 
'				reload timing won't sporadically cause the AI to type in the wrong field.
'20201001 - DJ: Made the ClickLoop function, replaced duplicative code throughout with references to ClickLoop, made the
'				proposal search as a function, so as to allow easy sharing with other scripts.
'20201006 - DJ: Changed the comboboxes to be traditional OR, had an issue with someone's odd resolution
'20201006 - DJ: Updated PPMProposalSearch function to look for the text Saved Searches to know Search click worked and commented .sync after that to prevent PPM from auto closing the menu popup
'20201006 - DJ: Removed unused function
'				Added function comments
'20201006 - DJ: Updated Expected next step for clicking completed to use an AI statement instead of a traditional OR statement
'20201008 - DJ: Found another example where the PPM application sometimes locks up, click statement with CLickLoop call
'20201012 - DJ: Found another example where the PPM application sometimes locks up, during a .type command, added logic to try again
'20201013 - DJ: Modified the ClickLoop retry counter to be 3 instead of 90
'===========================================================


'===========================================================
'Function to retry action if next step doesn't show up
'===========================================================
Function ClickLoop (AppContext, ClickStatement, SuccessStatement)
	
	Dim Counter
	
	Counter = 0
	Do
		ClickStatement.Click
		AppContext.Sync																				'Wait for the browser to stop spinning
		Counter = Counter + 1
		wait(1)
		If Counter >=3 Then
			msgbox("Something is broken, the Requests hasn't shown up")
			Reporter.ReportEvent micFail, "Click the Search text", "The Requests text didn't display within " & Counter & " attempts."
			Exit Do
		End If
	Loop Until SuccessStatement.Exist(1)
	AppContext.Sync																				'Wait for the browser to stop spinning

End Function

'===========================================================
'Function to search for the PPM proposal in the appropriate status
'===========================================================
Function PPMProposalSearch (CurrentStatus, NextAction)
	'===========================================================================================
	'BP:  Click the Search menu item
	'===========================================================================================
	Set ClickStatement = AIUtil.FindText("SEARCH", micFromTop, 1)
	Set SuccessStatement = AIUtil.FindText("Saved Searches")
	ClickLoop AppContext, ClickStatement, SuccessStatement
	'AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the Requests text
	'===========================================================================================
	Set ClickStatement = AIUtil.FindTextBlock("Requests", micFromTop, 1)
	Set SuccessStatement = AIUtil("text_box", "Request Type:")
	ClickLoop AppContext, ClickStatement, SuccessStatement
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter PFM - Proposal into the Request Type field
	'===========================================================================================
	AIUtil("text_box", "Request Type:").Type "PFM - Proposal"
	AIUtil.FindText("Status").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Enter a status of "New" into the Status field
	'===========================================================================================
	AIUtil("text_box", "Status").Type CurrentStatus
	
	'===========================================================================================
	'BP:  Click the Search button (OCR not seeing text, use traditional OR)
	'===========================================================================================
	Browser("Search Requests").Page("Search Requests").Link("Search").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	
	'===========================================================================================
	'BP:  Click the first record returned in the search results
	'===========================================================================================
	DataTable.Value("dtFirstReqID") = Browser("Search Requests").Page("Request Search Results").WebTable("Req #").GetCellData(2,2)
	Set ClickStatement = AIUtil.FindTextBlock(DataTable.Value("dtFirstReqID"))
	Set SuccessStatement = AIUtil.FindText(NextAction)
	ClickLoop AppContext, ClickStatement, SuccessStatement
	AppContext.Sync																				'Wait for the browser to stop spinning
	
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
'BP:  Click the Strategic Portfolio link
'===========================================================================================
AIUtil.FindText("Strategic Portfolio").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Barbara Getty (Business Relationship Manager) link to log in as Barabara Getty
'===========================================================================================
AIUtil.FindTextBlock("Barabara Getty").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("New Proposals").Exist
AppContext.Sync																				'Wait for the browser to stop spinning

PPMProposalSearch "New", "Approved"

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Approved")
Set SuccessStatement = Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
'When executing from Jenkins, the AI statement is failing the first time, 2nd time it runs, while 
'	investigating, replace with traditional OR step
'	AIUtil.FindText("Continue WorkflowAction").Click
Set ClickStatement = Browser("Search Requests").Page("Req More Information").WebElement("Continue Workflow Action")
Set SuccessStatement = AIUtil.FindTextBlock("Status: High-Level Business Case")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Enter "US" into the Region field
'===========================================================================================
'AIUtil("text_box", "*Region:").Type "US"
Browser("Search Requests").Page("Req Details").WebEdit("Region").Set "US"

'===========================================================================================
'BP:  Click the Completed button
'===========================================================================================
Set ClickStatement = AIUtil("button", "Completed")
Set SuccessStatement = AIUtil.FindTextBlock("Project Class")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Select "Innovation" in the Project Class
'===========================================================================================
'AIUtil("combobox", "Project Class").Select "Innovation"
Browser("Search Requests").Page("Req More Information").WebList("Project Class").Select "Innovation"

'===========================================================================================
'BP:  Select "Infrastructure" in the Asset Class
'===========================================================================================
'AIUtil("combobox", ":Asset Class").Select "Infrastructure"
Browser("Search Requests").Page("Req More Information").WebList("Asset Class").Select "Infrastructure"

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Continue WorkflowAction")
Set SuccessStatement = AIUtil.FindTextBlock("Status: 1st Level Review")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Approved button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Approved", micFromLeft, 1)
Set SuccessStatement = AIUtil("text_box", "Expected Finish Period")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Enter the Expected Start Period as June 2021
'===========================================================================================
AIUtil("text_box", "Expected Stan Period").Type "June " & (Year(Now)+1)

'===========================================================================================
'BP:  Enter the Expected Finish Period as December 2021
'===========================================================================================
AIUtil("text_box", "Expected Finish Period").Type "December " & (Year(Now)+1)

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Continue WorkflowAction")
Set SuccessStatement = AIUtil.FindTextBlock("Status: Detailed Business Case")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Completed button
'===========================================================================================
Set ClickStatement = AIUtil("button", "Completed")
Set SuccessStatement = AIUtil.FindTextBlock("Staffing Profile")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Create button
'===========================================================================================
Set ClickStatement = Browser("Search Requests").Page("Req More Information").WebElement("Create")
Set SuccessStatement = Browser("Create a Blank Staffing").Page("Create a Blank Staffing").WebButton("button.create")
ClickLoop AppContext, ClickStatement, SuccessStatement
AppContext2.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext2.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext2																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Click the Create button in the popup window
'===========================================================================================
Browser("Create a Blank Staffing").Page("Create a Blank Staffing").WebButton("button.create").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Select the Staffing Profile button
'===========================================================================================
Set ClickStatement = AIUtil("button", "Select the Staffing Profile")
Set SuccessStatement = AIUtil("text_box", "Staffing Profile:")
ClickLoop AppContext2, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Enter "A/R Billing Upgrade" into the Staffing Profile field
'===========================================================================================
Counter = 0
Do
	AIUtil("text_box", "Staffing Profile:").Type "A/R Billing Upgrade"
	AppContext2.Sync																				'Wait for the browser to stop spinning
	Counter = Counter + 1
	wait(1)
	If Counter >=3 Then
		msgbox("Something is broken, the Billing Upgrade text hasn't shown up")
		Reporter.ReportEvent micFail, "Type the Staffing Profile", "The Billing Upgrade text didn't display within " & Counter & " attempts."
		Exit Do
	End If
Loop Until AIUtil.FindText("Billing Upgrade").Exist(1)
'AIUtil("text_box", "Staffing Profile:").Type "A/R Billing Upgrade"
Set ClickStatement = AIUtil.FindText("Staffing Profile:", micFromBottom, 1)
Set SuccessStatement = Browser("Create a Blank Staffing").Page("Staffing Profile").Frame("copyPositionsDialogIF").Link("Import")
ClickLoop AppContext2, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the Import button
'===========================================================================================
Browser("Create a Blank Staffing").Page("Staffing Profile").Frame("copyPositionsDialogIF").Link("Import").Click
AppContext2.Sync																				'Wait for the browser to stop spinning

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
Do
	AppContext.Sync																				'Wait for the browser to stop spinning
	If AIUtil.FindTextBlock("Status: Finance Review").Exist(0) Then
		Exit Do
	End If
	AIUtil.FindText("Continue WorkflowAction").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	Counter = Counter + 1
	wait(1)
	If Counter >=90 Then
		msgbox("Something is broken, the project health override window hasn't opened.")
		Reporter.ReportEvent micFail, "Click the Continue WorkflowAction button", "The Status: Finance Review didn't display within " & Counter & " attempts."
		Exit Do
	End If
Loop Until AIUtil.FindTextBlock("Status: Finance Review").Exist(5)

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req Details").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Sign Out (Barbara Getty)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

