'===========================================================
'20200929 - DJ:  Updated project creation sync loop to break out and fail the report
'				 if it takes more than 90 seconds to create the project.
'20200929 - DJ:  Updated impoper syntax in the Exit Do
'20200929 - DJ: Added .sync statements after .click statements and additional tuning
'20200930 - DJ: Added additional step to click on the Status label before attempting to type in the status field
'20201001 - DJ: Added the ClickLoop and the PPMProposalSearch functions, removed duplicative code
'20201001 - DJ: Changed the Execute Now from button to text, on lower resolution settings button isn't properly identified
'===========================================================


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

Function ClickLoop (AppContext, ClickStatement, SuccessStatement)
	
	Dim Counter
	
	Counter = 0
	Do
		ClickStatement.Click
		AppContext.Sync																				'Wait for the browser to stop spinning
		Counter = Counter + 1
		wait(1)
		If Counter >=90 Then
			msgbox("Something is broken, the Requests hasn't shown up")
			Reporter.ReportEvent micFail, "Click the Search text", "The Requests text didn't display within " & Counter & " attempts."
			Exit Do
		End If
	Loop Until SuccessStatement.Exist(1)
	AppContext.Sync																				'Wait for the browser to stop spinning

End Function

Function PPMProposalSearch (CurrentStatus, NextAction)
	'===========================================================================================
	'BP:  Click the Search menu item
	'===========================================================================================
	Set ClickStatement = AIUtil.FindText("SEARCH", micFromTop, 1)
	Set SuccessStatement = AIUtil.FindTextBlock("Requests", micFromTop, 1)
	ClickLoop AppContext, ClickStatement, SuccessStatement
	AppContext.Sync																				'Wait for the browser to stop spinning
	
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
'BP:  Click the Executive Overview link
'===========================================================================================
AIUtil.FindText("Strategic Portfolio").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Jonathan Kaplan (Portfolio Manager) link to log in as Jonathan Kaplan
'===========================================================================================
AIUtil.FindTextBlock("Jonathan Kaplan").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Approval Queue - Key Attributes").Exist

'===========================================================================================
'BP:  Search for proposals in a status of "Standards Review"
'===========================================================================================
PPMProposalSearch "Finance Review", "Approved"

'===========================================================================================
'BP:  Click the left Approved button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Approved", micFromLeft, 1)
Set SuccessStatement = AIUtil.FindTextBlock("Status: Review Complete")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the remiaining Approved button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Approved")
Set SuccessStatement = AIUtil.FindTextBlock("Status: ITSC Review")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Click the remiaining Approved button
'===========================================================================================
Set ClickStatement = AIUtil.FindText("Approved")
Set SuccessStatement = AIUtil("text_box", "'Project Manager:")
ClickLoop AppContext, ClickStatement, SuccessStatement

'===========================================================================================
'BP:  Set the Project Manager to be Joseph Banks
'===========================================================================================
AIUtil("text_box", "'Project Manager:").Type "Joseph Banks"

'===========================================================================================
'BP:  Enter Standard Project (PPM) - Medium Size into the Projec Type field
'===========================================================================================
AIUtil("text_box", "Project Type:").Type "Standard Project (PPM) - Medium Size"

'===========================================================================================
'BP:  Click the Continue Workflow Action button
'===========================================================================================
AIUtil.FindText("Continue WorkflowAction").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Execute Now button
'===========================================================================================
AIUtil.FindTextBlock("Execute Now").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Wait for the Status:Closed (Approved) to show up
'===========================================================================================
Counter = 0
Do
	Counter = Counter + 1
	wait(1)
	If Counter >=90 Then
		msgbox("Something is broken, status of the request hasn't shown up to be approved.")
		Reporter.ReportEvent micFail, "Create Project", "The project creation didn't finish within " & Counter & " seconds."
		Exit Do
	End If
Loop Until AIUtil.FindTextBlock("Status: Closed (Approved)").Exist(1)
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Search Requests").Page("Req #42953: Details").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.FindTextBlock("Sign Out (Jonathan Kaplan)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

