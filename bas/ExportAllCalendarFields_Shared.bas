' ===================================================================================
' COMPREHENSIVE CALENDAR EXPORT SCRIPT ("Export All Fields")
' Description: This script exports an extensive list of properties for each
'              appointment in a selected calendar to Excel. It handles shared
'              calendars, converts technical codes to readable text, and provides
'              a detailed recurrence summary.
'
' How to Customize: To add or remove a field, simply edit the 'properties'
'                   array below and add a corresponding 'Case' statement in the
'                   main loop to handle the new property.
' ===================================================================================

Sub ExportAllCalendarFields_Shared()
    ' --- PART 1: DECLARE VARIABLES ---
    Dim olApp As Outlook.Application, olNs As Outlook.NameSpace, olRecipient As Outlook.Recipient
    Dim olFolder As Outlook.MAPIFolder, olItems As Outlook.Items, olRestrictedItems As Outlook.Items
    Dim objItem As Object, olApt As Outlook.AppointmentItem
    
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    
    Dim iRow As Long, iCol As Long
    Dim strStartDate As String, strEndDate As String, sFilter As String
    Dim choice As VbMsgBoxResult, strOwnerEmail As String
    
    ' --- DEFINE ALL PROPERTIES TO EXPORT IN THIS ARRAY ---
    Dim properties As Variant
    properties = Array( _
        "Subject", "Start Time", "End Time", "Duration (Mins)", "All Day Event", "Location", _
        "Organizer", "Required Attendees", "Optional Attendees", "Resource Attendees", "Response Status", _
        "Body", "Categories", "Importance", "Sensitivity", "Busy Status", _
        "Is Recurring", "Recurrence Pattern", _
        "Creation Time", "Last Modified Time", "Entry ID" _
    )
    
    ' --- PART 2: SETUP AND USER INPUT ---
    On Error GoTo ErrorHandler
    
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    
    choice = MsgBox("Do you want to export a SHARED calendar?", vbYesNoCancel + vbQuestion, "Select Calendar Type")
    If choice = vbCancel Then Exit Sub
    
    If choice = vbYes Then ' Shared Calendar
        strOwnerEmail = InputBox("Enter the email address of the shared calendar owner:", "Shared Calendar Owner", "ManagedServicesCal@granitenet.com")
        If strOwnerEmail = "" Then Exit Sub
        Set olRecipient = olNs.CreateRecipient(strOwnerEmail)
        olRecipient.Resolve
        If Not olRecipient.Resolved Then
            MsgBox "Could not find a user with the email address: " & strOwnerEmail, vbCritical
            Exit Sub
        End If
        On Error Resume Next
        Set olFolder = olNs.GetSharedDefaultFolder(olRecipient, olFolderCalendar)
        On Error GoTo ErrorHandler
        If olFolder Is Nothing Then
            MsgBox "Could not open the shared calendar. Check permissions and the email address.", vbCritical
            Exit Sub
        End If
    Else ' Own Calendar
        Set olFolder = olNs.PickFolder
        If olFolder Is Nothing Or olFolder.DefaultItemType <> olAppointmentItem Then
            MsgBox "No valid calendar folder was selected. Aborting.", vbExclamation
            Exit Sub
        End If
    End If
    
    strStartDate = InputBox("Enter Start Date (YYYY-MM-DD):", "Export Start Date", Format(Date, "yyyy-mm-dd"))
    If Not IsDate(strStartDate) Then Exit Sub
    strEndDate = InputBox("Enter End Date (YYYY-MM-DD):", "Export End Date", Format(Date + 30, "yyyy-mm-dd"))
    If Not IsDate(strEndDate) Then Exit Sub

    ' --- PART 3: PREPARE CALENDAR ITEMS ---
    Set olItems = olFolder.Items
    olItems.Sort "[Start]"
    olItems.IncludeRecurrences = True
    sFilter = "[Start] >= '" & Format(strStartDate, "ddddd hh:nn AMPM") & "' AND [End] <= '" & Format(strEndDate & " 11:59 PM", "ddddd hh:nn AMPM") & "'"
    Set olRestrictedItems = olItems.Restrict(sFilter)
    
    ' --- PART 4: CREATE EXCEL AND WRITE DATA ---
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)
    
    ' Write headers dynamically from the properties array
    For iCol = 0 To UBound(properties)
        xlWs.Cells(1, iCol + 1).Value = properties(iCol)
    Next iCol
    xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, UBound(properties) + 1)).Font.Bold = True
    
    iRow = 2 ' Start data from row 2
    
    For Each objItem In olRestrictedItems
        If TypeName(objItem) = "AppointmentItem" Then
            Set olApt = objItem
            
            ' Loop through the properties array and populate the values for each
            For iCol = 0 To UBound(properties)
                Dim propValue As Variant
                propValue = "" ' Default to blank
                
                ' Use a Select Case statement to get the value for each property
                Select Case properties(iCol)
                    Case "Subject": propValue = olApt.Subject
                    Case "Start Time": propValue = olApt.Start
                    Case "End Time": propValue = olApt.End
                    Case "Duration (Mins)": propValue = olApt.Duration
                    Case "All Day Event": propValue = olApt.AllDayEvent
                    Case "Location": propValue = olApt.Location
                    Case "Organizer": propValue = olApt.Organizer
                    Case "Required Attendees": propValue = GetAttendees(olApt, 1)
                    Case "Optional Attendees": propValue = GetAttendees(olApt, 2)
                    Case "Resource Attendees": propValue = GetAttendees(olApt, 3)
                    Case "Response Status": propValue = GetResponseStatusText(olApt.ResponseStatus)
                    Case "Body": propValue = olApt.Body
                    Case "Categories": propValue = olApt.Categories
                    Case "Importance": propValue = GetImportanceText(olApt.Importance)
                    Case "Sensitivity": propValue = GetSensitivityText(olApt.Sensitivity)
                    Case "Busy Status": propValue = GetBusyStatusText(olApt.BusyStatus)
                    Case "Is Recurring": propValue = olApt.IsRecurring
                    Case "Recurrence Pattern":
                        If olApt.IsRecurring Then
                            propValue = GetRecurrenceSummary(olApt.GetRecurrencePattern)
                        Else
                            propValue = "Not recurring"
                        End If
                    Case "Creation Time": propValue = olApt.CreationTime
                    Case "Last Modified Time": propValue = olApt.LastModificationTime
                    Case "Entry ID": propValue = olApt.EntryID
                End Select
                
                ' Write the value to the cell
                xlWs.Cells(iRow, iCol + 1).Value = propValue
            Next iCol
            
            iRow = iRow + 1
        End If
    Next objItem
    
    xlWs.Columns.AutoFit
    MsgBox "Comprehensive export complete! " & (iRow - 2) & " appointments exported.", vbInformation

CleanExit:
    Set olApt = Nothing: Set olRestrictedItems = Nothing: Set olItems = Nothing: Set olFolder = Nothing
    Set olRecipient = Nothing: Set olNs = Nothing: Set olApp = Nothing
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' --- HELPER FUNCTIONS TO CONVERT CODES TO TEXT ---

Function GetAttendees(ByVal oApt As Outlook.AppointmentItem, ByVal attendeeType As Integer) As String
    Dim oRecipients As Outlook.Recipients, oRecipient As Outlook.Recipient, sAttendees As String
    Set oRecipients = oApt.Recipients
    For Each oRecipient In oRecipients
        If oRecipient.Type = attendeeType Then sAttendees = sAttendees & oRecipient.Name & "; "
    Next
    If Len(sAttendees) > 2 Then GetAttendees = Left(sAttendees, Len(sAttendees) - 2)
End Function

Function GetImportanceText(ByVal importanceCode As Long) As String
    Select Case importanceCode
        Case 0: GetImportanceText = "Low"
        Case 1: GetImportanceText = "Normal"
        Case 2: GetImportanceText = "High"
        Case Else: GetImportanceText = "Unknown"
    End Select
End Function

Function GetSensitivityText(ByVal sensitivityCode As Long) As String
    Select Case sensitivityCode
        Case 0: GetSensitivityText = "Normal"
        Case 1: GetSensitivityText = "Personal"
        Case 2: GetSensitivityText = "Private"
        Case 3: GetSensitivityText = "Confidential"
        Case Else: GetSensitivityText = "Unknown"
    End Select
End Function

Function GetBusyStatusText(ByVal busyCode As Long) As String
    Select Case busyCode
        Case 0: GetBusyStatusText = "Free"
        Case 1: GetBusyStatusText = "Tentative"
        Case 2: GetBusyStatusText = "Busy"
        Case 3: GetBusyStatusText = "Out of Office"
        Case 4: GetBusyStatusText = "Working Elsewhere"
        Case Else: GetBusyStatusText = "Unknown"
    End Select
End Function

Function GetResponseStatusText(ByVal responseCode As Long) As String
    Select Case responseCode
        Case 0: GetResponseStatusText = "None"
        Case 1: GetResponseStatusText = "Organized"
        Case 2: GetResponseStatusText = "Tentative"
        Case 3: GetResponseStatusText = "Accepted"
        Case 4: GetResponseStatusText = "Declined"
        Case 5: GetResponseStatusText = "Not Responded"
        Case Else: GetResponseStatusText = "Unknown"
    End Select
End Function

Function GetRecurrenceSummary(ByVal oPattern As Outlook.RecurrencePattern) As String
    Dim summary As String
    If oPattern Is Nothing Then GetRecurrenceSummary = "": Exit Function
    
    Select Case oPattern.RecurrenceType
        Case 0 ' olRecursDaily
            summary = "Daily"
            If oPattern.Interval > 1 Then summary = "Every " & oPattern.Interval & " days"
        Case 1 ' olRecursWeekly
            summary = "Weekly on " & GetDayOfWeek(oPattern.DayOfWeekMask)
            If oPattern.Interval > 1 Then summary = "Every " & oPattern.Interval & " weeks on " & GetDayOfWeek(oPattern.DayOfWeekMask)
        Case 2 ' olRecursMonthly
            summary = "Monthly on day " & oPattern.DayOfMonth
            If oPattern.Interval > 1 Then summary = "Every " & oPattern.Interval & " months on day " & oPattern.DayOfMonth
        Case 3 ' olRecursMonthNth
            summary = "Monthly on the " & GetNth(oPattern.instance) & " " & GetDayOfWeek(oPattern.DayOfWeekMask)
        Case 5 ' olRecursYearly
            summary = "Yearly on " & MonthName(oPattern.MonthOfYear) & " " & oPattern.DayOfMonth
        Case 6 ' olRecursYearNth
            summary = "Yearly on the " & GetNth(oPattern.instance) & " " & GetDayOfWeek(oPattern.DayOfWeekMask) & " of " & MonthName(oPattern.MonthOfYear)
    End Select
    GetRecurrenceSummary = summary
End Function

Private Function GetDayOfWeek(mask As Integer) As String
    Dim days As String
    If mask And 1 Then days = days & "Sunday, "
    If mask And 2 Then days = days & "Monday, "
    If mask And 4 Then days = days & "Tuesday, "
    If mask And 8 Then days = days & "Wednesday, "
    If mask And 16 Then days = days & "Thursday, "
    If mask And 32 Then days = days & "Friday, "
    If mask And 64 Then days = days & "Saturday, "
    If Len(days) > 2 Then GetDayOfWeek = Left(days, Len(days) - 2)
End Function

Private Function GetNth(instance As Integer) As String
    Select Case instance
        Case 1: GetNth = "first"
        Case 2: GetNth = "second"
        Case 3: GetNth = "third"
        Case 4: GetNth = "fourth"
        Case 5: GetNth = "last"
    End Select
End Function

