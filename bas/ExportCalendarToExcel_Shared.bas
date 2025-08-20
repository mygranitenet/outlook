Attribute VB_Name = "ExportCalendarToExcel_Shared"
' ===================================================================================
' EXPORT OUTLOOK CALENDAR TO EXCEL SCRIPT (with Shared Calendar support)
' Description: This script exports appointments from either a user's own calendar
'              or a shared calendar to a new Excel spreadsheet for a specified
'              date range. It correctly handles recurring meetings.
' ===================================================================================

Sub ExportCalendarToExcel_Shared()
    ' --- PART 1: DECLARE VARIABLES ---
    ' Outlook objects
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.NameSpace
    Dim olRecipient As Outlook.Recipient
    Dim olFolder As Outlook.MAPIFolder
    Dim olItems As Outlook.Items
    Dim olRestrictedItems As Outlook.Items
    Dim olApt As Outlook.AppointmentItem
    
    ' Excel objects (using "late binding")
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    
    ' Other variables
    Dim iRow As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim sFilter As String
    Dim objItem As Object
    Dim choice As VbMsgBoxResult
    Dim strOwnerEmail As String
    
    ' --- PART 2: SETUP AND USER INPUT ---
    On Error GoTo ErrorHandler
    
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    
    ' Ask user if they want to export a shared calendar or their own
    choice = MsgBox("Do you want to export a SHARED calendar?" & vbCrLf & vbCrLf & _
                    "Click 'Yes' for a shared calendar." & vbCrLf & _
                    "Click 'No' for one of your own calendars.", _
                    vbYesNoCancel + vbQuestion, "Select Calendar Type")

    If choice = vbCancel Then
        Exit Sub ' User cancelled
    ElseIf choice = vbYes Then
        ' --- A) Handle SHARED Calendar ---
        strOwnerEmail = InputBox("Enter the full email address of the person whose calendar you want to export:", "Shared Calendar Owner")
        If strOwnerEmail = "" Then Exit Sub ' User cancelled
        
        Set olRecipient = olNs.CreateRecipient(strOwnerEmail)
        olRecipient.Resolve
        
        If olRecipient.Resolved Then
            ' Use GetSharedDefaultFolder to access the calendar
            ' An error will occur here if you don't have permissions
            On Error Resume Next ' Temporarily disable error handling to check for permissions
            Set olFolder = olNs.GetSharedDefaultFolder(olRecipient, olFolderCalendar)
            On Error GoTo ErrorHandler ' Re-enable standard error handling
            
            If olFolder Is Nothing Then
                MsgBox "Could not open the shared calendar. Please ensure:" & vbCrLf & _
                       "1. The email address is correct." & vbCrLf & _
                       "2. You have permissions to view the calendar.", vbCritical
                Exit Sub
            End If
        Else
            MsgBox "Could not find a user with the email address: " & strOwnerEmail, vbCritical
            Exit Sub
        End If
        
    Else ' choice = vbNo
        ' --- B) Handle OWN Calendar ---
        ' Prompt user to select a calendar folder from their own mailbox
        Set olFolder = olNs.PickFolder
        If olFolder Is Nothing Or olFolder.DefaultItemType <> olAppointmentItem Then
            MsgBox "No valid calendar folder was selected. Aborting.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Get date range from user
    strStartDate = InputBox("Enter the Start Date (e.g., YYYY-MM-DD):", "Export Start Date", Format(Date, "yyyy-mm-dd"))
    If Not IsDate(strStartDate) Then
        MsgBox "Invalid start date format. Aborting.", vbCritical
        Exit Sub
    End If
    
    strEndDate = InputBox("Enter the End Date (e.g., YYYY-MM-DD):", "Export End Date", Format(Date + 30, "yyyy-mm-dd"))
    If Not IsDate(strEndDate) Then
        MsgBox "Invalid end date format. Aborting.", vbCritical
        Exit Sub
    End If

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
    
    With xlWs
        .Cells(1, 1).Value = "Subject"
        .Cells(1, 2).Value = "Start Time"
        .Cells(1, 3).Value = "End Time"
        .Cells(1, 4).Value = "Duration (Minutes)"
        .Cells(1, 5).Value = "Location"
        .Cells(1, 6).Value = "Organizer"
        .Cells(1, 7).Value = "Required Attendees"
        .Cells(1, 8).Value = "Optional Attendees"
        .Cells(1, 9).Value = "Categories"
        .Cells(1, 10).Value = "Is Recurring"
        
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.ColorIndex = 15 ' Light Gray
    End With
    
    iRow = 2 ' Start writing data from row 2
    
    For Each objItem In olRestrictedItems
        If TypeName(objItem) = "AppointmentItem" Then
            Set olApt = objItem
            
            With xlWs
                .Cells(iRow, 1).Value = olApt.Subject
                .Cells(iRow, 2).Value = olApt.Start
                .Cells(iRow, 3).Value = olApt.End
                .Cells(iRow, 4).Value = olApt.Duration
                .Cells(iRow, 5).Value = olApt.Location
                .Cells(iRow, 6).Value = olApt.Organizer
                .Cells(iRow, 7).Value = GetAttendees(olApt, 1) ' 1 for Required
                .Cells(iRow, 8).Value = GetAttendees(olApt, 2) ' 2 for Optional
                .Cells(iRow, 9).Value = olApt.Categories
                .Cells(iRow, 10).Value = olApt.IsRecurring
            End With
            
            iRow = iRow + 1
        End If
    Next objItem
    
    xlWs.Columns.AutoFit
    
    MsgBox "Calendar export complete! " & (iRow - 2) & " appointments were exported from '" & olFolder.Name & "'.", vbInformation

    ' --- PART 5: CLEANUP ---
CleanExit:
    Set olApt = Nothing
    Set olRestrictedItems = Nothing
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olRecipient = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


' --- HELPER FUNCTION to get attendees as a string ---
Function GetAttendees(ByVal oAppointment As Outlook.AppointmentItem, ByVal attendeeType As Integer) As String
    Dim oRecipients As Outlook.Recipients
    Dim oRecipient As Outlook.Recipient
    Dim sAttendees As String
    
    Set oRecipients = oAppointment.Recipients
    sAttendees = ""
    
    For Each oRecipient In oRecipients
        If oRecipient.Type = attendeeType Then
            sAttendees = sAttendees & oRecipient.Name & "; "
        End If
    Next
    
    If Len(sAttendees) > 2 Then
        GetAttendees = Left(sAttendees, Len(sAttendees) - 2)
    Else
        GetAttendees = ""
    End If
    
    Set oRecipient = Nothing
    Set oRecipients = Nothing
End Function

