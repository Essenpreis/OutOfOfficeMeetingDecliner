Sub CancelMeetings()
On Error GoTo eh

   Dim msg As String
    msg = InputBox("What is the reason to cancel the meeting(s)? Leave empty to send default ooo message.")
    If msg = "" Then
        msg = "Cancellation Reason: I am out of office. If we should repeat the meeting, respectively my attendance is required, please propose a new date."
    End If
    
    Dim myaccount As String
    myaccount = Application.Session.CurrentUser

    Dim Session As Outlook.NameSpace
    Dim currentExplorer As Explorer
    Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
    Set Selection = currentExplorer.Selection
    
    Dim app As AppointmentItem
    Dim response As Outlook.MeetingItem
    Dim app2 As AppointmentItem
    Dim mItem As MeetingItem
            
        
    'For all items selected...
    For Each app In Selection
      'app.Display
                  
      
      If app.Organizer = myaccount Then
        'handle meetings I organized
        
        app.ForceUpdateToAllAttendees = True 'careful: when opening meeting of other person this way one could send updates in the name of the author
        
        If app.MeetingStatus = olMeetingCanceled Then
            app.Delete
        ElseIf app.Class = olMeetingRequest Then
            'if we have a meeting request benefit of its cancellation handling
            Set app2 = app.GetAssociatedAppointment(False)
            If Not app2 Is Nothing Then
                Set mItem = app2.Respond(olMeetingCancellation, True, False)
                mItem.Body = msg
                mItem.Send
            End If
        ElseIf app.Recipients.Count = 0 Then
            app.Delete
        ElseIf app.Recipients.Count = 1 And app.Recipients.Item(1) = myaccount Then
            app.Delete
        
        ElseIf app.Class = olAppointment Then
            app.MeetingStatus = olMeetingCanceled
            app.Body = msg
            app.Subject = "[Meeting cancellation] " & app.Subject
            app.BodyFormat = olFormatPlain
            app.Save
            app.Send
        Else
            MsgBox ("Could not identify action for meeting " & app.Subject)
        End If
      
      ElseIf Not app.Organizer = myaccount Then
        'handle meetings I didn't organize
        
        app.ForceUpdateToAllAttendees = False
        
        If app.MeetingStatus = olMeetingReceivedAndCanceled Then
            app.Delete
        ElseIf app.Class = olMeetingRequest Then
            'if we have a meeting request benefit of its cancellation handling
            Set app2 = app.GetAssociatedAppointment(False)
            If Not app2 Is Nothing Then
                Set mItem = app2.Respond(olMeetingDeclined, True, False)
                mItem.Body = msg
                mItem.Send
            End If
            
        ElseIf app.Class = olAppointment And (app.MeetingStatus = olMeetingReceived) Then
        ' And Not (app.ResponseStatus = olResponseTentative Or app.ResponseStatus = olResponseNotResponded) Then
                
            'copied from https://www.slipstick.com/developer/accept-or-decline-a-meeting-request-using-vba/
            Dim cAppt As AppointmentItem
            Dim oResponse
            
            Set cAppt = app 'GetCurrentItem.GetAssociatedAppointment(True)
            Set oResponse = cAppt.Respond(olMeetingDeclined, True)
            
            If cAppt.ResponseRequested = True Then
                oResponse.Send 'TODO: for some invitations vba exits when sending...
            End If
            
            Set cAppt = Nothing
            Set oResponse = Nothing
        
        Else
            MsgBox ("Could not identify action for meeting " & app.Subject)
        End If
        
            
      Else
            MsgBox ("Could not identify action for meeting " & app.Subject)
                 
      End If 'big if
    
        
    Next
    

Exit Sub
    
eh:
    MsgBox ("Error" & Err.Description & " with meeting " & app.Subject)
    Debug.Print "Error number: " & Err.Number _
            & " " & Err.Description & " with meeting " & app.Subject

End Sub
