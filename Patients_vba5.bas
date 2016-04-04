Private Sub Form_BeforeUpdate(Cancel As Integer)
If Me.NewRecord Then
recordcreatedby = "whoever"
Me.Individual_Created_Date = Now()
Else
recordlasteditedby = "whoever"
Me.Individual_Last_Modified = Now() Or Date
End If
End Sub

'Start at new record on open
Private Sub Form_Open(Cancel As Integer)
DoCmd.GoToRecord , , acNewRec
End Sub

'If in more than one trial box is checked, make Trial 2 boxes visible
Private Sub Check185_AfterUpdate()
If Check185 = 0 Then
Me.Trial2.Visible = False
Me.Trial_ID2.Visible = False
Me.Label366.Visible = False
Me.Hospital2.Visible = False
Me.Hospital_Trial_ID2.Visible = False
Me.Label385.Visible = False
Me.Label142.Visible = False
Me.Label325.Visible = False


ElseIf Check185 = -1 Then
Me.Trial2.Visible = True
Me.Trial_ID2.Visible = True
Me.Label366.Visible = True
Me.Hospital2.Visible = True
Me.Hospital_Trial_ID2.Visible = True
Me.Label385.Visible = True
Me.Label142.Visible = True
Me.Label325.Visible = True

End If
End Sub

Private Sub DOB_AfterUpdate()
' Check to see if the Initials and DOB are already in the database (after DOB update)
    Dim Init As String
    Dim DateOfBirth As Date
    Dim Initials As String
    Dim DOB As Date
    Dim rsc As DAO.Recordset

    Set rsc = Me.RecordsetClone

    Init = Nz(Me.Initials.Value)
    DateOfBirth = Nz(Me.DOB.Value)
    rsc.FindFirst ("[Initials]= '" & UCase(Init) & "' And [DOB]= #" & Format(DateOfBirth, "dd-mmm-yyyy") & "#")
    'Check Patients table for duplicate Initials and DOB
    If Not rsc.NoMatch Then
        'Undo duplicate entry
        Me.Undo
        'Message box warning of duplication
        MsgBox "Warning: Initials " & Init & " and Date of Birth " & DateOfBirth & " is already in the database." & vbNewLine & "This duplicate entry has been deleted and you will now be" & vbNewLine & "taken to the original record."
        'Go to record of original Initials and DOB entry
        Me.Bookmark = rsc.Bookmark
    End If
    
    Set rsc = Nothing
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
'Shoot out error if NHS number isn't 10 digits
Const INPUTMASK_VIOLATION = 2279
   If DataErr = INPUTMASK_VIOLATION Then
      MsgBox "An NHS number must be 10 digits!"
      Response = acDataErrContinue
   Else
    'Display a standard error message
    Response = acdatadisplay
    End If
End Sub


Private Sub Hospital_AfterUpdate()
' Set SQL working to current database
Dim db As Database
Set db = CurrentDb()
Me.Refresh

'Set Hospital_ID to null if hospital is not chosen
If IsNull(Hospital) = True Then
'Delete previous entry for this patient in the Hosp table
CurrentDb.Execute "UPDATE tblPATIENTS SET [Hospital_ID1] = NULL WHERE tblPATIENTS.PatientID = " & Me.PatientID
CurrentDb.Execute "UPDATE tblPATIENTS SET [Hospital_Trial_ID1] = NULL WHERE tblPATIENTS.PatientID = " & Me.PatientID
    Exit Sub

Else
    CurrentDb.Execute "UPDATE tblPATIENTS INNER JOIN tblREF_Hospitals ON tblPATIENTS.Hospital = tblREF_Hospitals.HospitalID SET tblPATIENTS.Hospital_ID1 = tblREF_Hospitals.[Centre ID] WHERE tblPATIENTS.PatientID = " & Me.PatientID
    CurrentDb.Execute "UPDATE tblPATIENTS SET [Hospital_Trial_ID1] = '" & Me.[Hospital_ID1] & "_" & Me.Trial_ID1 & "' WHERE tblPATIENTS.PatientID = " & Me.PatientID
End If
End Sub

Private Sub Hospital2_AfterUpdate()
' Set SQL working to current database
Dim db As Database
Set db = CurrentDb()
Me.Refresh

'Set Hospital_ID2 to null if hospital is not chosen
If IsNull(Hospital2) = True Then
'Delete previous entry for this patient in the Hosp table
CurrentDb.Execute "UPDATE tblPATIENTS SET [Hospital_ID2] = NULL WHERE tblPATIENTS.PatientID = " & Me.PatientID
CurrentDb.Execute "UPDATE tblPATIENTS SET [Hospital_Trial_ID2] = NULL WHERE tblPATIENTS.PatientID = " & Me.PatientID
    Exit Sub

Else
    CurrentDb.Execute "UPDATE tblPATIENTS INNER JOIN tblREF_Hospitals ON tblPATIENTS.Hospital2 = tblREF_Hospitals.HospitalID SET tblPATIENTS.Hospital_ID2 = tblREF_Hospitals.[Centre ID] WHERE tblPATIENTS.PatientID = " & Me.PatientID
    CurrentDb.Execute "UPDATE tblPATIENTS SET [Hospital_Trial_ID2] = '" & Me.[Hospital_ID2] & "_" & Me.Trial_ID2 & "' WHERE tblPATIENTS.PatientID =" & Me.PatientID
End If
End Sub

Private Sub Initials_AfterUpdate()
' Check to see if the Initials and DOB are already in the database (after initials update)
    Dim Init As String
    Dim DateOfBirth As Date
    Dim Initials As String
    Dim DOB As Date
    Dim rsc As DAO.Recordset

    Set rsc = Me.RecordsetClone

    Init = Nz(Me.Initials.Value)
    DateOfBirth = Nz(Me.DOB.Value)
    rsc.FindFirst ("[Initials]= '" & Init & "' And [DOB]= #" & Format(DateOfBirth, "dd-mmm-yyyy") & "#")
    'Check Patients table for duplicate Initials and DOB
    If Not rsc.NoMatch Then
        'Undo duplicate entry
        Me.Undo
        'Message box warning of duplication
        MsgBox "Warning: Initials " & UCase(Init) & " and Date of Birth " & DateOfBirth & " is already in the database." & vbNewLine & "This duplicate entry has been deleted and you will now be" & vbNewLine & "taken to the original record."
        'Go to record of original Initials and DOB entry
        Me.Bookmark = rsc.Bookmark
    End If
    
    Set rsc = Nothing
End Sub

Private Sub NHS_number_AfterUpdate()
    ' Check to see if the NHS number is already in the database
    Dim NHSno As String
    Dim NHS_number As String
    Dim rsc As DAO.Recordset

    Set rsc = Me.RecordsetClone

    NHSno = Me.NHS_number.Value
    stLinkCriteria = "[NHS_number]= '" & NHSno & "'"

    'Check Patients table for duplicate NHS number
    If DCount("[NHS_number]", "tblPATIENTS", stLinkCriteria) > 0 Then
        'Undo duplicate entry
        Me.Undo
        'Message box warning of duplication
        MsgBox "Warning: NHS number " & NHSno & " is already in the database." & vbNewLine & "This duplicate entry has been deleted and you will now be" & vbNewLine & "taken to the original record."
        'Go to record of original NHS number entry
        rsc.FindFirst stLinkCriteria
        Me.Bookmark = rsc.Bookmark
    End If
    
    Set rsc = Nothing
End Sub
