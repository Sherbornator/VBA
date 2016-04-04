Option Compare Database
'MUK7
Private Sub Check1257_AfterUpdate()
    Dim db As Database
    Set db = CurrentDb()
    Dim Trial1 As String
        Trial1 = Nz(DLookup("Trial1", "tblPATIENTS", "PatientID =" & Me.Patient_number))
    Dim Trial2 As String
        Trial2 = Nz(DLookup("Trial2", "tblPATIENTS", "PatientID =" & Me.Patient_number))


If Check1257 = -1 And Not (Trial1 = "MUK 7" Or Trial2 = "MUK 7") Then
    MsgBox "MUK 7 not entered for this patient, are you sure?"
    Me.Check1257.SetFocus
End If
    
If Check1257 = -1 And (Trial1 = "MUK 7" Or Trial2 = "MUK 7") Then
    Me.fsubMUK7.Requery
    Me.fsubMUK7.Visible = True
End If

If Check1257 = -1 And DCount("[Lab_numberID]", "tblSample_BMA", "Lab_numberID = '" & Me.Lab_numberID & "'") = 0 Then
    CurrentDb.Execute "INSERT INTO tblSample_BMA (PatientID, Lab_numberID, Sample_type, Sample_type_short, [CD138 RLT Name], [Waste Name]) VALUES ('" & Me.Patient_number & "', '" & Me.Lab_numberID & "', 'Bone Marrow Aspirate', 'BMA', '" & Me.Lab_numberID & "' & '_BMA_CD138_RLT', '" & Me.Lab_numberID & "' & '_BMA_WASTE')"
End If

' Define WasteName as the current fields calculated in tblSample_BMA table
Dim WasteName As String
    WasteName = Me.Lab_numberID & "_BMA_W138"
   
'Delete all previous entries for this lab number in the CellsLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & WasteName & "'"

'CELLS
' Work out the last entry in the Cells Location table
Dim LastContainerCells As Integer
    LastContainerCells = DLast("Cells_Container", "tblLoc_Cells")
Dim LastPositionCells As String
    LastPositionCells = DLast("Cells_Position", "tblLoc_Cells")
Dim LastPosIDCells As Integer
    LastPosIDCells = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionCells & "'")
Dim NewPosIDCells As Integer
    NewPosIDCells = LastPosIDCells + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDCells = 82 Then
        LastContainerCells = LastContainerCells + 1
        NewPosIDCells = 1
        Else
        LastContainerCells = LastContainerCells
        NewPosIDCells = NewPosIDCells
    End If
    
Dim NewPos81Cells As String
    NewPos81Cells = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDCells)
[fsubMUK7].Requery
    
'Waste cells added in Cells box
If Check1257 = -1 And DCount("[Cells_Sample]", "tblLoc_Cells", "Cells_Sample = '" & WasteName & "'") = 0 Then
'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [Waste DMSO Container] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [Waste DMSO Position] = '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & WasteName & "')"

End If

'DELETE ENTIRE SAMPLE
'If MUK7 is unchecked, hide the subform and delete entries in the BMA table with this Lab number AND entries in ALL the location tables
If Check1257 = 0 Then
    'Undo duplicate entry
        Dim intResponse As Integer
        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String
        Dim strSQL4 As String
        Dim strSQL5 As String
        Dim strSQL6 As String
        Dim strSQL7 As String
        Dim RLTName As String
            RLTName = Me.Lab_numberID & "_BMA_CD138_RLT"
        Dim FISHName As String
            FISHName = Me.Lab_numberID & "_BMA_FISH"
        Dim TCPName As String
            TCPName = Me.Lab_numberID & "_BMA_TCP"
        Dim SerumName As String
            SerumName = Me.Lab_numberID & "_PB_SER"
        Dim WCPName As String
            WCPName = Me.Lab_numberID & "_PB_WCP"
        intResponse = MsgBox("Warning: Did you mean to delete this entire MUK 7 entry?", vbYesNo + vbQuestion + vbDefaultButton2)
            If intResponse = vbYes Then
            strSQL1 = "DELETE * FROM tblSample_BMA WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
            strSQL2 = "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & WasteName & "'"
            strSQL3 = "DELETE * FROM [tblLoc_CD138RLTMUK7] WHERE RLT_Sample = '" & RLTName & "'"
            strSQL4 = "DELETE * FROM [tblLoc_CD138FISHMUK7] WHERE FISH_Sample = '" & FISHName & "'"
            strSQL5 = "DELETE * FROM [tblLoc_TCP] WHERE TCP_Sample = '" & TCPName & "'"
            strSQL6 = "DELETE * FROM [tblLoc_SerumMUK7] WHERE Serum_Sample = '" & SerumName & "'"
            strSQL7 = "DELETE * FROM [tblLoc_WCPMUK7] WHERE WCP_Sample = '" & WCPName & "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL1
            DoCmd.RunSQL strSQL2
            DoCmd.RunSQL strSQL3
            DoCmd.RunSQL strSQL4
            DoCmd.RunSQL strSQL5
            DoCmd.RunSQL strSQL6
            DoCmd.RunSQL strSQL7
            DoCmd.SetWarnings True
            MsgBox "The entire MUK7 sample has been deleted."
            Me.fsubMUK7.Visible = False
            Me.Requery
        Else
            Check1257 = -1
        Me.fsubMUK7.Visible = True
        End If
End If

End Sub

'DATE SAMPLE RECEIVED
Private Sub Date_sample_received_BeforeUpdate(Cancel As Integer)

'Warn that the End Date is before the Begin Date
If IsNull(Me.Date_sample_taken) Then
    Exit Sub
ElseIf Me.Date_sample_received <= Me.Date_sample_taken Then
    MsgBox "Date sample received must be after sample taken."
End If
End Sub

'BONE MARROW ASPIRATE
Private Sub Bone_Marrow_Aspirate_check_AfterUpdate()
'Check to see if Lab number ID has been filled in
If IsNull(Me.[Lab#]) Then
    MsgBox "Please enter lab number first.", vbOKOnly
    Bone_Marrow_Aspirate_check = False
    Me.[Lab#].SetFocus
    Exit Sub
Else:

'If Bone Marrow Aspirate is checked, make the subform visible and add entries in the BMA table
' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
    
If Bone_Marrow_Aspirate_check = -1 And DCount("[Lab_numberID]", "tblSample_BMA", "Lab_numberID = '" & Me.Lab_numberID & "'") = 0 Then
    CurrentDb.Execute "INSERT INTO tblSample_BMA (PatientID, Lab_numberID, Sample_type, Sample_type_short, [CD138 RLT Name], [Waste Name]) VALUES ('" & Me.Patient_number & "', '" & Me.Lab_numberID & "', 'Bone Marrow Aspirate', 'BMA', '" & Me.Lab_numberID & "' & '_BMA_CD138_RLT', '" & Me.Lab_numberID & "' & '_BMA_WASTE')"
    Me.fsubBMA.Requery
    Me.fsubBMA.Visible = True
End If

' Define WasteName as the current fields calculated in tblSample_BMA table
Dim WasteName As String
    WasteName = Me.Lab_numberID & "_BMA_W138"
   
'Delete all previous entries for this lab number in the CellsLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & WasteName & "'"

'CELLS
' Work out the last entry in the Cells Location table
Dim LastContainerCells As Integer
    LastContainerCells = DLast("Cells_Container", "tblLoc_Cells")
Dim LastPositionCells As String
    LastPositionCells = DLast("Cells_Position", "tblLoc_Cells")
Dim LastPosIDCells As Integer
    LastPosIDCells = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionCells & "'")
Dim NewPosIDCells As Integer
    NewPosIDCells = LastPosIDCells + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDCells = 82 Then
        LastContainerCells = LastContainerCells + 1
        NewPosIDCells = 1
        Else
        LastContainerCells = LastContainerCells
        NewPosIDCells = NewPosIDCells
    End If
    
Dim NewPos81Cells As String
    NewPos81Cells = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDCells)
[fsubBMA].Requery
    
'Waste cells added in Cells box
If Bone_Marrow_Aspirate_check = -1 And DCount("[Cells_Sample]", "tblLoc_Cells", "Cells_Sample = '" & WasteName & "'") = 0 Then
'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [Waste DMSO Container] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [Waste DMSO Position] = '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & WasteName & "')"

'WASTE in -20
' Define WasteName as the current fields calculated in tblSample_BMA table
Dim WasteCellsName As String
    WasteCellsName = Me.Lab_numberID & "_BMA_W138_Cells"

'Delete all previous entries for this lab number in the WasteLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_Waste WHERE Waste_Sample = '" & WasteCellsName & "'"
    
' Work out the last entry in the Waste Location table NB. box of 100!
Dim LastContainerWaste As Integer
    LastContainerWaste = DLast("Waste_Container", "tblLoc_Waste")
Dim LastPositionWaste As Integer
    LastPositionWaste = DLast("Waste_Position", "tblLoc_Waste")
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If LastPositionWaste = 100 Then
        LastContainerWaste = LastContainerWaste + 1
        LastPositionWaste = 0
        Else: LastContainerWaste = LastContainerWaste
    End If
[fsubBMA].Requery

'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [Waste Cells Container] = '" & LastContainerWaste & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [Waste Cells Position] = '" & LastPositionWaste & "' +1 WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO tblLoc_Waste (Waste_Container, Waste_Position, Waste_Sample) VALUES ('" & LastContainerWaste & "', '" & LastPositionWaste & "' +1, '" & WasteCellsName & "')"
End If
[fsubBMA].Requery

'DELETE ENTIRE SAMPLE
'If Bone Marrow Aspirate is unchecked, hide the subform and delete entries in the BMA table with this Lab number AND entries in ALL the location tables
If Bone_Marrow_Aspirate_check = 0 Then
    'Undo duplicate entry
        Dim intResponse As Integer
        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String
        Dim strSQL4 As String
        Dim strSQL5 As String
        Dim strSQL6 As String
        Dim RLTName As String
            RLTName = Me.Lab_numberID & "_BMA_CD138_RLT"
        Dim FISHName As String
            FISHName = Me.Lab_numberID & "_BMA_FISH"
        Dim TCPName As String
            TCPName = Me.Lab_numberID & "_BMA_TCP"
            
        intResponse = MsgBox("Warning: Did you mean to delete this entire Bone Marrow Aspirate entry?", vbYesNo + vbQuestion + vbDefaultButton2)
            If intResponse = vbYes Then
            strSQL1 = "DELETE * FROM tblSample_BMA WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
            strSQL2 = "DELETE * FROM tblLoc_Waste WHERE Waste_Sample = '" & WasteCellsName & "'"
            strSQL3 = "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & WasteName & "'"
            strSQL4 = "DELETE * FROM [tblLoc_CD138RLT] WHERE RLT_Sample = '" & RLTName & "'"
            strSQL5 = "DELETE * FROM [tblLoc_CD138FISH] WHERE FISH_Sample = '" & FISHName & "'"
            strSQL6 = "DELETE * FROM [tblLoc_TCP] WHERE TCP_Sample = '" & TCPName & "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL1
            DoCmd.RunSQL strSQL2
            DoCmd.RunSQL strSQL3
            DoCmd.RunSQL strSQL4
            DoCmd.RunSQL strSQL5
            DoCmd.RunSQL strSQL6
            DoCmd.SetWarnings True
            MsgBox "The entire BMA sample has been deleted."
            Me.fsubBMA.Visible = False
            Me.Requery
        Else
            Bone_Marrow_Aspirate_check = -1
        Me.fsubBMA.Visible = True
        End If
End If
End If

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
If Me.NewRecord Then
recordcreatedby = "whoever"
Me.Sample_Created_Date = Now()
Else
recordlasteditedby = "whoever"
Me.Sample_Last_Modified = Now()
End If
End Sub

Private Sub Form_Current()

If Me.NewRecord Then
    Me.fsubTreatment.Visible = False
    Me.fsubBMA.Visible = False
    Me.fsubMUK7.Visible = False
    Me.fsubPB__.Visible = False
    Me.Peripheral_blood_check = False
    Me.Treatment_check = False
    Me.Bone_Marrow_Aspirate_check = False
    Me.Bone_Marrow_Smears_check = False
Else:

Dim Trial1 As String
    Trial1 = Nz(DLookup("Trial1", "tblPATIENTS", "PatientID =" & Me.Patient_number))
Dim Trial2 As String
    Trial2 = Nz(DLookup("Trial2", "tblPATIENTS", "PatientID =" & Me.Patient_number))
        
If Not IsNull(Me.[Lab#]) = True And DCount("[Lab_numberID]", "tblSample_PB", "Lab_numberID = '" & Me.[Lab#] & "'") <> 0 Then
    Me.fsubPB__.Visible = True
    Me.Peripheral_blood_check = True
    Else
    Me.fsubPB__.Visible = False
    Me.Peripheral_blood_check = False
End If

If Not IsNull(Me.[Lab#]) = True And (DCount("[Lab_numberID]", "tblSample_BMA", "Lab_numberID = '" & Me.[Lab#] & "'") <> 0 And (Not (Trial1 = "MUK 7" Or Trial2 = "MUK 7"))) Then
    Me.fsubBMA.Visible = True
    Me.Bone_Marrow_Aspirate_check = True
    Else
    Me.fsubBMA.Visible = False
    Me.Bone_Marrow_Aspirate_check = False
End If

If Not IsNull(Me.[Lab#]) = True And (DCount("[Lab_numberID]", "tblSample_BMA", "Lab_numberID = '" & Me.[Lab#] & "'") <> 0 And (Trial1 = "MUK 7" Or Trial2 = "MUK 7")) Then
    Me.fsubMUK7.Visible = True
    Me.Check1257 = True
    Else
    Me.fsubMUK7.Visible = False
    Me.Check1257 = False
End If

If Not IsNull(Me.[Lab#]) = True And DCount("[Lab_numberID]", "tblSample_Treatment", "Lab_numberID = '" & Me.[Lab#] & "'") <> 0 Then
    Me.fsubTreatment.Visible = True
    Me.Treatment_check = True
    Else
    Me.fsubTreatment.Visible = False
    Me.Treatment_check = False
End If
End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
'If an error occurs because of missing Lab ID
Const conErrRequiredData = 3058
     
If DataErr = conErrRequiredData Then
    MsgBox "You must enter a Lab Number to save this record"
    Me.[Lab#].SetFocus
    Response = acDataErrContinue
    Else
    'Display a standard error message
    Response = acdatadisplay
    End If
End Sub
'BONE MARROW SMEARS
Private Sub Bone_Marrow_Smears_check_AfterUpdate()
'Check to see if Lab number ID has been filled in
If IsNull(Me.[Lab#]) Then
    MsgBox "Please enter lab number first.", vbOKOnly
    Bone_Marrow_Smears_check = False
    Me.[Lab#].SetFocus
    Exit Sub
Else:

' Set SQL working to current database
Dim db As Database
    Set db = CurrentDb()
' If BMS check box is selected and there isn't already a record, add record in tblSample_BMS with the current Patient ID and Lab number
If Bone_Marrow_Smears_check = -1 And DCount("[Lab_numberID]", "tblSample_BMS", "Lab_numberID = '" & Me.Lab_numberID & "'") = 0 Then
    CurrentDb.Execute "INSERT INTO tblSample_BMS (PatientID, Lab_numberID, Sample_type, Sample_type_short) VALUES ('" & Me.Patient_number & "', '" & Me.Lab_numberID & "', 'Bone Marrow Smear', 'BMS')"
End If

' If BMS is left empty, delete the record in tblSample_BMS with that patient ID
If Bone_Marrow_Smears_check = 0 Then
    'Undo duplicate entry
        Dim intResponse As Integer
        Dim strSQL As String
        intResponse = MsgBox("Warning: Did you mean to delete this entire Bone Marrow Smear entry?", vbYesNo + vbQuestion + vbDefaultButton2)
            If intResponse = vbYes Then
            strSQL = "DELETE * FROM tblSample_BMS WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            MsgBox "The entire BMS sample has been deleted."
            Me.Requery
        Else
            Bone_Marrow_Smears_check = -1
        End If
End If
End If
End Sub


'PERIPHERAL BLOOD
Private Sub Peripheral_blood_check_AfterUpdate()
'Check to see if Lab number ID has been filled in
If IsNull(Me.[Lab#]) Then
    MsgBox "Please enter lab number first.", vbOKOnly
    Peripheral_blood_check = False
    Me.[Lab#].SetFocus
    Exit Sub
Else:

' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
' Refresh the fsubPB form
    [fsubPB__].Requery
' Show fsubPB if check box is selected
If Peripheral_blood_check = -1 Then
    Me.fsubPB__.Visible = True
End If

' Work out the last entry in the Serum Location table
Dim LastContainerSerum As Integer
        LastContainerSerum = DLast("Serum_Container", "tblLoc_Serum")
Dim LastPositionSerum As String
    LastPositionSerum = DLast("Serum_Position", "tblLoc_Serum")
Dim LastPosIDSerum As Integer
    LastPosIDSerum = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionSerum & "'")
Dim NewPosIDSerum As Integer
    NewPosIDSerum = LastPosIDSerum + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDSerum = 82 Then
        LastContainerSerum = LastContainerSerum + 1
        NewPosIDSerum = 1
        Else
        LastContainerSerum = LastContainerSerum
        NewPosIDSerum = NewPosIDSerum
    End If
    
Dim NewPos81Serum As String
    NewPos81Serum = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDSerum)
[fsubBMA].Requery

' Work out the last entry in the WCP Location table
Dim LastContainerWCP As Integer
        LastContainerWCP = DLast("WCP_Container", "tblLoc_WCP")
 Dim LastPositionWCP As String
    LastPositionWCP = DLast("WCP_Position", "tblLoc_WCP")
Dim LastPosIDWCP As Integer
    LastPosIDWCP = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionWCP & "'")
Dim NewPosIDWCP As Integer
    NewPosIDWCP = LastPosIDWCP + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDWCP = 82 Then
        LastContainerWCP = LastContainerWCP + 1
        NewPosIDWCP = 1
        Else
        LastContainerWCP = LastContainerWCP
        NewPosIDWCP = NewPosIDWCP
    End If
    
Dim NewPos81WCP As String
    NewPos81WCP = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDWCP)
[fsubPB__].Requery

' Define SerumName and WBCName as the current fields calculated in tblSample_PB table
Dim SerumName As String
    SerumName = Me.Lab_numberID & "_PB_SER"
Dim WCPName As String
    WCPName = Me.Lab_numberID & "_PB_WCP"
            
' If PB check box is selected and there isn't already a record, add record in tblSample_PB with the current Patient ID and Lab number
If Peripheral_blood_check = -1 And DCount("[Lab_numberID]", "tblSample_PB", "Lab_numberID = '" & Me.Lab_numberID & "'") = 0 Then
    CurrentDb.Execute "INSERT INTO tblSample_PB (PatientID, Lab_numberID, Sample_type, Sample_type_short, Serum_Name, Serum_Container, Serum_Position, WCP_Name, WCP_Container, WCP_Position) VALUES ('" & Me.Patient_number & "', '" & Me.Lab_numberID & "', 'Peripheral Blood', 'PB', '" & SerumName & "', '" & LastContainerSerum & "', '" & NewPos81Serum & "', '" & WCPName & "', '" & LastContainerWCP & "', '" & NewPos81WCP & "')"
    ' Add the serum and WCP samples into their respective location table
    CurrentDb.Execute "INSERT INTO tblLoc_Serum (Serum_Container, Serum_Position, Serum_Sample) VALUES ('" & LastContainerSerum & "', '" & NewPos81Serum & "', '" & SerumName & "')"
    CurrentDb.Execute "INSERT INTO tblLoc_WCP (WCP_Container, WCP_Position, WCP_Sample) VALUES ('" & LastContainerWCP & "', '" & NewPos81WCP & "', '" & WCPName & "')"
    [fsubPB__].Requery
End If

' If PB is left empty, delete the record in tblSample_PB with that patient ID as long as the serum container field is empty
If Peripheral_blood_check = 0 Then
    'Undo duplicate entry
        Dim intResponse As Integer
        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String
        intResponse = MsgBox("Warning: Did you mean to delete this entire Peripheral Blood entry?", vbYesNo + vbQuestion + vbDefaultButton2)
            If intResponse = vbYes Then
            strSQL1 = "DELETE * FROM tblSample_PB WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
            strSQL2 = "DELETE * FROM tblLoc_Serum WHERE Serum_Sample = '" & SerumName & "'"
            strSQL3 = "DELETE * FROM tblLoc_WCP WHERE WCP_Sample = '" & WCPName & "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL1
            DoCmd.RunSQL strSQL2
            DoCmd.RunSQL strSQL3
            DoCmd.SetWarnings True
            MsgBox "The entire Peripheral Blood sample has been deleted."
            Me.fsubPB__.Visible = False
            Me.Requery
        Else
            Peripheral_blood_check = -1
        Me.fsubPB__.Visible = True
        End If
End If
End If

End Sub


Private Sub Timepoint_AfterUpdate()
Me.Timepoint_short = Timepoint.Column(2)
End Sub

'TREATMENT
'Make subform visible
Private Sub Treatment_check_AfterUpdate()
'Check to see if Lab number ID has been filled in
If IsNull(Me.[Lab#]) Then
    MsgBox "Please enter lab number first.", vbOKOnly
    Treatment_check = False
    Me.[Lab#].SetFocus
    Exit Sub
Else:

' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()

If Treatment_check = -1 Then
    Me.fsubTreatment.Visible = True
End If
'Add entry with current patient ID and lab number
If Treatment_check = -1 And DCount("[Lab_numberID]", "tblSample_Treatment", "Lab_numberID = '" & Me.Lab_numberID & "'") = 0 Then
    CurrentDb.Execute "INSERT INTO tblSample_Treatment (PatientID, Lab_numberID) VALUES ('" & Me.Patient_number & "', '" & Me.Lab_numberID & "')"
    fsubTreatment.Requery
End If
    
If Treatment_check = 0 Then
    'Undo duplicate entry
        Dim intResponse As Integer
        Dim strSQL1 As String
        Dim strSQL2 As String
        intResponse = MsgBox("Warning: Did you mean to delete this entire entry?", vbYesNo + vbQuestion + vbDefaultButton2)
        If intResponse = vbYes Then
            strSQL1 = "DELETE * FROM tblSample_Treatment WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
            strSQL2 = "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & CellsName & "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL1
            DoCmd.RunSQL strSQL2
            DoCmd.SetWarnings True
            MsgBox "The entire Treatment sample has been deleted."
            Me.fsubTreatment.Visible = False
            Me.Requery
        Else
            Treatment_check = -1
        Me.fsubTreatment.Visible = True
        End If
End If
End If
End Sub
