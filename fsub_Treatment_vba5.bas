Private Sub Form_Current()
If Me.NewRecord Then
    Me.Label15.Visible = False
    Me.Label18.Visible = False
    Me.Label177.Visible = False
    Me.Cells_Conc__x_10_6_.Visible = False
    Me.Cells_Frozen.Visible = False
    Me.Cells_Name.Visible = False
End If

If Not IsNull(Me.Lab_numberID) = True And DCount("[Lab_numberID]", "tblSample_Treatment", "Lab_numberID = '" & Me.Lab_numberID & "'") <> 0 Then
    Me.Label15.Visible = True
    Me.Label18.Visible = True
    Me.Label177.Visible = True
    Me.Cells_Conc__x_10_6_.Visible = True
    Me.Cells_Frozen.Visible = True
    Me.Cells_Name.Visible = True
    Else
    Me.Label15.Visible = False
    Me.Label18.Visible = False
    Me.Label177.Visible = False
    Me.Cells_Conc__x_10_6_.Visible = False
    Me.Cells_Frozen.Visible = False
    Me.Cells_Name.Visible = False
End If

End Sub

Private Sub Sample_type_AfterUpdate()
' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
    Me.Requery

If IsNull(Sample_type) = True Then
    Me.Label15.Visible = False
    Me.Label18.Visible = False
    Me.Label177.Visible = False
    Me.Cells_Conc__x_10_6_.Visible = False
    Me.Cells_Frozen.Visible = False
    Me.Cells_Name.Visible = False
Else
    Me.Label15.Visible = True
    Me.Label18.Visible = True
    Me.Label177.Visible = True
    Me.Cells_Conc__x_10_6_.Visible = True
    Me.Cells_Frozen.Visible = True
    Me.Cells_Name.Visible = True
End If

'Add entry with current patient ID and lab number
If Sample_type = "Stem Cell Harvest" Then
    CurrentDb.Execute "UPDATE tblSample_Treatment SET Sample_type_short = 'SCH' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
Else
    CurrentDb.Execute "UPDATE tblSample_Treatment SET Sample_type_short = 'T' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
End If


End Sub

Private Sub Cells_Frozen_AfterUpdate()
' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
    Me.Requery
    
' Define CellsName as the current fields calculated in tblSample_Treatment table
Dim CellsName As String
    CellsName = Me.Lab_numberID & "_" & Me.Sample_type_short & "_BMA_UnselectedCells"
    
'Delete any previous entries for container and position if cells frozen is set to null
If IsNull(Cells_Frozen) = True Then
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
'Delete all previous entries for this lab number in the CellsLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & CellsName & "'"
    Exit Sub
End If
    
'Delete all previous entries for this lab number in the CellsLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & CellsName & "'"

'Ensure all container and position boxes are invisible
Me.Cells_Container1.Visible = False
Me.Cells_Position1.Visible = False
Me.Label382.Visible = False
Me.Cells_Container2.Visible = False
Me.Cells_Position2.Visible = False
Me.Label385.Visible = False
Me.Cells_Container3.Visible = False
Me.Cells_Position3.Visible = False
Me.Label387.Visible = False

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
Me.Requery

'CELLS FROZEN 1
If Cells_Frozen = 1 Or Cells_Frozen = 2 Or Cells_Frozen = 3 Then
'Delete all previous entries for this lab number in the BMA Sample table
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_Treatment SET tblSample_Treatment.CellsName = '" & CellsName & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container1] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position1] = '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

'Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CellsName & "')"

'Make container and position visible for one tube
Me.Cells_Container1.Visible = True
Me.Cells_Position1.Visible = True
Me.Label382.Visible = True
End If

'CELLS FROZEN 2
If Cells_Frozen = 2 Or Cells_Frozen = 3 Then

'Add one to position for second tube, including moving to new box if necessary
NewPosIDCells = NewPosIDCells + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDCells = 82 Then
        LastContainerCells = LastContainerCells + 1
        NewPosIDCells = 1
        Else
        LastContainerCells = LastContainerCells
        NewPosIDCells = NewPosIDCells
    End If

NewPos81Cells = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDCells)

'Add in container entry for second tube
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container2] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position2] = '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the second sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CellsName & "')"

'Make container and position visible for second tube
Me.Cells_Container2.Visible = True
Me.Cells_Position2.Visible = True
Me.Label385.Visible = True
End If

'CELLS FROZEN 3
If Cells_Frozen = 3 Then

'Add one to position for third tube, including moving to new box if necessary
NewPosIDCells = NewPosIDCells + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDCells = 82 Then
        LastContainerCells = LastContainerCells + 1
        NewPosIDCells = 1
        Else
        LastContainerCells = LastContainerCells
        NewPosIDCells = NewPosIDCells
    End If

NewPos81Cells = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDCells)

'Add in container entry for third tube
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Container3] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_Treatment SET [Cells Position3] = '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the second sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CellsName & "')"

Me.Cells_Container3.Visible = True
Me.Cells_Position3.Visible = True
Me.Label387.Visible = True
End If

End Sub

