
Private Sub CD138_Cells_Frozen_AfterUpdate()
' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
' Refresh the TBMA entry form
Me.Requery

' Define CD138 CellsName as the current fields calculated in tblSample_BMA table
Dim CD138CellsName As String
    CD138CellsName = Me.Lab_numberID & "_BMA_138_CELLS"
    
'Delete any previous entries for container and position if cells frozen is set to null
If IsNull(CD138_Cells_Frozen) = True Then
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
'Delete all previous entries for this lab number in the CellsLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_Cells WHERE Cells_Sample = '" & CD138CellsName & "'"

'Ensure all container and position boxes are invisible
Me.[CD138 Cells Container1].Visible = False
Me.[CD138 Cells Position1].Visible = False
Me.Label1453.Visible = False
Me.[CD138 Cells Container2].Visible = False
Me.[CD138 Cells Position2].Visible = False
Me.Label1458.Visible = False
Me.[CD138 Cells Container3].Visible = False
Me.[CD138 Cells Position3].Visible = False
Me.Label1460.Visible = False
Me.[CD138 Cells Container4].Visible = False
Me.[CD138 Cells Position4].Visible = False
Me.Label1462.Visible = False
    Exit Sub
End If

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
    
'CD138 CELLS 1
If CD138_Cells_Frozen = 1 Or CD138_Cells_Frozen = 2 Or CD138_Cells_Frozen = 3 Or CD138_Cells_Frozen = 4 Then
'Delete all previous entries for this lab number in the BMA Sample table
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Name] = '" & CD138CellsName & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container1] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position1] = '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CD138CellsName & "')"

'Make container and position visible for one tube
Me.[CD138 Cells Container1].Visible = True
Me.[CD138 Cells Position1].Visible = True
Me.Label1453.Visible = True
End If

'CD138 CELLS 2
If CD138_Cells_Frozen = 2 Or CD138_Cells_Frozen = 3 Or CD138_Cells_Frozen = 4 Then
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
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container2] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position2] =  '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the second sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CD138CellsName & "')"

Me.[CD138 Cells Container2].Visible = True
Me.[CD138 Cells Position2].Visible = True
Me.Label1458.Visible = True
End If

'CD138 CELLS 3
If CD138_Cells_Frozen = 3 Or CD138_Cells_Frozen = 4 Then
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

'Add in container entry for third tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container3] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position3] =  '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the third sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CD138CellsName & "')"

Me.[CD138 Cells Container3].Visible = True
Me.[CD138 Cells Position3].Visible = True
Me.Label1460.Visible = True
End If

'CD138 CELLS 4
If CD138_Cells_Frozen = 4 Then
'Add one to position for fourth tube, including moving to new box if necessary
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

'Add in container entry for fourth tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Container4] = '" & LastContainerCells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 Cells Position4] =  '" & NewPos81Cells & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the third sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_Cells (Cells_Container, Cells_Position, Cells_Sample) VALUES ('" & LastContainerCells & "', '" & NewPos81Cells & "', '" & CD138CellsName & "')"

Me.[CD138 Cells Container4].Visible = True
Me.[CD138 Cells Position4].Visible = True
Me.Label1462.Visible = True
End If

End Sub


Private Sub TCP_Frozen_AfterUpdate()
' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
' Refresh the Treatment entry form
Me.Requery

'TCP
' Define TCPName as the current fields calculated in tblSample_BMA table
Dim TCPName As String
    TCPName = Me.Lab_numberID & "_BMA_TCP"
    
'Delete any previous entries for container and position if TCP frozen is set to null
If IsNull(TCP_Frozen) = True Then
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
'Delete all previous entries for this lab number in the TCPLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_TCP WHERE TCP_Sample = '" & TCPName & "'"
    Exit Sub
End If
    
'Delete all previous entries for this lab number in the TCPLoc table
CurrentDb.Execute "DELETE * FROM tblLoc_TCP WHERE TCP_Sample = '" & TCPName & "'"

'Ensure all container and position boxes are invisible
Me.[TCP Container1].Visible = False
Me.[TCP Position1].Visible = False
Me.Label1453.Visible = False
Me.[TCP Container2].Visible = False
Me.[TCP Position2].Visible = False
Me.Label1458.Visible = False
Me.[TCP Container3].Visible = False
Me.[TCP Position3].Visible = False
Me.Label1460.Visible = False
Me.[TCP Container4].Visible = False
Me.[TCP Position4].Visible = False
Me.Label1462.Visible = False

' Work out the last entry in the TCP Location table
Dim LastContainerTCP As Integer
        LastContainerTCP = DLast("TCP_Container", "tblLoc_TCP")
Dim LastPositionTCP As String
    LastPositionTCP = DLast("TCP_Position", "tblLoc_TCP")
Dim LastPosIDTCP As Integer
    LastPosIDTCP = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionTCP & "'")
Dim NewPosIDTCP As Integer
    NewPosIDTCP = LastPosIDTCP + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDTCP = 82 Then
        LastContainerTCP = LastContainerTCP + 1
        NewPosIDTCP = 1
        Else
        LastContainerTCP = LastContainerTCP
        NewPosIDTCP = NewPosIDTCP
    End If
    
Dim NewPos81TCP As String
    NewPos81TCP = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDTCP)

    
'TCP 1
If TCP_Frozen = 1 Or TCP_Frozen = 2 Or TCP_Frozen = 3 Or TCP_Frozen = 4 Then
'Delete all previous entries for this lab number from the BMA sample table
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position1] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position2] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position3] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position4] = NULL WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Name] = '" & TCPName & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container1] = '" & LastContainerTCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position1] = '" & NewPos81WCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO tblLoc_TCP (TCP_Container, TCP_Position, TCP_Sample) VALUES ('" & LastContainerTCP & "', '" & NewPos81TCP & "', '" & TCPName & "')"


'Make container and position visible for one tube
Me.[TCP Container1].Visible = True
Me.[TCP Position1].Visible = True
Me.Label1453.Visible = True
End If

'TCP 2
If TCP_Frozen = 2 Or TCP_Frozen = 3 Or TCP_Frozen = 4 Then
'Add one to position for second tube, including moving to new box if necessary
NewPosIDTCP = NewPosIDTCP + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDTCP = 82 Then
        LastContainerTCP = LastContainerTCP + 1
        NewPosIDTCP = 1
        Else
        LastContainerTCP = LastContainerTCP
        NewPosIDTCP = NewPosIDTCP
    End If

NewPos81TCP = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDTCP)

'Add in container entry for second tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container2] = '" & LastContainerTCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position2] = '" & NewPos81TCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the second sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_TCP (TCP_Container, TCP_Position, TCP_Sample) VALUES ('" & LastContainerTCP & "', '" & NewPos81TCP & "', '" & TCPName & "')"

Me.[TCP Container2].Visible = True
Me.[TCP Position2].Visible = True
Me.Label1458.Visible = True
End If

'TCP 3
If TCP_Frozen = 3 Or TCP_Frozen = 4 Then
'Add one to position for second tube, including moving to new box if necessary
NewPosIDTCP = NewPosIDTCP + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDTCP = 82 Then
        LastContainerTCP = LastContainerTCP + 1
        NewPosIDTCP = 1
        Else
        LastContainerTCP = LastContainerTCP
        NewPosIDTCP = NewPosIDTCP
    End If

NewPos81TCP = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDTCP)

'Add in container entry for third tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container3] = '" & LastContainerTCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position3] = '" & NewPos81TCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the third sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_TCP (TCP_Container, TCP_Position, TCP_Sample) VALUES ('" & LastContainerTCP & "', '" & NewPos81TCP & "', '" & TCPName & "')"

Me.[TCP Container3].Visible = True
Me.[TCP Position3].Visible = True
Me.Label1460.Visible = True
End If

'TCP 4
If TCP_Frozen = 4 Then
'Add one to position for fourth tube, including moving to new box if necessary
NewPosIDTCP = NewPosIDTCP + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDTCP = 82 Then
        LastContainerTCP = LastContainerTCP + 1
        NewPosIDTCP = 1
        Else
        LastContainerTCP = LastContainerTCP
        NewPosIDTCP = NewPosIDTCP
    End If

NewPos81TCP = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDTCP)

'Add in container entry for fourth tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Container4] = '" & LastContainerTCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [TCP Position4] = '" & NewPos81TCP & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
' Add the third sample into the location table
CurrentDb.Execute "INSERT INTO tblLoc_TCP (TCP_Container, TCP_Position, TCP_Sample) VALUES ('" & LastContainerTCP & "', '" & NewPos81TCP & "', '" & TCPName & "')"

Me.[TCP Container4].Visible = True
Me.[TCP Position4].Visible = True
Me.Label1462.Visible = True
End If

End Sub


'FISH
Private Sub FISH_Conc__x_10_6__AfterUpdate()

' Define FISHName as the current fields calculated in tblSample_BMA table
Dim FISHName As String
    FISHName = Me.Lab_numberID & "_BMA_FISH"
        
'Delete all previous entries for this lab number in the TCPLoc table
CurrentDb.Execute "DELETE * FROM [tblLoc_CD138FISH] WHERE FISH_Sample = '" & FISHName & "'"

' Work out the last entry in the FISH Location table
Dim LastContainerFISH As Integer
    LastContainerFISH = DLast("FISH_Container", "tblLoc_CD138FISH")
Dim LastPositionFISH As String
    LastPositionFISH = DLast("FISH_Position", "tblLoc_CD138FISH")
Dim LastPosIDFISH As Integer
    LastPosIDFISH = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionFISH & "'")
Dim NewPosIDFISH As Integer
    NewPosIDFISH = LastPosIDFISH + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDFISH = 82 Then
        LastContainerFISH = LastContainerFISH + 1
        NewPosIDFISH = 1
        Else
        LastContainerFISH = LastContainerFISH
        NewPosIDFISH = NewPosIDFISH
    End If
    
Dim NewPos81FISH As String
    NewPos81FISH = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDFISH)
Me.Requery

If FISH_Conc__x_10_6_ = 0 Or FISH_Conc__x_10_6_ = Null Then
'If box 0 or null delete name, container and position entry for tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 FISH Name] = Null WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 FISH Container] = Null WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 FISH Position] = Null WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

'FISH added in BMA sample
ElseIf FISH_Conc__x_10_6_ <> 0 Then
'Add in name, container and position entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 FISH Name] = '" & FISHName & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 FISH Container] = '" & LastContainerFISH & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 FISH Position] = '" & NewPos81FISH & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"


' Add the FISH samples into their respective location table
CurrentDb.Execute "INSERT INTO [tblLoc_CD138FISH] (FISH_Container, FISH_Position, FISH_Sample) VALUES ('" & LastContainerFISH & "', '" & NewPos81FISH & "', '" & FISHName & "')"
End If
Me.Requery

End Sub

Private Sub RLT_Conc__x_10_6__AfterUpdate()
' Set SQL working to current database
    Dim db As Database
    Set db = CurrentDb()
' Refresh the Treatment entry form
Me.Requery

Me.[CD138 RLT Name].Visible = False
Me.[CD138 RLT Container].Visible = False
Me.[CD138 RLT Position].Visible = False

' Define TCPName as the current fields calculated in tblSample_BMA table
Dim RLTName As String
    RLTName = Me.Lab_numberID & "_BMA_R138"
    
'Delete all previous entries for this lab number in the TCPLoc table
CurrentDb.Execute "DELETE * FROM [tblLoc_CD138RLT] WHERE RLT_Sample = '" & RLTName & "'"

'Delete all previous entries for this lab number in the BMA Sample table
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 RLT Container] = Null WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 RLT Position] = Null WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Work out the last entry in the RLT Location table
Dim LastContainerRLT As Integer
    LastContainerRLT = DLast("RLT_Container", "tblLoc_CD138RLT")
Dim LastPositionRLT As String
    LastPositionRLT = DLast("RLT_Position", "tblLoc_CD138RLT")
Dim LastPosIDRLT As Integer
    LastPosIDRLT = DLookup("PosID", "tblREF_81well", "Pos ='" & LastPositionRLT & "'")
Dim NewPosIDRLT As Integer
    NewPosIDRLT = LastPosIDRLT + 1
' If the last position in the box is filled you'll move onto a new box, so add one to the container
    If NewPosIDRLT = 82 Then
        LastContainerRLT = LastContainerRLT + 1
        NewPosIDRLT = 1
        Else
        LastContainerRLT = LastContainerRLT
        NewPosIDRLT = NewPosIDRLT
    End If
    
Dim NewPos81RLT As String
    NewPos81RLT = DLookup("Pos", "tblREF_81well", "PosID =" & NewPosIDRLT)
Me.Requery
        
If RLT_Conc__x_10_6_ <> 0 And DCount("[RLT_Sample]", "tblLoc_CD138RLT", "RLT_Sample = '" & RLTName & "'") = 0 Then
'Add in container entry for one tube
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 RLT Container] = '" & LastContainerRLT & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"
CurrentDb.Execute "UPDATE tblSample_BMA SET [CD138 RLT Position] = '" & NewPos81RLT & "' WHERE Lab_numberID = '" & Me.Lab_numberID & "'"

' Add the Cells samples into their respective location table
CurrentDb.Execute "INSERT INTO [tblLoc_CD138RLT] (RLT_Container, RLT_Position, RLT_Sample) VALUES ('" & LastContainerRLT & "', '" & NewPos81RLT & "', '" & RLTName & "')"

Me.[CD138 RLT Name].Visible = True
Me.[CD138 RLT Container].Visible = True
Me.[CD138 RLT Position].Visible = True
End If
    
End Sub
