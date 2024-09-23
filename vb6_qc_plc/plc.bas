Option Explicit

Public bY0(0 To 15) As Boolean
Public bY20(20 To 35) As Boolean
Public bX0(0 To 15) As Boolean
Public bX20(20 To 35) As Boolean
Public strPlcX0 As String
Public strPlcX20 As String
Public strPlcY0 As String
Public strPlcY20 As String
Public PLCdata(19) As Double
Public MCstate As Boolean          'Machine On=True
Public iPLC1_OnCommCount As Integer
Public bNext As Boolean
Public bShowPLC As Boolean

Dim iNext As Integer


Public Sub PlcYState()
    If bNext = True Then
       iNext = iNext + 1
    
       If iNext < 3 Then
          Exit Sub
       End If
    End If
    
    iNext = 0
    bNext = False
    
    Dim strPlcX0Temp As String
    Dim strPlcX20Temp As String
    Dim strPlcY0Temp As String
    Dim strPlcY20Temp As String
    Dim i, j, k As Integer
    
    For j = 0 To 15
        strPlcX0Temp = Mid(strPlcX0, j + 1, 1) & strPlcX0Temp
        strPlcX20Temp = Mid(strPlcX20, j + 1, 1) & strPlcX20Temp
        strPlcY0Temp = Mid(strPlcY0, j + 1, 1) & strPlcY0Temp
        strPlcY20Temp = Mid(strPlcY20, j + 1, 1) & strPlcY20Temp
    Next
    
    For i = 0 To 15
        bY0(i) = Mid(strPlcY0Temp, i + 1, 1)
        bX0(i) = Mid(strPlcX0Temp, i + 1, 1)
        bX20(i + 20) = Mid(strPlcX20Temp, i + 1, 1)
        bY20(i + 20) = Mid(strPlcY20Temp, i + 1, 1)
    Next
    
    If bShowPLC = True Then
       For k = 0 To 15
           If bY0(k) = 0 Then
              PLCstatefrm.ShY(k).FillColor = &HC0FFFF
           Else
              PLCstatefrm.ShY(k).FillColor = &HFF&
           End If
           
           If bX0(k) = 0 Then
              PLCstatefrm.ShX(k).FillColor = &HC0FFFF
           Else
              PLCstatefrm.ShX(k).FillColor = &HFF&
           End If
           
           If bX20(k + 20) = 0 Then
              PLCstatefrm.ShX(k + 16).FillColor = &HC0FFFF
           Else
              PLCstatefrm.ShX(k + 16).FillColor = &HFF&
           End If
            
           If bY20(k + 20) = 0 Then
              PLCstatefrm.ShY(k + 16).FillColor = &HC0FFFF
           Else
              PLCstatefrm.ShY(k + 16).FillColor = &HFF&
           End If
       Next
    End If
    
    If bY0(0) = False Then
       Inspectfrm.cmdPLCOnOff.BackColor = QBColor(7) ' "MC OFF"
       Inspectfrm.cmdPLCOnOff.Caption = "Spasi ON"
    Else
       Inspectfrm.cmdPLCOnOff.BackColor = QBColor(10) ' "MC ON"
       Inspectfrm.cmdPLCOnOff.Caption = "Spasi OFF"
    End If
    
    If bY0(1) = False Then
       Inspectfrm.cmdPLCFW.BackColor = QBColor(12) '"Ç°³µ"
       Inspectfrm.cmdPLCBW.BackColor = QBColor(7)
    Else
       Inspectfrm.cmdPLCBW.BackColor = QBColor(12) '"µ¹³µ"
       Inspectfrm.cmdPLCFW.BackColor = QBColor(7)
    End If
End Sub

Public Sub sNumToBin()
    strPlcX0 = ""
    strPlcX20 = ""
    strPlcY0 = ""
    strPlcY20 = ""

    Dim nLoop As Integer
    For nLoop = 0 To 15
        strPlcX0 = IIf(PLCdata(5) And 2 ^ nLoop, "1", "0") & strPlcX0
        strPlcX20 = IIf(PLCdata(6) And 2 ^ nLoop, "1", "0") & strPlcX20
        strPlcY0 = IIf(PLCdata(7) And 2 ^ nLoop, "1", "0") & strPlcY0
        strPlcY20 = IIf(PLCdata(8) And 2 ^ nLoop, "1", "0") & strPlcY20
    Next
End Sub

Public Sub GetPLCdata()
    Inspectfrm.PLC.ReadPLCdata PLCdata
    sNumToBin
End Sub

Public Sub StateCheck()
    Dim i, j As Integer
    bNext = True
    For i = 0 To 15
        If bY0(i) = False Then
            PLCstatefrm.ShY(i).FillColor = &HC0FFFF
        Else
            PLCstatefrm.ShY(i).FillColor = &HFF&
        End If
        
        If bX0(i) = False Then
            PLCstatefrm.ShX(i).FillColor = &HC0FFFF
        Else
            PLCstatefrm.ShX(i).FillColor = &HFF&
        End If
        
        If bX20(i + 20) = False Then
            PLCstatefrm.ShX(i + 16).FillColor = &HC0FFFF
        Else
            PLCstatefrm.ShX(i + 16).FillColor = &HFF&
        End If
    
        If bY20(i + 20) = False Then
            PLCstatefrm.ShY(i + 16).FillColor = &HC0FFFF
        Else
            PLCstatefrm.ShY(i + 16).FillColor = &HFF&
        End If
    Next
End Sub

