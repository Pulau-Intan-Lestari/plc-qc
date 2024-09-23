Option Explicit

'Dim rsRolKain As ADODB.Recordset
'Dim rsDefectRol As ADODB.Recordset

'Dim bFirstInspectRol As Boolean
Dim nPLCStopLen, nYards As Single

'Dim sRolSelected As String
Private Sub cmdF4NewRol_Click()
    NewRolfrm.Show vbModal
End Sub

Private Sub cmdF5Spec_Click()
    Specfrm.Show vbModal
End Sub

Private Sub cmdF6Edit_Click()
    Editfrm.Show vbModal
End Sub

Private Sub cmdOpenOrder_Click()
    OrderListfrm.Show vbModal       'NumPadCtl_KeyDown 120   'F9
End Sub

Private Sub cmdPLCBW_Click()
    PLC.Switch_M 120
End Sub

Private Sub cmdPLCBW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.SetFocus
End Sub

Private Sub cmdPLCFW_Click()
    PLC.Switch_M 110
End Sub

Private Sub cmdPLCFW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.SetFocus
End Sub

Private Sub cmdPLCOnOff_Click()
    PLC.Switch_M 101
End Sub

Private Sub cmdPLCOnOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.SetFocus
End Sub

Private Sub cmdSelesai_Click()
    NumPadCtl.Enable = False
    If MsgBox("SELESAI INSPECT PO " & txtNoso & " ?", vbQuestion + vbYesNo, _
              "Finish Inspect") = vbYes Then
       CloseOrder
    End If
    NumPadCtl.Enable = True
End Sub



Private Sub dgRolKain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    sRolNoInspected = Null2String(dgRolKain.Columns(1))
End Sub

Private Sub XForm_Activate()
    'NB :
    'di non aktifkan karena menggangu pemilihan order di form Orderlist
    
    'NumPadCtl.Enable = True
    
    'jika tombol NumLock OFF maka ON / Aktifkan
    'If NumPadCtl.NumLockState = False Then
    '   NumPadCtl.NumLockPress
    'End If
    
    'cmdF4NewRol.SetFocus
End Sub

Private Sub Form_Activate()
'    Me.NumPadCtl.SetFocus
    NumPadCtl.Enable = True
End Sub

Private Sub Form_GotFocus()
    bNumKeyIsDown = False
End Sub

Private Sub Form_Initialize()
    ModeFirstLoad
End Sub

Public Sub ModeFirstLoad()
    bGetAutoStop = True     'COUNTER AUTO STOP ROLLING
    
    nSOBatchNo = -9999      'BATCHNO PO YG SEDANG DI INSPECT
    sRolNoInspected = ""    'NO ROL YG SEDANG DI INSPECT
    sKeyPressed = ""        'HURUF/ANGKA TOMBOL YG DI TEKAN
    
    'Set rsRolKain = New ADODB.Recordset
    'Set rsDefectRol = New ADODB.Recordset
   
    ClearTextBox        'clear detail so
    ShowDefectList      'tampilkan daftar defect
    
    ShowDetailRol       'tampilkan detail rol
    ShowDefectData      'tampilkan detail defect rol
    
    Me.cmdF2SaveRol.Enabled = False
    Me.cmdF3SaveOrder.Enabled = False
    Me.cmdF4NewRol.Enabled = False
    Me.cmdF5Spec.Enabled = False
    Me.cmdF6Edit.Enabled = False
    Me.cmdSelesai.Enabled = False
    Me.cmdOpenOrder.SetFocus
End Sub

Public Sub ModeReady()
    bGetAutoStop = True     'COUNTER AUTO STOP ROLLING
    
    'Tampilkan Detail Rol Data Inspect yg ada sebelumnya
    ShowDetailRol
    ShowDefectData
    
    txtTmp(0) = nSOBatchNo
    txtTmp(1) = sRolNoInspected
    
    Me.cmdF2SaveRol.Enabled = False
    Me.cmdF3SaveOrder.Enabled = False
    Me.cmdF4NewRol.Enabled = True
    Me.cmdF5Spec.Enabled = True
    Me.cmdF6Edit.Enabled = True
    Me.cmdSelesai.Enabled = True
    Me.cmdOpenOrder.Enabled = True
    'Me.cmdF4NewRol.SetFocus
       
    'SUDAH DI FORM INPUT ROL BARU
    'ResetCounter
    
    NumPadCtl.Enable = True
End Sub

Private Sub Form_Load()
    'MsgBox Me.Left & " " & Me.Top & " " & Me.Width & " " & Me.Height
    
'    Set rsRolKain = New ADODB.Recordset
'    Set rsDefectRol = New ADODB.Recordset
    
    With Me
      .Left = 0
      .Top = 0
      .Width = 15300
      .Height = 10005
    End With
    
    Dim i As Integer
    For i = 0 To 19
        txtDef1(i).Locked = True
    Next
    
    lblLength = "0.00"
    lblWeight = "0.00"
    
    'bGetAutoStop = True        'counter utk Auto Stop Rolling nn.nn Yd
    
    PLC.PLCportOpen = True
    
    'tidak usah di reset pada saat awal membuka program
    'If bAutoReset Then PLC.Switch_M 140
    
    'reset PLC counter yard
    PLC.Switch_M 140
    
    Dim Ttimer As Single
    Ttimer = Timer
    Do While Timer < Ttimer + 0.2
       DoEvents
    Loop
    
    'counter NumLock Control
    bNumKeyIsDown = False
    
    If bViewDataPLC Then
       frDataPLC.Visible = True
    Else
       frDataPLC.Visible = False
    End If
End Sub

Private Sub ClearTextBox()
    txtOrderNo = ""
    txtWarna = ""
    txtLot = ""
    
    txtFabric = ""
    txtYarn = ""
    txtKdwn = ""
    txtJenis = ""
    
    txtFinish = ""
    txtGrm2 = ""
    txtNoso = ""
    'txtCust = ""
    
    ShowDefectList
End Sub

Public Function OpenOrder()
    'With rsOrder
    '  If .State Then .Close
    '  .Open "Select", dbQCHAE
    '  If .BOF And .EOF Then
    '     MsgBox "NOMOR ORDER " & txtNoso & " TIDAK ADA !", vbExclamation, App.Title
    '     Exit Sub
    '  End If
    'End With
End Function

Public Function CloseOrder()
    ModeFirstLoad
    ClearTextBox
End Function

Public Sub ShowDetailRol()  '(Optional ByVal bGetRolNo As Boolean)
                            'If IsMissing(bGetRolNo) Then bGetRolNo = True
    'tampilkan Detail Rol Kain
    With rsRolKain
      Set dgRolKain.DataSource = Nothing
      
      If .State Then .Close
      .Open "Select BATCHNO,ROLNO,LOT,LENGTH,WEIGHTNETTO,WIDTH,WEIGHT,DEFECTPOINTS," & _
            "GRADE,INSPECTOR,INSPECTTIME From DETAILROL Where BATCHNO=" & nSOBatchNo & _
            " Order By INSPECTTIME", dbQCHAE
            
      '.Open "Select BATCHNO,ROLNO,LOT,WIDTH,WEIGHT,DEFECTPOINTS,GRADE," & _
            "INSPECTOR,INSPECTTIME From DETAILROL Where BATCHNO=" & nSOBatchNo & _
            " Order By INSPECTTIME", dbQCHAE
      
      Set dgRolKain.DataSource = rsRolKain
      
      SetGridRol
      
      sRolNoInspected = ""
      
      If Not (.BOF And .EOF) Then
        .MoveLast
         sRolNoInspected = !ROLNO         'If bGetRolNo Then sRolNoInspected = !ROLNO
         
         Dim rsTmp As adodb.Recordset
         Set rsTmp = New adodb.Recordset
         
         With rsTmp
           .Open "Select DEF0,DEF1,DEF2,DEF3,DEF4,DEF5,DEF6,DEF7,DEF8,DEF9," & _
                 "DEF10,DEF11,DEF12,DEF13,DEF14,DEF15,DEF16,DEF17,DEF18,DEF19" & _
                 " From DETAILROL Where BATCHNO=" & nSOBatchNo & _
                 " And ROLNO='" & sRolNoInspected & "'", dbQCHAE
                 
           If Not (.BOF And .EOF) Then
              Dim i As Integer
              For i = 0 To 19
                  txtDef1(i) = .Fields(i)
              Next
           End If
           .Close
           Set rsTmp = Nothing
         End With
      End If
    End With
End Sub

Private Sub SetGridRol()
    Dim i As Integer
    
    With dgRolKain
      'For i = 0 To rsRolKain.Fields.Count - 1   'kRJmlKol - 1     '
      '   .Columns(i).Width = 0
      'Next
      
      .Columns(0).Width = 0
      
      .Columns(1).Caption = "NO. ROL"
      .Columns(1).Width = 900
      .Columns(1).Alignment = dbgCenter
      
      .Columns(2).Caption = "LOT"
      .Columns(2).Width = 900
      .Columns(2).Alignment = dbgCenter
      
      .Columns(3).Caption = "LENGTH"
      .Columns(3).Width = 1100
      .Columns(3).Alignment = dbgCenter
      
      .Columns(4).Caption = "WEIGHT (KG)"
      .Columns(4).Width = 1100
      .Columns(4).Alignment = dbgCenter
      
      .Columns(5).Caption = "WIDTH"
      .Columns(5).Width = 1100
      .Columns(5).Alignment = dbgCenter
      
      .Columns(6).Caption = "WEIGHT"
      .Columns(6).Width = 1100
      .Columns(6).Alignment = dbgCenter
      
      .Columns(7).Caption = "DEFECT POINTS"
      .Columns(7).Width = 1200
      .Columns(7).Alignment = dbgCenter
      
      .Columns(8).Caption = "GRADE"
      .Columns(8).Width = 0 '900
      .Columns(8).Alignment = dbgCenter
      
      .Columns(9).Caption = "QC BY"
      .Columns(9).Width = 1600
      .Columns(9).Alignment = dbgLeft
      
      .Columns(10).Width = 0
    End With
End Sub

Public Sub ShowDefectData()
    'tampilkan Detail Defect tiap Rol Kain yg di "select"
    With rsDefectRol
      Set dgDefects.DataSource = Nothing
      
      If .State Then .Close
      
      'MsgBox nSOBatchNo & " " & sRolNoInspected
      
      .Open "Select r.BATCHNO,r.ROLNO,r.DEFCODE,d.DEFNAME,r.DEFPOINTS," & _
            "r.DEFPOSITION,r.REMARK From ROLDEFECTS r" & _
            " Left Join DEFECTS d On r.DEFCODE=d.DEFCODE" & _
            " Where r.BATCHNO=" & nSOBatchNo & _
            " And r.ROLNO='" & sRolNoInspected & "'", dbQCHAE
      
      Set dgDefects.DataSource = rsDefectRol
            
      SetGridDefectRol
      
      If Not (.BOF And .EOF) Then
        .MoveLast
      End If
    End With
End Sub

Private Sub SetGridDefectRol()
    Dim i As Integer
    
    With dgDefects
      'For i = 0 To rsDefectRol.Fields.Count - 1     'kDJmlKol - 1     '
      '   .Columns(i).Width = 0
      'Next
      
      .Columns(0).Width = 0
      .Columns(1).Width = 0
      .Columns(2).Width = 0
      
      .Columns(3).Caption = "DEFECT FOUNDED"
      .Columns(3).Width = 1700
      .Columns(3).Alignment = dbgLeft
      
      .Columns(4).Caption = "POINTS"
      .Columns(4).Width = 900
      .Columns(4).Alignment = dbgCenter
      
      .Columns(5).Caption = "POSITION"
      .Columns(5).Width = 1100
      .Columns(5).Alignment = dbgCenter
      
      .Columns(6).Caption = "REMARK"
      .Columns(6).Width = 3000
      .Columns(6).Alignment = dbgLeft
    End With
End Sub

Public Sub ShowDefectList(Optional tmpDefID As String)
    Dim rsDef As adodb.Recordset
    Dim i, trec As Integer
    
    Set rsDef = New adodb.Recordset
    
    With rsDef
       .Open "Select * From DEFECTS Order By DEFCODE", dbQCHAE
       
       If (.BOF And .EOF) Then
          MsgBox "DAFTAR DEFECT BELUM DI-ISI !", vbCritical
          Exit Sub
       Else
          trec = .RecordCount
          
          For i = 0 To 19
              lblDef1(i) = ""
              cmdKey1(i).Caption = ""
              txtDef1(i) = 0
          Next
          
          For i = 1 To trec
              Def(i - 1) = Trim(!DEFNAME)   'Trim(.Fields(i + 3))
              lblDef1(i - 1) = Trim(!DEFNAME)
              cmdKey1(i - 1).Caption = Trim(!DEFKEY)
              .MoveNext
          Next i
       End If
         
      .Close
       Set rsDef = Nothing
    End With
End Sub

Private Sub ResetCounter()
    If bAutoReset Then
       If bAskReset Then
          If MsgBox("RESET COUNTER YARDS PANJANG KAIN KE ANGKA NOL (0.00) ?", vbExclamation + vbYesNo, "Reset Counter Yards") = vbYes Then
             PLC.Switch_M 140
          End If
       End If
    End If
End Sub

Private Sub lblLength_Click()

End Sub

Private Sub NumPadCtl_KeyDown(KeyCode As Integer)

'    If bNumKeyIsDown = False Then
    
       If Not (KeyCode >= 48 And KeyCode <= 57) Then
          If Not (KeyCode >= 65 And KeyCode <= 90) Then
             If Not (KeyCode >= 96 And KeyCode <= 105) Then
                If Not (KeyCode = 107) Then
                   If Not (KeyCode = 109) Then
                      If Not (KeyCode >= 115 And KeyCode <= 120) Then
                         'If Not (KeyCode = 13) Then
                             If Not (KeyCode = 32) Then
                                
                                bNumKeyIsDown = True
                                Exit Sub
                                
                             End If
                         'End If
                      End If
                   End If
                End If
             End If
          End If
       End If
       
       'If (KeyCode >= 65 And KeyCode <= 90) Or _
          (KeyCode >= 48 And KeyCode <= 57) Or _
          (KeyCode >= 96 And KeyCode <= 105) Then
          
'          If (KeyCode >= 96 And KeyCode <= 105) Then
'             Select Case KeyCode
'             Case 96: sKeyPressed = "0"
'             Case 97: sKeyPressed = "1"
'             Case 98: sKeyPressed = "2"
'             Case 99: sKeyPressed = "3"
'             Case 100: sKeyPressed = "4"
'             Case 101: sKeyPressed = "5"
'             Case 102: sKeyPressed = "6"
'             Case 103: sKeyPressed = "7"
'             Case 104: sKeyPressed = "8"
'             Case 105: sKeyPressed = "9"
'             End Select
'          Else
'             sKeyPressed = Chr(KeyCode)
'          End If
'
'          Dim i As Integer
'          Dim bFound As Boolean
'          bFound = False
'          For i = 0 To 19
'              If sKeyPressed = cmdKey1(i).Caption Then
'                 bFound = True
'              End If
'          Next
'
'          Pointsfrm.Show vbModal
'
'          Dim Ttimer As Single
'          Ttimer = Timer
'          Do While Timer < Ttimer + 1
'             DoEvents
'          Loop
       'End If
       
       Select Case KeyCode
       
       Case 113    'F2 SaveRollData
            'If Me.cmdSaveRollData.Enabled = True Then
            '   Me.cmdSaveRollData.Value = True
            'End If
       
       Case 114    'F3 SaveBatchData
            'If Me.cmdSaveBatchData.Enabled = True Then
            '   Me.cmdSaveBatchData.Value = True
            'End If
       
       Case 115     'F4 New Roll
            If txtNoso <> "" Then NewRolfrm.Show vbModal
            
       Case 116     'F5 Spec Fabric
            If txtNoso <> "" Then Specfrm.Show vbModal
            
       Case 117     'F6 Edit
            If txtNoso <> "" Then Editfrm.Show vbModal
            
       Case 118     'F7 Note
            If txtNoso <> "" Then Notefrm.Show vbModal
            
       Case 119     'F8 Selesai
            If txtNoso <> "" Then cmdSelesai.Value = True
            
       Case 120     'F9 Open Order
            OrderListfrm.Show vbModal
            
'       Case 13    'ENTER
'            cmdPLCOnOff.Value = True           'Me.cmdOnOff.Value = True
       
       Case 32    'SPACE
            cmdPLCOnOff.Value = True           'Me.cmdOnOff.Value = True
       
       Case 107    '+
            cmdPLCFW.Value = True
       
       Case 109    '-
            cmdPLCBW.Value = True
       
       Case Else
            
            If (KeyCode >= 96 And KeyCode <= 105) Then
               Select Case KeyCode
               Case 96: sKeyPressed = "0"
               Case 97: sKeyPressed = "1"
               Case 98: sKeyPressed = "2"
               Case 99: sKeyPressed = "3"
               Case 100: sKeyPressed = "4"
               Case 101: sKeyPressed = "5"
               Case 102: sKeyPressed = "6"
               Case 103: sKeyPressed = "7"
               Case 104: sKeyPressed = "8"
               Case 105: sKeyPressed = "9"
               End Select
            Else
               sKeyPressed = Chr(KeyCode)
            End If
          
            Dim i As Integer
            Dim bFound As Boolean
            bFound = False
            For i = 0 To 19
                If sKeyPressed = cmdKey1(i).Caption Then
                   bFound = True
                   Pointsfrm.txtDefName = lblDef1(i).Caption
                   Pointsfrm.lblDefNo.Caption = i
                End If
            Next
         
            If bFound Then
               
               'tampilkan form isi data inspect
               Pointsfrm.Show vbModal
               
               bNumKeyIsDown = True
               
               'Dim Ttimer As Single
               'Ttimer = Timer
               'Do While Timer < Ttimer + 1
               '   DoEvents
               'Loop
            End If
         End Select
'    End If
    
    bNumKeyIsDown = True
End Sub

Private Sub NumPadCtl_KeyUp(KeyCode As Integer)
'    If bNumKeyIsDown = True Then
'       bNumKeyIsDown = False
'    End If
End Sub

Private Sub Picture1_Click()
    Picture1.SetFocus
End Sub

Private Sub Picture1_GotFocus()
    Picture1.BackColor = vbRed
End Sub

Private Sub Picture1_LostFocus()
    Picture1.BackColor = vbBlack
End Sub

Private Sub PLC_OnComm()

    GetPLCdata
    PlcYState
    
    lblDataPlc.Caption = PLCdata(0)
    
    nYards = Format((PLCdata(0) / 10) + nYardsAdjustment, "##0.00")
    
    lblLength.Caption = Format((PLCdata(0) / 10) + nYardsAdjustment, "##0.00")

    If bAutoStop Then       'JIKA AUTO STOP ROLLING DI AKTIFKAN
       
       If bGetAutoStop Then 'JIKA BELUM AUTO STOP / POSISI KEPALA KAIN
    
          If nYards >= nAutoStop Then   'JIKA POSISI AUTO STOP
             
             bGetAutoStop = False
             
             cmdPLCOnOff.Value = True
             
             'Specfrm.Show vbModal
             
             'NewRolfrm.Show vbModal
          End If
       End If
    End If
End Sub

Public Function M2Yd(tM As Single) As Single
    M2Yd = Round(tM / (25.4 * 12 * 3 / 1000), 2)
End Function

Public Function Yd2M(tYD As Single) As Single
    '1ydÂë£½25.4*12*3/1000 m
    'yd2m = Round(tYD * 25.4 * 12 * 3 / 1000, 2)
    Yd2M = Round(tYD * 0.914, 2)
End Function
