Option Explicit

'Regional Setting Set
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'-----

Public dbQCHAE As adodb.Connection      'MS SQL QCHAE
Public dbPPC As adodb.Connection        'MS SQL PPC
Public MyPPC As adodb.Connection        'MY SQL PPC
Public dbPRODUCTION As adodb.Connection        'MS SQL PRODUCTION
Public dbKNITTING As adodb.Connection        'MS SQL PRODUCTION

Public rsOrder As New adodb.Recordset
Public rsRolKain As New adodb.Recordset
Public rsDefectRol As New adodb.Recordset


Public ServerName, sNetLib As String
Public DatabaseName As String

Public dbUserName As String
Public dbPassword As String

Public sUserID, sUserName As String
Public nUserLevel, nUserGroup As Integer

Public Def(20) As String
Public sMachID, sBatchno, sKeyPressed As String
Public sNoso, sFabric, sWarna, sLot, sFinish, sWeight As String

Public nSOBatchNo As Long
Public sRolNoInspected As String

Public bAutoStop, bRolByKey, bAutoReset, bAskReset, _
       bNumKeyIsDown, bViewDataPLC, bGetAutoStop As Boolean

Public nAutoStop, nYardsAdjustment As Single     'auto stop rolling machine when beginning fabric inspection
Public sMachineID As String     'machine identity

Public QCName As String         'nama petugas sortir

Public Const LvlUser = 0
Public Const LvlSpv = 1
Public Const LvlAcct = 2
Public Const LvlPower = 3
Public Const LvlAdmin = 4

'Regional Setting Declar
Private Const REG_SZ = 1

Private Const HKEY_CURRENT_USER = &H80000001
Private Const LOCALE_SENGCOUNTRY     As Long = &H1002 '  English name of country
Private Const LOCALE_USER_DEFAULT    As Long = &H400
Private Const REGIONAL_SETTING       As String = "United States"
'------------------------------------------

'---------------------------------------
'variabel saat compile pada koneksi
'---------------------------------------
'Public Const koneksiprogram = "lokal"
Public Const koneksiprogram = "server"
'---------------------------------------

Sub Main()
'On Error GoTo StartErr

    If App.PrevInstance Then
      MsgBox "ANDA SUDAH MEMBUKA PROGRAM !;LIHAT DI BARIS START (Windows Taskbar);ATAU TEKAN ALT+TAB !", vbInformation, "PIL - Inspection System"
      End
    End If

    Call setDefRegionalSetting

    MouseP 11
    
    '*** ReadIni  'hapus parameter readini 18/11/2014
    
    sMachID = ""
    'If Dir("C:\MAC01.HAE") <> "" Then sMachID = "MAC01"
    'If Dir("C:\MAC02.HAE") <> "" Then sMachID = "MAC02"
    'If Dir("C:\MAC03.HAE") <> "" Then sMachID = "MAC03"
    'If Dir("C:\MAC04.HAE") <> "" Then sMachID = "MAC04"
    
    
    ' open connection to QCHAE database MS SQL
    Set dbQCHAE = New adodb.Connection
    
    With dbQCHAE
      .CursorLocation = adUseClient   'adUseServer
    
    If koneksiprogram = "lokal" Then
      .Open "PROVIDER=SQLOLEDB;" & _
            "Data Source=VMWARE-XP-AGUS;" & _
            "Initial Catalog=QCHAE;" & _
            "User Id=sa;" & _
            "Password=310124;"
    Else
      .Open "PROVIDER=SQLOLEDB;" & _
            "Data Source=192.168.3.4;" & _
            "Initial Catalog=QCHAE;" & _
            "User Id=sa;" & _
            "Password=admin89;"
    End If
    
      .Execute "set dateformat dmy"
    End With
    
    
    ' open connection to PPC database MS SQL
    Set dbPPC = New adodb.Connection
    
    With dbPPC
      .CursorLocation = adUseClient
    If koneksiprogram = "lokal" Then
      .Open "Provider=SQLOLEDB.1;" & _
            "Persist Security Info=False;" & _
            "User ID=sa;" & _
            "Password=310124;" & _
            "Initial Catalog=PPC;" & _
            "Data Source=VMWARE-XP-AGUS"
    Else
      .Open "Provider=SQLOLEDB.1;" & _
            "Persist Security Info=False;" & _
            "User ID=sa;" & _
            "Password=admin89;" & _
            "Initial Catalog=PPC;" & _
            "Data Source=192.168.3.4"
    End If
       .Execute "set dateformat dmy"
    End With
    


    ' open connection to PPC database MySql
    Set MyPPC = New adodb.Connection

    With MyPPC
      .CursorLocation = adUseClient   'adUseServer
      .ConnectionTimeout = 5
      
    If koneksiprogram = "lokal" Then
      .Open "DRIVER={MySQL ODBC 3.51 Driver};" _
           & "SERVER=127.0.0.1;" _
           & "DATABASE=ppc1;" _
           & "UID=user;" _
           & "PWD=user;" _
           & "OPTION=16387;STMT=;" & Chr$(34)
    Else
      .Open "DRIVER={MySQL ODBC 3.51 Driver};" _
           & "SERVER=192.168.3.1;" _
           & "DATABASE=ppc;" _
           & "UID=user;" _
           & "PWD=user;" _
           & "OPTION=16387;STMT=;" & Chr$(34)
    End If
    
    End With
    
    
    ' open connection to PRODUCTION database MSSQL
    Set dbPRODUCTION = New adodb.Connection
    
    With dbPRODUCTION
      .CursorLocation = adUseClient
    If koneksiprogram = "lokal" Then
      .Open "Provider=SQLOLEDB.1;" & _
            "Persist Security Info=False;" & _
            "User ID=sa;" & _
            "Password=310124;" & _
            "Initial Catalog=PRODUCTION;" & _
            "Data Source=VMWARE-XP-AGUS"
    Else
      .Open "Provider=SQLOLEDB.1;" & _
            "Persist Security Info=False;" & _
            "User ID=sa;" & _
            "Password=admin89;" & _
            "Initial Catalog=PRODUCTION;" & _
            "Data Source=192.168.3.4"
    End If
       .Execute "set dateformat dmy"
    End With
    
    ' open connection to KNITTING database MSSQL
    Set dbKNITTING = New adodb.Connection
    
    With dbKNITTING
      .CursorLocation = adUseClient
    If koneksiprogram = "lokal" Then
      .Open "Provider=SQLOLEDB.1;" & _
            "Persist Security Info=False;" & _
            "User ID=sa;" & _
            "Password=310124;" & _
            "Initial Catalog=KNITTING;" & _
            "Data Source=VMWARE-XP-AGUS"
    Else
      .Open "Provider=SQLOLEDB.1;" & _
            "Persist Security Info=False;" & _
            "User ID=sa;" & _
            "Password=admin89;" & _
            "Initial Catalog=KNITTING;" & _
            "Data Source=192.168.3.4"
    End If
       .Execute "set dateformat dmy"
    End With
    
    
    
' rubah error trap utk client yg di-Deepfreeze
' tidak bisa edit jam & tanggal di pc lokal
On Error Resume Next

    Dim rsTemp As adodb.Recordset
    Set rsTemp = New adodb.Recordset
    
    rsTemp.Open "Select CURRENT_TIMESTAMP();", MyPPC
    
    Date = rsTemp.Fields(0)
    Time = rsTemp.Fields(0) 'format(, "hh:mm:ss")
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    MouseP
    
    
    '---Set Paramater Inspect System
    On Error Resume Next
    
    Dim rsTmp As adodb.Recordset
    Set rsTmp = New adodb.Recordset

    'Set rsTmp = dbQCHAE.Execute("GetServerTime", , adCmdStoredProc)
    '
    'Date = rsTmp.Fields(0)
    'Time = rsTmp.Fields(0)
      
    With rsTmp
      '.Close
      
      .Open "Select * From SYSPARS Where MACHID='" & sMachID & "'", dbQCHAE
      
      If Not (.BOF And .EOF) Then
         
         'AUTO STOP INSPECT KEPALA KAIN
         bAutoStop = False
         If !STOPROLLING = "Y" Then bAutoStop = True
         
         'AUTO STOP POSITION
         nAutoStop = Format(!AUTOSTOP, "#0.0")
         
         'YARDS LENGTH ADJUSTMENT
         nYardsAdjustment = Format(!YARDSADJUSTMENT, "#0.0")
         
         'AUTO ROLLING SETELAH INPUT NO ROL BARU/BERIKUTNYA
         bRolByKey = False
         If !ROLBYKEY = "Y" Then bRolByKey = True
         
         'AUTO RESET COUNTER YARDS KAIN KE ANGKA NOL (0)
         bAutoReset = False
         If !AUTORESET = "Y" Then bAutoReset = True
         
         'TANYA DULU SEBELUM COUNTER YARDS KAIN DI-RESET KE ANGKA NOL (0)
         bAskReset = False
         If !ASKRESETCOUNTER = "Y" Then bAskReset = True
         
         bViewDataPLC = False
         If !VIEWDATAPLC = "Y" Then bViewDataPLC = True
      End If
      
      .Close
      Set rsTmp = Nothing
    End With
    
    'set recordset
    Set rsOrder = New adodb.Recordset
    Set rsRolKain = New adodb.Recordset
    Set rsDefectRol = New adodb.Recordset
    
    
    nSOBatchNo = -1000          'BATCHNO PO YG SEDANG DI INSPECT
    sRolNoInspected = ""    'NO ROL YG SEDANG DI INSPECT
    
    'MainForm.Show
    Load Inspectfrm
    
    Exit Sub
    

StartErr:
    MsgBox Err.Description, vbCritical, App.Title
    End
End Sub
                                                                                                            
'Public Sub ReadIni()
'    Dim sText, sViewDataPLC As String
'    Dim i As Integer
'
'    If Dir(App.Path & "\AGUS.INI") <> "" Then
'       Open App.Path & "\AGUS.INI" For Input Shared As #1
'    Else
'       Open App.Path & "\QCHAE.INI" For Input Shared As #1
'    End If
'
'    For i = 1 To 7
'        Input #1, sText
'
'        sText = Trim(sText)
'
'        Select Case i
'          Case 1: ServerName = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'          Case 2: DatabaseName = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'          Case 3: dbUserName = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'          Case 4: dbPassword = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'          Case 5: sNetLib = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'          Case 6: sViewDataPLC = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'          Case 7: sUserName = Right(sText, Len(sText) - InStr(1, sText, "=", vbTextCompare))
'        End Select
'    Next i
'
'    Close #1
'
'    bViewDataPLC = False
'    If UCase(sViewDataPLC) = "Y" Then bViewDataPLC = True
'End Sub

Public Sub BlockText()
    SendKeys "{Home}+{end}"
End Sub


Function Null2String(sValue As Variant)
    If IsNull(sValue) Then
        Null2String = ""
    Else
        If Not IsEmpty(sValue) Then
           Null2String = sValue
        Else
           Null2String = ""
        End If
    End If
End Function


Function Null2Zero(nValue As Variant, Optional nType As Integer)
'On Error GoTo ErrorNih
    
    If IsMissing(nType) Then _
       nType = VarType(nValue)
    
    If IsNull(nValue) Then
       Null2Zero = CInt(0)
    
    ElseIf IsEmpty(nValue) Then
           Null2Zero = CInt(0)
           
    ElseIf Val(nValue) = 0 Then
           Null2Zero = CInt(0)
    Else
        Select Case nType           'VarType(nValue)
        Case vbInteger
             Null2Zero = CInt(nValue)
        Case vbLong
             Null2Zero = CLng(nValue)
        Case vbSingle
             Null2Zero = CSng(nValue)
        Case vbDouble
             Null2Zero = CDbl(nValue)
        Case vbCurrency
             Null2Zero = CCur(nValue)
        Case vbDecimal
             Null2Zero = CDec(nValue)
        Case Else
             Null2Zero = CSng(nValue)
        End Select
    End If
    
    Exit Function
    
ErrorNih:
    MsgBox nValue & " " & nType
End Function


Function MouseP(Optional mVal As Integer)
    If IsMissing(mVal) Then mVal = 0
    Screen.MousePointer = mVal
End Function


Function InaCMONTH(dDate) As String
    Dim nBulan As Integer
    
    If VarType(dDate) = vbDate Then
       nBulan = Month(dDate)
    Else
       nBulan = Val(dDate)
       'If VarType(dDate) = vbInteger Then
       '   nBulan = dDate
       'Else
       '   nBulan = 0
       'End If
    End If
    
    Select Case nBulan
       Case 1
            InaCMONTH = "Januari"
       Case 2
            InaCMONTH = "Februari"
       Case 3
            InaCMONTH = "Maret"
       Case 4
            InaCMONTH = "April"
       Case 5
            InaCMONTH = "Mei"
       Case 6
            InaCMONTH = "Juni"
       Case 7
            InaCMONTH = "Juli"
       Case 8
            InaCMONTH = "Agustus"
       Case 9
            InaCMONTH = "September"
       Case 10
            InaCMONTH = "Oktober"
       Case 11
            InaCMONTH = "November"
       Case 12
            InaCMONTH = "Desember"
       Case Else
            InaCMONTH = "."
    End Select
End Function


Public Function IsFormLoaded(sFormName As String) As Boolean
Dim lForm As Long
    
    For lForm = 0 To VB.Forms.Count - 1
        If VB.Forms(lForm).Name = sFormName Then
            IsFormLoaded = True
            Exit For
        End If
    Next
End Function


Function ReplaceKey(KeyAscii As Integer, Optional bUpperCase As Boolean) As Integer
    
    If IsMissing(bUpperCase) Then bUpperCase = False
    
    If KeyAscii = Asc("'") Then
       ReplaceKey = Asc("`")
    Else
       If bUpperCase Then
          ReplaceKey = Asc(UCase(Chr(KeyAscii)))
       Else
          ReplaceKey = KeyAscii
       End If
    End If
End Function


Function IsNumberKey(KeyAscii As Integer) As Integer
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
       KeyAscii = Asc(".") Or KeyAscii = Asc("-") Or KeyAscii = vbKeyBack Then
       
       IsNumberKey = KeyAscii
    Else
       IsNumberKey = 0
    End If
End Function


Function IsCodeKey(KeyCode As Integer) As Boolean
    If (KeyCode >= Asc("A") And KeyCode <= Asc("Z")) Or _
       (KeyCode >= Asc("a") And KeyCode <= Asc("z")) Or _
       (KeyCode >= Asc("0") And KeyCode <= Asc("9")) Or _
       KeyCode = Asc(".") Or KeyCode = Asc("-") Or KeyCode = vbKeyBack Then
       
       IsCodeKey = True
    Else
       IsCodeKey = False
    End If
End Function

Function BlankTextBox(sForm As Form)
    Dim i As Integer
    
    For i = 0 To sForm.Controls.Count - 1
    
        If TypeOf sForm.Controls(i) Is TextBox Then
           sForm.Controls(i).Text = ""
        End If
    Next
End Function

Private Sub saveToWindowsRegistry(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim ret As Long

    On Error Resume Next

    'Create a new key
    RegCreateKey hKey, strPath, ret
    'Save a string to the key
    RegSetValueEx ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey ret
End Sub

Private Function getInfoRegionalSettings(ByVal lInfo As Long) As String
    Dim buffer As String, ret As String

    On Error Resume Next

    buffer = String$(256, 0)
    ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, buffer, Len(buffer))
    If ret > 0 Then
        getInfoRegionalSettings = Left$(buffer, ret - 1)
    Else
        getInfoRegionalSettings = ""
    End If
End Function

Public Sub setDefRegionalSetting()
Dim TypeRegional As String

TypeRegional = getInfoRegionalSettings(LOCALE_SENGCOUNTRY)
    If getInfoRegionalSettings(LOCALE_SENGCOUNTRY) <> REGIONAL_SETTING Then
    MsgBox "Terditeksi Regional Setting [" & TypeRegional & "] ,Setting Regional di Control Panel Harus .:. English-US .:. ", vbExclamation + vbOKOnly
    If MsgBox("System Akan Mengaktifkan Regional Setting .:. English-US .:. , Ganti Sekarang ?", vbQuestion + vbYesNo, "Perhatian") = vbYes Then
    'MsgBox "ok"
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iCountry", "1")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iCurrDigits", "2")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iCurrency", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iDate", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iDigits", "2")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iLZero", "1")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iMeasure", "1")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iNegCurr", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iTime", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iTLZero", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "Locale", "00000409")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "s1159", "AM")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "s2359", "PM")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sCountry", "United States")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sCurrency", "$")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sDate", "/")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sDecimal", ".")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sLanguage", "ENU")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sList", ",")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sLongDate", "dddd, MMMM dd, yyyy")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate", "M/d/yyyy")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sThousand", ",")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sTime", ":")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sTimeFormat", "h:mm:ss tt")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iTimePrefix", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sMonDecimalSep", ".")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sMonThousandSep", ",")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iNegNumber", "1")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sNativeDigits", "0123456789")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "NumShape", "1")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iCalendarType", "1")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iFirstDayOfWeek", "6")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "iFirstWeekOfYear", "0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sGrouping", "3;0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sMonGrouping", "3;0")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sPositiveSign", "")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International", "sNegativeSign", "-")
        Call saveToWindowsRegistry(HKEY_CURRENT_USER, "Control Panel\International\Geo", "Nation", "244")
Else
    MsgBox "Exit Program ..!!! "
    End
End If
    End If
End Sub
