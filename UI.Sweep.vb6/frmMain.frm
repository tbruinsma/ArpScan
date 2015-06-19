VERSION 5.00
Object = "{F82B5D50-EC04-4350-9FBB-95E421B15F88}#1.0#0"; "PacketX.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "EtherARP"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   7440
      ScaleHeight     =   285
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox cmdPassive 
      Caption         =   "&Passive"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4620
      Width           =   1395
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   26
      Top             =   5070
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Ether ARP Atatus"
            TextSave        =   "Ether ARP Atatus"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboAdapter 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3675
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   2820
      TabIndex        =   23
      Top             =   420
      Width           =   5055
      Begin VB.CheckBox chkTimeOnly 
         Caption         =   "Show Time Only"
         Height          =   255
         Left            =   180
         TabIndex        =   32
         Top             =   630
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkOnPassiveResponseTime 
         Caption         =   "On Passive Get Response Time"
         Height          =   195
         Left            =   2340
         TabIndex        =   31
         Top             =   390
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkResolveOnPassive 
         Caption         =   "On Passive Detect"
         Height          =   195
         Left            =   2340
         TabIndex        =   30
         Top             =   180
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkResolveAfter 
         Caption         =   "Resolve Following Scan"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   375
         Width           =   2115
      End
      Begin VB.CheckBox chkResolveDuring 
         Caption         =   "Resolve During Scan"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   1875
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5160
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C47
            Key             =   "network"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEndIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "255"
      Top             =   1020
      Width           =   315
   End
   Begin VB.TextBox txtEndIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   20
      Text            =   "255"
      Top             =   1020
      Width           =   315
   End
   Begin VB.TextBox txtEndIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   19
      Text            =   "255"
      Top             =   1020
      Width           =   315
   End
   Begin VB.TextBox txtEndIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   18
      Text            =   "255"
      Top             =   1020
      Width           =   315
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1260
      TabIndex        =   17
      Text            =   " ."
      Top             =   1020
      Width           =   135
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1740
      TabIndex        =   16
      Text            =   " ."
      Top             =   1020
      Width           =   135
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2220
      TabIndex        =   15
      Text            =   " ."
      Top             =   1020
      Width           =   135
   End
   Begin VB.TextBox txtStartIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "255"
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox txtStartIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "255"
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox txtStartIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "255"
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox txtStartIP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "255"
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1260
      TabIndex        =   9
      Text            =   " ."
      Top             =   600
      Width           =   135
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1740
      TabIndex        =   8
      Text            =   " ."
      Top             =   600
      Width           =   135
   End
   Begin VB.TextBox txt_ipdot 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2220
      TabIndex        =   7
      Text            =   " ."
      Top             =   600
      Width           =   135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7020
      TabIndex        =   6
      Top             =   4620
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvResults 
      Height          =   2775
      Left            =   60
      TabIndex        =   5
      Top             =   1740
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   767
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MAC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Response Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Detected"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4620
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   315
      Left            =   900
      ScaleHeight     =   255
      ScaleWidth      =   1815
      TabIndex        =   14
      Top             =   540
      Width           =   1875
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   315
      Left            =   900
      ScaleHeight     =   255
      ScaleWidth      =   1815
      TabIndex        =   22
      Top             =   960
      Width           =   1875
   End
   Begin PacketXLibCtl.PacketXCtrl PacketXCtrl1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmMain.frx":B4A2
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblStartIP 
      Caption         =   "Start IP:"
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblEndIP 
      Caption         =   "End IP:"
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1020
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuPassive 
         Caption         =   "&Passive Mode"
      End
      Begin VB.Menu mnuScan 
         Caption         =   "&Scan"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel Scan"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents oSweep As ArpSweep.CArpScanner
Attribute oSweep.VB_VarHelpID = -1

Public cNetAdapters As clsNetworkAdapters
Public oPktCol As PacketCollection

Private Type Host
    IPAddress As String
    MAC As String
    HostName As String
    ResponseTime As Long
End Type

Private oHosts() As Host

Private Sub cboAdapter_Click()
    Dim sStartIP As String
    Dim sEndIP As String
    Dim tempSM As String
    Dim tempip As String
    Dim aryIP() As String
    Dim sBinIP As String
    Dim sBinMask As String
    
    Dim sBinNetIPStart As String
    Dim sNetIPStart As String
    Dim iRange As Integer
    
    oSweep.Adapter = cboAdapter.ListIndex
    tempip = sAdapter(3, cboAdapter.ListIndex)
    tempSM = cNetAdapters.MaskFromIP(tempip)
    
    If tempSM = "" Then Exit Sub
    
    aryIP = Split(tempip, ".")
    sBinIP = convert_to_binary(aryIP(0)) & "." & convert_to_binary(aryIP(1)) & "." & convert_to_binary(aryIP(2)) & "." & convert_to_binary(aryIP(3))
    
    aryMask = Split(tempSM, ".")
    sBinMask = convert_to_binary(aryMask(0)) & "." & convert_to_binary(aryMask(1)) & "." & convert_to_binary(aryMask(2)) & "." & convert_to_binary(aryMask(3))
    sBinNetIPStart = GetBinNetID(sBinIP, sBinMask)
    sNetIPStart = ConvertBinToIP(sBinNetIPStart)
    
    iRange = GetRange(tempSM)
    sEndIP = GetBroadcast(sNetIPStart, tempSM, iRange)
    sStartIP = GetNetStart(sNetIPStart)
    Me.SetStartIP sStartIP
    Me.SetEndIP sEndIP
    
   
End Sub

Private Sub chkResolveAfter_Click()
    If Me.chkResolveDuring.Value Then Me.chkResolveDuring.Value = Abs(Not CBool(Me.chkResolveAfter.Value))
End Sub

Private Sub chkResolveDuring_Click()
    If Me.chkResolveAfter.Value Then Me.chkResolveAfter.Value = Abs(Not CBool(Me.chkResolveDuring.Value))
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdPassive.Enabled = True
    oSweep.StopScan
    ProgressBar1.Value = ProgressBar1.Max
End Sub

Private Sub cmdPassive_Click()
Static bQuiet As Boolean
    If Me.cboAdapter.Text = "" Then
        If Not bQuiet Then MsgBox "You must first select an adapter that will be used to scan.", vbInformation
        bQuiet = True
        cmdPassive.Value = 0
        bQuiet = False
    Else
        If cmdPassive.Value Then
            Me.PacketXCtrl1.Adapter = Me.PacketXCtrl1.Adapters(cboAdapter.ListIndex + 1)
            Me.PacketXCtrl1.Adapter.BuffSize = 32000
            Me.PacketXCtrl1.Adapter.BuffMinToCopy = 0
        
            Me.PacketXCtrl1.Start
        
        
            Me.StatusBar1.Panels(1).Text = "Passive Mode Activated"
            cmdStart.Enabled = False
            Me.cboAdapter.Enabled = False
        Else
            Me.PacketXCtrl1.Stop
            Me.StatusBar1.Panels(1).Text = ""
            cmdStart.Enabled = True
            Me.cboAdapter.Enabled = True
        End If
    
    
    End If
    
End Sub

Private Sub cmdStart_Click()
    oSweep.StartIP = GetStartIP()
    oSweep.EndIP = GetEndIP()
    ProgressBar1.Max = oSweep.HostCount - 1
    lvResults.ListItems.Clear
    Me.cmdCancel.Enabled = True
    Me.cmdPassive.Enabled = False
    oSweep.StartScan
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim x As Integer
    Dim sMAC As String
    Dim sIP As String
    Dim sDesc As String
    Dim sName As String
    
    'Set cAdapters = New clsAdapters
    Set oSweep = New ArpSweep.CArpScanner
    Set cNetAdapters = New clsNetworkAdapters
    Me.PacketXCtrl1.Reset
    
    oSweep.Key = 328760823
    ReDim oHosts(0)
    
    cboAdapter.Clear
    ReDim sAdapter(4, oSweep.AdapterCount - 1)
    
    For x = 0 To oSweep.AdapterCount - 1
        Call oSweep.GetAdapterInfo(x, sName, sDesc, sIP, sMAC)
        If InStr(1, sDesc, "(Microsoft") > 0 Then sDesc = Left(sDesc, InStr(1, sDesc, "(Microsoft") - 1)
        
        sAdapter(0, x) = x
        sAdapter(1, x) = sName
        sAdapter(2, x) = sDesc
        sAdapter(3, x) = sIP
        sAdapter(4, x) = sMAC
        
        Me.cboAdapter.AddItem sDesc, x
    Next
    
    Call ClearStartIP
    Call ClearEndIP
    
    With Me.lvResults.ColumnHeaders
        .Item(1).Width = GetSetting("EtherARP", "frmMain", "lvResults.Col1", 434.8347)
        .Item(2).Width = GetSetting("EtherARP", "frmMain", "lvResults.Col2", 1440)
        .Item(3).Width = GetSetting("EtherARP", "frmMain", "lvResults.Col3", 1470.047)
        .Item(5).Width = GetSetting("EtherARP", "frmMain", "lvResults.Col5", 1440)
        .Item(5).Width = GetSetting("EtherARP", "frmMain", "lvResults.Col5", 1440)
    
        .Item(4).Width = GetSetting("EtherARP", "frmMain", "lvResults.Col4", lvResults.Width - (.Item(1).Width + .Item(2).Width + .Item(3).Width + .Item(5).Width + .Item(6).Width + 50))
        
    End With

End Sub


Public Function GetBroadcast(ByVal strNetID As String, ByVal strmask As String, ByVal ipRange As Integer) As String
  Dim inet As Integer
  Dim ibroad As Integer
  Dim x As Integer
  Dim ipleft As String
  Dim iptemp
  Dim imaskend As Integer
  
  imaskend = GetMaskEndPosition(strmask)
  strNetID = strNetID & "."
  iptemp = Split(strNetID, ".")
  If ipRange = 0 Then ipRange = 256
  For x = 0 To imaskend - 1
    ipleft = ipleft & iptemp(x) & "."
  Next x
  
  If Not ipRange = 256 Then
    For x = 0 To 255 Step ipRange
      Select Case UBound(iptemp) + 1
          Case 1
            sStartIP = ipleft & x & ".0.0.0"
            sEndIP = ipleft & x + (ipRange - 1) & ".255.255.255"
          Case 2
            sStartIP = ipleft & x & ".0.0"
            sEndIP = ipleft & x + (ipRange - 1) & ".255.255"
          Case 3
            sStartIP = ipleft & x & ".0"
            sEndIP = ipleft & x + (ipRange - 1) & ".255"
          Case 4
            sStartIP = ipleft & x
            sEndIP = ipleft & x + (ipRange - 1)
        End Select
      
      If sStartIP = strNetID Then
        GetBroadcast = sEndIP
        Exit Function
      End If
      DoEvents
    Next x
  Else
       iptemp = Split(ipleft & "x", ".")
      Select Case UBound(iptemp) + 1
        Case 1
          sEndIP = ipleft & (ipRange - 1) & ".255.255.255"
        Case 2
          sEndIP = ipleft & (ipRange - 1) & ".255.255"
        Case 3
          sEndIP = ipleft & (ipRange - 1) & ".255"
        Case 4
          sEndIP = ipleft & (ipRange - 1)
      End Select
    DoEvents
  End If
  
  GetBroadcast = sEndIP
End Function

Private Function GetMaskEndPosition(strmask As String) As Integer
  Dim tmpmask, x As Integer
  strmask = strmask & "."
  tmpmask = Split(strmask, ".")
  For x = 0 To UBound(tmpmask) - 1
    Select Case tmpmask(x)
      Case "255"
        GetMaskEndPosition = x + 1
        iMaskEndPosition = x + 1
      Case Else
        Exit For
    End Select
  Next x
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With Me.lvResults.ColumnHeaders
        SaveSetting "EtherARP", "frmMain", "lvResults.Col1", .Item(1).Width
        SaveSetting "EtherARP", "frmMain", "lvResults.Col2", .Item(2).Width
        SaveSetting "EtherARP", "frmMain", "lvResults.Col3", .Item(3).Width
        SaveSetting "EtherARP", "frmMain", "lvResults.Col4", .Item(4).Width
        SaveSetting "EtherARP", "frmMain", "lvResults.Col5", .Item(5).Width
    End With
End Sub

Private Sub Form_Resize()
    If Not frmMain.WindowState = vbMinimized Then
        If frmMain.ScaleHeight >= 2296 + Me.StatusBar1.Height Then
            lvResults.Height = Me.ScaleHeight - 2295 - Me.StatusBar1.Height
            cmdStart.Top = lvResults.Top + lvResults.Height + 105
            cmdCancel.Top = lvResults.Top + lvResults.Height + 105
            cmdPassive.Top = lvResults.Top + lvResults.Height + 105
            
        End If
        If frmMain.ScaleWidth >= 4141 Then
            lvResults.Width = frmMain.ScaleWidth - 120
            cmdCancel.Left = Me.ScaleWidth - 1380
            cmdStart.Left = Me.ScaleWidth - 2760
            cmdPassive.Left = Me.ScaleWidth - 4140
            Me.ProgressBar1.Width = Me.ScaleWidth - 120
            Me.StatusBar1.Panels(1).Width = Me.ScaleWidth
        End If
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub oSweep_OnScanHost(ByVal IP As String, ByVal MAC As String, ByVal response_time As Long)
    'Me.txtResults.Text = Me.txtResults.Text & vbCrLf & IP & vbTab & mac & vbTab & response_time
    Dim Item As ListItem
    
    If response_time > 0 Then
        Set Item = lvResults.ListItems.Add(, "|" & lvResults.ListItems.Count + 1, "", , 1)
        
        Item.SubItems(1) = IP
        Item.SubItems(2) = MAC
        Item.SubItems(3) = ""
        Item.SubItems(4) = response_time
        Item.SubItems(5) = Now()
    End If
    Call UpdateProgress(oSweep.ScannedHostCount)
End Sub

Private Sub oSweep_OnStartStop()
    Call UpdateProgress(oSweep.HostCount - 1)
End Sub

Public Sub UpdateProgress(sCount As Long)
    If Me.ProgressBar1.Max >= sCount Then
        Me.ProgressBar1.Value = sCount
    Else
        Me.cmdCancel.Enabled = False
        Me.cmdPassive.Enabled = True
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    End If
End Sub

Public Sub ClearStartIP()
    Dim Cnt As Integer
    With txtStartIP
        For Cnt = 0 To .UBound
            .Item(Cnt).Text = ""
            .Item(Cnt).Tag = False
        Next Cnt
    End With
End Sub

Public Sub ClearEndIP()
    Dim Cnt As Integer
    With txtEndIP
        For Cnt = 0 To .UBound
            .Item(Cnt).Text = ""
            .Item(Cnt).Tag = False
        Next Cnt
    End With
End Sub

Public Sub SetStartIP(sIP As String)
    Dim aryIP() As String
    Dim x As Integer
    aryIP = Split(sIP, ".")
    For x = 0 To UBound(aryIP)
        If x >= 4 Then Exit For
        txtStartIP(x) = aryIP(x)
    Next x
    
End Sub

Public Sub SetEndIP(sIP As String)
    Dim aryIP() As String
    Dim x As Integer
    aryIP = Split(sIP, ".")
    For x = 0 To UBound(aryIP)
        If x >= 4 Then Exit For
        txtEndIP(x) = aryIP(x)
    Next x
    
End Sub

Public Function GetStartIP() As String
    Dim Cnt As Integer
    With txtStartIP
        For Cnt = 0 To .UBound
            GetStartIP = GetStartIP & .Item(Cnt).Text & "."
        Next Cnt
    End With
    
    GetStartIP = Left(GetStartIP, Len(GetStartIP) - 1)
End Function

Public Function GetEndIP() As String
    Dim Cnt As Integer
    With txtEndIP
        For Cnt = 0 To .UBound
            GetEndIP = GetEndIP & .Item(Cnt).Text & "."
        Next Cnt
    End With
    
    GetEndIP = Left(GetEndIP, Len(GetEndIP) - 1)
End Function

Private Sub PacketXCtrl1_OnPacket(ByVal pPacket As PacketXLibCtl.IPktXPacket)
    'Debug.Print pPacket.SourceIpAddress
    'Debug.Print pPacket.SourceMacAddress
    'Debug.Print CStr(pPacket.Data)
    'If pPacket.Data(16) = 8 Or pPacket.Data(17) = 8 Then
    '    Debug.Print pPacket.Protocol & " " & Hex(pPacket.Data(16)) & " " & Hex(pPacket.Data(17))
    'End If
    'Debug.Print pPacket.SourceMacAddress
    Dim sSourceIP As String
    
    Select Case pPacket.Protocol
        Case PktXProtocolTypeIP
        Case PktXProtocolTypeTCP
        Case PktXProtocolTypeEthernet
            If pPacket.Data(16) = 8 And pPacket.Data(17) = 0 Then
                With pPacket
                    'Debug.Print "SRC: " & .Data(28) & "." & .Data(29) & "." & .Data(30) & "." & .Data(31) & " to " & _
                    '            "DST: " & .Data(38) & "." & .Data(39) & "." & .Data(40) & "." & .Data(41)
                    'Debug.Print .SourceMacAddress & " to " & .DestMacAddress
                '
                    sSourceIP = .Data(28) & "." & .Data(29) & "." & .Data(30) & "." & .Data(31)
                End With
                If Not HostExists(sSourceIP, pPacket.SourceMacAddress) Then
                    Call AddHost(CStr(sSourceIP), CStr(pPacket.SourceMacAddress), "", 0)
                End If
            End If
            
        Case PktXProtocolTypeUDP
    End Select
End Sub

Function HostExists(IP As String, MAC As String) As Boolean
    Dim x As Integer
    HostExists = False
    For x = 0 To UBound(oHosts)
        If oHosts(x).IPAddress = IP And oHosts(x).MAC = MAC Then
            HostExists = True
            Exit For
        End If
    Next x
End Function

Public Sub AddHost(IP As String, MAC As String, Host As String, ResponseTime As Long)
    Dim iHost As Integer
    iHost = UBound(oHosts) + 1
    ReDim Preserve oHosts(iHost)
    Dim response_time As String
    
    
    oHosts(iHost).IPAddress = IP
    oHosts(iHost).MAC = MAC
    oHosts(iHost).HostName = Host
    oHosts(iHost).ResponseTime = ResponseTime
    
    
    Dim Item As ListItem
    Set Item = lvResults.ListItems.Add(, "|" & lvResults.ListItems.Count + 1, "", , 1)
    If Me.cmdPassive.Value Then
        response_time = "NA"
        If Me.chkResolveOnPassive.Value Then Host = Resolve.GetHostNameFromIP(IP)
        If Me.chkOnPassiveResponseTime.Value Then
            ResponseTime = oSweep.Scan(Me.cboAdapter.ListIndex, IP, MAC)
            response_time = ResponseTime & " ms"
        End If
    Else
        response_time = ResponseTime & " ms"
        If Me.chkResolveDuring.Value Then Host = Resolve.GetHostNameFromIP(IP)
    End If
    
    
    
    Dim sTime As String
    
    If Me.chkTimeOnly Then
        sTime = Time()
    Else
        sTime = Now()
    End If
    
    Item.SubItems(1) = IP
    Item.SubItems(3) = Host
    Item.SubItems(2) = MAC
    Item.SubItems(4) = response_time
    Item.SubItems(5) = sTime
    
    AltLVBackground Me.lvResults, Me.pic
End Sub

Private Sub AltLVBackground(lv As ListView, pic As PictureBox, _
    Optional ByVal StartAtOddRow As Boolean = False, _
        Optional ByVal AltBackColor As OLE_COLOR = -1)

Dim h               As Single
Dim sw              As Single
Dim oAltBackColor   As OLE_COLOR

'//If AltBackColor is not passed, default to picturebox backcolor
If AltBackColor = -1 Then
    oAltBackColor = pic.BackColor
Else
    oAltBackColor = AltBackColor
End If

With lv
    If .View = lvwReport Then
        If .ListItems.Count Then
            .PictureAlignment = lvwTile
            h = .ListItems(1).Height
            With pic
                .Visible = False
                .BackColor = lv.BackColor
                .BorderStyle = 0
                .Height = h * 2
                .Width = 10 * Screen.TwipsPerPixelX
                sw = .ScaleWidth
                .AutoRedraw = True
                If StartAtOddRow Then
                    pic.Line (0, 0)-Step(sw, h - Screen.TwipsPerPixelY), oAltBackColor, BF
                Else
                    pic.Line (0, h)-Step(sw, h), oAltBackColor, BF
                End If
                Set lv.Picture = .Image
                .AutoRedraw = False
                .BackColor = oAltBackColor
            End With
            .Refresh
            Exit Sub
        End If
    End If
    Set .Picture = Nothing
End With
    
End Sub
