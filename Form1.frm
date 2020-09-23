VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Scanner Testing"
   ClientHeight    =   2220
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   5124
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5124
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1560
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2200
      Left            =   840
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   240
      Top             =   1440
   End
   Begin VB.TextBox txtinput 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   1584
      ItemData        =   "Form1.frx":0ECA
      Left            =   2760
      List            =   "Form1.frx":0ECC
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2175
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   1968
      Width           =   5124
      _ExtentX        =   9038
      _ExtentY        =   445
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   529
            MinWidth        =   529
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   529
            MinWidth        =   529
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   529
            MinWidth        =   529
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5362
            MinWidth        =   5362
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   1024
      RThreshold      =   13
      RTSEnable       =   -1  'True
      ParitySetting   =   2
      SThreshold      =   2
   End
   Begin VB.Label Label2 
      Caption         =   "History scanned:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "What you just scanned:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Menu mnuport 
      Caption         =   "Port"
      Begin VB.Menu mnucom1 
         Caption         =   "COM1"
      End
      Begin VB.Menu mnucom2 
         Caption         =   "COM2"
      End
      Begin VB.Menu mnucom3 
         Caption         =   "COM3"
      End
      Begin VB.Menu mnucom4 
         Caption         =   "COM4"
      End
   End
   Begin VB.Menu mnuSpeed 
      Caption         =   "Speed"
      Begin VB.Menu mnu28800 
         Caption         =   "28800"
      End
      Begin VB.Menu mnu19200 
         Caption         =   "19200"
      End
      Begin VB.Menu mnu14400 
         Caption         =   "14400"
      End
      Begin VB.Menu mnu9600 
         Caption         =   "9600"
      End
      Begin VB.Menu mnu4800 
         Caption         =   "4800"
      End
      Begin VB.Menu mnu2400 
         Caption         =   "2400"
      End
      Begin VB.Menu mnu1200 
         Caption         =   "1200"
      End
      Begin VB.Menu mnuparity 
         Caption         =   "Parity"
         Begin VB.Menu mnun 
            Caption         =   "None"
         End
         Begin VB.Menu mnuodd 
            Caption         =   "Odd"
         End
         Begin VB.Menu mnueven 
            Caption         =   "Even"
         End
         Begin VB.Menu mnumark 
            Caption         =   "Mark"
         End
         Begin VB.Menu mnuspace 
            Caption         =   "Space"
         End
      End
      Begin VB.Menu mnudata 
         Caption         =   "Data"
         Begin VB.Menu mnu7 
            Caption         =   "7"
         End
         Begin VB.Menu mnu8 
            Caption         =   "8"
         End
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop Bit"
         Begin VB.Menu mnu1 
            Caption         =   "1"
         End
         Begin VB.Menu mnu2 
            Caption         =   "2"
         End
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear Display"
   End
   Begin VB.Menu mnuopen 
      Caption         =   "Open Port"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Version 1.10A"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   *************************************************************************
'   *       This program designed by Troy M. Reed for test scanning a       *
'   *       hand held serial scanner.                                       *
'   *************************************************************************
'   *                         Design Date:  04/03/00                        *
'   *************************************************************************
'   * frmMain is the veiw screen for user.                                  *
'   *************************************************************************


Private Sub Form_Load()
  
  MsgBox "Choose your setting and press Open Port before scanning.", vbCritical, App.ProductName
  With StatusBar1
       .Panels(1).Text = "COM1"
       .Panels(2).Text = "9600"
       .Panels(3).Text = "E"
       .Panels(4).Text = "8"
       .Panels(5).Text = "1"
       .Panels(6).Text = "Press Open Port to use setting"
  End With
End Sub

Private Sub mnu1_Click()
  mnuopen.Enabled = True
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(5).Text = "1"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu1200_Click()

  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "1200"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu14400_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "14400"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu19200_Click()

  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "19200"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu2_Click()

  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(5).Text = "2"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu2400_Click()
 
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "2400"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu28800_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "28800"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu4800_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "4800"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu7_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(4).Text = "7"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu8_Click()

  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(4).Text = "8"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnu9600_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(2).Text = "9600"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnuClear_Click()
  List1.Clear
End Sub

Private Sub mnucom1_Click()
  If MSComm1.PortOpen = True Then
     MSComm1.PortOpen = False
  End If
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(1).Text = "COM1"
  MSComm1.CommPort = 1
End Sub

Private Sub mnucom2_Click()
  If MSComm1.PortOpen = True Then
     MSComm1.PortOpen = False
  End If
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(1).Text = "COM2"
  MSComm1.CommPort = 2
End Sub

Private Sub mnucom3_Click()
  If MSComm1.PortOpen = True Then
     MSComm1.PortOpen = False
  End If
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(1).Text = "COM3"
  MSComm1.CommPort = 3
End Sub

Private Sub mnucom4_Click()
  If MSComm1.PortOpen = True Then
     MSComm1.PortOpen = False
  End If
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(1).Text = "COM4"
  MSComm1.CommPort = 4
End Sub

Private Sub mnueven_Click()
 
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(3).Text = "E"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnuexit_Click()
  Unload Me
  End
End Sub

Private Sub mnumark_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(3).Text = "M"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnun_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(3).Text = "N"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnuodd_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(3).Text = "O"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnuopen_Click()

  If MSComm1.PortOpen = True Then
     MSComm1.PortOpen = False
  End If
  StatusBar1.Panels(6).Text = "Ready for scanner input."
  MSComm1.PortOpen = True
  
End Sub
Private Sub mnuspace_Click()
  
  StatusBar1.Panels(6).Text = "Press Open Port to use setting."
  StatusBar1.Panels(3).Text = "S"
  CheckMySettings = StatusBar1.Panels(2).Text + "," + StatusBar1.Panels(3).Text + "," + StatusBar1.Panels(4).Text + "," + StatusBar1.Panels(5).Text
  MSComm1.Settings = CheckMySettings
End Sub

Private Sub mnuVer_Click()

 Dim j As Integer
 fmrreadme.Caption = "Programmed by Mike Reed."
  Timer3.Enabled = True
  For j = 1 To 6
    FlashWindow Me.hwnd, 1
    Sleep 100
  Next j
  fmrreadme.Show
End Sub


Private Sub MSComm1_OnComm()
Dim CheckMyScan
Dim CheckForCR
Dim CountMe As Integer
Dim CountMy As Integer
Dim Number As Integer
On Error GoTo Mscomm11:
  If MSComm1.CommEvent = 2 Then
     txtinput.Text = MSComm1.Input
     CheckMyScan = txtinput.Text
     CountMe = Len(txtinput.Text)
     Number = 0
       Do Until CountMy = CountMe
            txtinput.SelStart = Number
            txtinput.SelLength = Len(txtinput.Text)
             CheckForCR = txtinput.SelText
             If CheckForCR = vbCr Then
                StatusBar1.Panels(6).Text = "Scanner is programed with CR."
                CheckMyScan = txtinput.Text
                List1.AddItem (CheckMyScan)
                Timer2.Enabled = True
                MSComm1.PortOpen = False
                MSComm1.PortOpen = True
                Exit Sub
             End If
             DoEvents
             Number = Number + 1
             CountMy = CountMy + 1
       Loop
     List1.AddItem (CheckMyScan)
     Timer1.Enabled = True
     MSComm1.PortOpen = False
     MSComm1.PortOpen = True
  End If
  Exit Sub
Mscomm11:
  MsgBox "A error in reading this bar code", vbOKOnly, App.ProductName
End Sub

Private Sub Timer1_Timer()
  StatusBar1.Panels(6).Text = "Ready for scanner input."
  Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
  StatusBar1.Panels(6).Text = "Ready for scanner input."
  Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
 Timer3.Enabled = False
End Sub
