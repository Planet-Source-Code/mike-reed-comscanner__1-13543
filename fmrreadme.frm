VERSION 5.00
Begin VB.Form fmrreadme 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3240
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5088
   Icon            =   "fmrreadme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5088
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox LstRead 
      BackColor       =   &H80000004&
      Height          =   2352
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdreadmeok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "fmrreadme"
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
'   * fmrreadme is the veiw the instuction.                                  *
'   *************************************************************************

Private Sub cmdreadmeok_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim ReadMeFileName As String
Dim fileNum As Integer
Dim TempName As String
 ChDrive App.Path
 ChDir App.Path
On Error GoTo ReadMe1
 ReadMeFileName = "Read_Me.txt"
 fileNum = FreeFile
 LstRead.Clear
Open ReadMeFileName For Input As fileNum
  Do Until EOF(fileNum)
   Input #fileNum, TempName
   LstRead.AddItem TempName
   DoEvents
  Loop
  Close fileNum
  Exit Sub
ReadMe1:
 MsgBox Error$, vbInformation, "Unable to load Read Me File"
End Sub
