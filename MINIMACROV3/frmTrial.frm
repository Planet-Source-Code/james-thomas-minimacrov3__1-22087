VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notice"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmTrial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLabel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   255
      ScaleWidth      =   1695
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trial Version"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4095
      Begin VB.Label lblStats 
         Height          =   855
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "Validate"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar prgbrMeter 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblTrialMeter 
      Caption         =   "Trial Usage Meter"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    On Error GoTo fin

Unload Me
End


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub cmdOk_Click()

    On Error GoTo fin

Load frmEditor
Unload Me



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub cmdValidate_Click()

    On Error GoTo fin

frmSerial.Show
Unload Me


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub Form_Activate()

    On Error GoTo fin

Dim WindowRegion As Long
Dim strMsg As String
picLabel.ScaleMode = vbPixels
picLabel.AutoRedraw = True
picLabel.AutoSize = True
picLabel.BorderStyle = vbBSNone
strMsg = Max_Times_Loaded & " Day Trial Version"
picLabel.Print strMsg
Call AutoFormShape(picLabel)



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub

Private Sub Form_Load()

    On Error GoTo fin

lblStats.Caption = "You are using a trial Version of " & App.Title & vbCrLf

If Expired = True Then
    lblStats.Caption = lblStats.Caption & App.Title & " has Expired. Click Validate for Ordering Information."
    cmdOk.Enabled = False
Else
    lblStats.Caption = lblStats.Caption & "It has been use for " & TimesLoaded & " Days since first run on " & FirstRun
    cmdOk.Enabled = True
End If
prgbrMeter.Max = Max_Times_Loaded
prgbrMeter.Min = 0
prgbrMeter.Value = Max_Times_Loaded - TimesLoaded



Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Sub

