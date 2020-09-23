VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSerial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Serial Number"
   ClientHeight    =   4080
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordering Information"
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.Label lblOrder 
         Height          =   3255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3015
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraNumber 
      Caption         =   "Enter Valid Serial Number"
      Height          =   1095
      Left            =   3840
      TabIndex        =   6
      Top             =   2760
      Width           =   3735
      Begin MSMask.MaskEdBox mebN3 
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "####"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mebA1 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "????"
         Mask            =   "????"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mebN2 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "####"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mebN1 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Format          =   "####"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Line Line3 
         X1              =   2760
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   1920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   840
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"frmSerial.frx":0000
      Height          =   1695
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "www.technolord.bizland.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Click to go to the Web Site."
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim lMouse As Long

Private Sub CancelButton_Click()
frmTrial.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim strInstruct As String
lMouse = Screen.MousePointer
strInstruct = App.Title & " is a Product of TECHNOLORD Software."
strInstruct = strInstruct & vbCrLf & "To register Click on the Web Site in the center of this form."
strInstruct = strInstruct & vbCrLf & "After registering, you will be given an activation code to unlock the program."
lblOrder.Caption = App.Title & ":" & vbCrLf
lblOrder.Caption = lblOrder.Caption & "To place an Order, Click on the Web site in the center of this form."
lblOrder.Caption = lblOrder.Caption & "Email TECHNOLORD or Fill out the Registration Form. Your Serial Number will be sent to you via email."
lblOrder.Caption = lblOrder.Caption & vbCrLf
lblOrder.Caption = lblOrder.Caption & vbCrLf
lblOrder.Caption = lblOrder.Caption & App.LegalCopyright

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = lMouse
End Sub

Private Sub Label1_Click()
ExecuteLink (Label1.Caption)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbUpArrow

End Sub

Private Sub OKButton_Click()
Dim scode As String
scode = mebN1.Text & mebN2.Text & mebA1.Text & mebN3.Text
If ValidCode(scode) = True Then
    MsgBox "Thank you for registering with TECHNOLORD Software."
    frmEditor.Show
    Unload Me
Else
    MsgBox "You have entered an invalid serial number!" & vbCrLf & _
    "Register with TECHNOLORD Software or enter a valid Serial Number."
    mebN1.SetFocus
End If

End Sub
