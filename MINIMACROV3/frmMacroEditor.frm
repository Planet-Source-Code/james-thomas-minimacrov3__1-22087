VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMacroEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Macro Editor"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5640
   Icon            =   "frmMacroEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtMess 
      Height          =   1575
      Left            =   2760
      TabIndex        =   31
      Top             =   2880
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMacroEditor.frx":014A
   End
   Begin MSComctlLib.TreeView HistTree 
      Height          =   1815
      Left            =   2760
      TabIndex        =   30
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3201
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.TextBox txtHold 
      Height          =   285
      Left            =   1440
      TabIndex        =   29
      Text            =   "Text13"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtAscii 
      Height          =   285
      Left            =   1440
      TabIndex        =   28
      Text            =   "Text12"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtCap 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Text            =   "Text11"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtShift 
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Text            =   "Text10"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtCntl 
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Text            =   "Text9"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtAlt 
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Text            =   "Text8"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtLast_Evt 
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Text            =   "Text7"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtDrag 
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Text            =   "Text6"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtButton 
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtWait 
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Text            =   "Text4"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtEvt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Button Hold"
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Ascii Key No."
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Cap State"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Shift State"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Control State"
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Alt State"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Last Event No."
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Dragging"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Button Number"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Wait Time"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Y Coordinate"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "X Coordinate"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblEvent 
      Caption         =   "Event Number"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMacroEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim macarray() As mMacro
Dim EVT As Long
Dim Max_Evt As Long
Public Function GetNumbers(Exp As Variant) As Long
Dim i As Long
For i = 1 To Len(Exp)
    If IsNumeric(Mid(Exp, i, 1)) Then
        GetNumbers = GetNumbers & Val(Mid(Exp, i, 1))
    End If
Next i
End Function

Sub SetForm(evt_num)
On Error Resume Next
EVT = evt_num
txtEvt.Text = EVT
txtLast_Evt.Text = macarray(evt_num).last_evt
txtX.Text = macarray(evt_num).x
txtY.Text = macarray(evt_num).Y
txtWait.Text = macarray(evt_num).wait
txtButton.Text = macarray(evt_num).Button
txtDrag.Text = macarray(evt_num).dragging
txtAlt.Text = macarray(evt_num).AltState
txtCntl.Text = macarray(evt_num).CtrlState
txtShift.Text = macarray(evt_num).ShiftState
txtCap.Text = macarray(evt_num).CapState
txtAscii.Text = Chr(macarray(evt_num).AsciiKey)
txtHold.Text = macarray(evt_num).AsciiHold
End Sub

Private Sub Command2_Click()
If EVT <> Max_Evt Then
    EVT = EVT + 1
    SetForm (EVT)
End If

End Sub

Private Sub Command1_Click()
If EVT <> 0 Then
    EVT = EVT - 1
    SetForm (EVT)
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
CreateHistory
End Sub

Private Sub Form_Load()
Max_Evt = GetMacro(macarray())
SetForm (0)

End Sub

Public Sub CreateHistory()
Dim nod As Node
Dim kybd(0 To 255) As Byte
Dim i As Integer
Dim ret As Long
Dim ky As Long
Dim root As String
HistTree.Nodes.Clear
root = "r" & CStr(EVT)
i = EVT
Set nod = HistTree.Nodes.Add(, , root, "Step: " & i)
Do Until i > Max_Evt
    If root <> "r" & CStr(i) Then
        root = "r" & CStr(i)
        Set nod = HistTree.Nodes.Add(, , root, "Step: " & i)
    End If
    HistTree.Nodes.Add root, tvwChild, "X" & root, "X: " & macarray(i).x
    HistTree.Nodes.Add root, tvwChild, "Y" & root, "Y: " & macarray(i).Y
    HistTree.Nodes.Add root, tvwChild, "wait" & root, "wait: " & macarray(i).wait
    HistTree.Nodes.Add root, tvwChild, "Button" & root, "Button: " & macarray(i).Button
    HistTree.Nodes.Add root, tvwChild, "dragging" & root, "dragging: " & macarray(i).dragging
    HistTree.Nodes.Add root, tvwChild, "last_evt" & root, "last_evt: " & macarray(i).last_evt
    HistTree.Nodes.Add root, tvwChild, "AltState" & root, "AltState: " & macarray(i).AltState
    HistTree.Nodes.Add root, tvwChild, "CtrlState" & root, "CtrlState: " & macarray(i).CtrlState
    HistTree.Nodes.Add root, tvwChild, "ShiftState" & root, "ShiftState: " & macarray(i).ShiftState
    HistTree.Nodes.Add root, tvwChild, "CapState" & root, "CapState: " & macarray(i).CapState
    HistTree.Nodes.Add root, tvwChild, "AsciiKey" & root, "AsciiKey: " & macarray(i).AsciiKey
    HistTree.Nodes.Add root, tvwChild, "AsciiHold" & root, "AsciiHold: " & macarray(i).AsciiHold
    Dim j As Long
    If macarray(i).AsciiKey > 0 Then
        GetKeyboardState kybd(0)
            If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
                '===== Win95 or higher
                kybd(vbKeyShift) = macarray(i).ShiftState
                SetKeyboardState kybd(0)
                If macarray(i).AsciiKey <> vbKeyShift Then
                    ret = PostMessage(txtMess.hwnd, WM_KEYDOWN, macarray(i).AsciiKey, &H20002)
                End If
                kybd(vbKeyShift) = 0
                SetKeyboardState kybd(0)
            ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
                '===== WinNT
                'Simulate Key Press
                Call keybd_event(vbKeyShift, MapVirtualKey(VK_SHIFT, 0), KEYEVENTF_EXTENDEDKEY Or 0, 0&)
                Call PostMessage(txtMess.hwnd, WM_KEYDOWN, macarray(i).AsciiKey, 0)
                Call keybd_event(vbKeyShift, MapVirtualKey(vbKeyShift, 0), KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0&)
            End If

    End If
    i = i + 1
Loop
Exit Sub
End Sub

Private Sub HistTree_NodeClick(ByVal Node As MSComctlLib.Node)
SetForm (GetNumbers(Node.Text))

End Sub


