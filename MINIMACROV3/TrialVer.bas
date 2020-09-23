Attribute VB_Name = "TrialVer"
Option Explicit
Global TrialVers As Boolean
Global Expired As Boolean
Global FirstRun As String
Global TimesLoaded As Integer

Public Const Max_Times_Loaded = 60
Public Const N1_Min = 1000
Public Const N1_Max = 1050
Public Const N2_Min = 3126
Public Const N2_Max = 3135
Public Const A1_Min = 65
Public Const A1_Max = 68
Public Function Crypt(Text As String) As String
On Error GoTo fin
Dim strTempChar As String, i As Integer
For i = 1 To Len(Text)
    If Asc(Mid$(Text, i, 1)) < 128 Then
      strTempChar = Asc(Mid$(Text, i, 1)) + 128
    ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
      strTempChar = Asc(Mid$(Text, i, 1)) - 128
    End If
    Mid$(Text, i, 1) = Chr(strTempChar)
Next i
Crypt = Text

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select
End Function

Public Sub Main()

    On Error GoTo fin

Dim Pass_Stored As String
'Get the stored registration key
Pass_Stored = GetSetting(App.ProductName, "Settings", "RegKey", "")

'See if it matches the valid registration key
If Pass_Stored <> "" And ValidCode(Pass_Stored) = True Then
    'If so, bypass Trial Version window
    'and go directly to the application
    TrialVers = False
    frmEditor.Show
    Exit Sub
End If

'See when the app was first run
FirstRun = GetSetting(App.ProductName, "Settings", "FirstRun", "")
'If it hasn't been run before, save today's date
'as the date it was first run
If FirstRun = "" Then
    FirstRun = Format(Date, "MM/DD/YYYY")
    SaveSetting App.ProductName, "Settings", "FirstRun", FirstRun
End If
If CVDate(FirstRun) < CVDate(Format(Date, "mm/dd/yyyy")) Then
    'See how many Days the app has been loaded
    TimesLoaded = DateDiff("d", CVDate(FirstRun), CVDate(Format(Date, "mm/dd/yyyy")))
End If
'If it hasn't been loaded before, save it as being
'loaded 1 time
SaveSetting App.ProductName, "Settings", "TimesLoaded", TimesLoaded

'If it has been loaded more than the times it can be,
'tell the application that the trial version has expired
If TimesLoaded > Max_Times_Loaded Then
    Expired = True
End If

'Set the Trial Version to True so that the
'application knows it's a trial version
TrialVers = True
'Show the Trial Version window
frmTrial.Show


Exit Sub
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select


End Sub
Public Function ValidCode(scode As String) As Boolean

    On Error GoTo fin

Dim iA1Indx As Integer
Dim bLetters As Boolean
Dim bValid As Boolean
Dim N1 As String, N2 As String, A1 As String, N3 As String
bValid = False
If Len(scode) < 16 Then
    ValidCode = False
    Exit Function
End If
N1 = Mid(scode, 1, 4)
N2 = Mid(scode, 5, 4)
A1 = Mid(scode, 9, 4)
N3 = Mid(scode, 13, 4)
If Val(N1) >= N1_Min And Val(N1) <= N1_Max Then
    'part 1
    If Val(N2) >= N2_Min And Val(N2) <= N2_Max Then
        'part 2
        For iA1Indx = 1 To Len(A1)
            'part 3
            If Asc(Mid(UCase(A1), iA1Indx, 1)) >= A1_Min And Asc(Mid(UCase(A1), iA1Indx, 1)) <= A1_Max Then
                bLetters = True
            Else
                bLetters = False
            End If
        Next iA1Indx
        If bLetters = True And CLng(N3) = 700 Then
            'part 4 everthing is valid
            bValid = True
            Call SaveSetting(App.ProductName, "Settings", "RegKey", scode)
        End If
    End If
End If
ValidCode = bValid
Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function
Public Function WriteValidCode(FN As String)

    On Error GoTo fin

'This is a 4 tier code system
'
Dim scode As String
Dim N1 As Long, N2 As Long, A1 As String, N3 As Long
Dim A1_1 As Integer, A1_2 As Integer, A1_3 As Integer, A1_4 As Integer
Dim iFNum As Integer
iFNum = FreeFile()
N3 = 700
Open FN For Output As iFNum
    For N1 = N1_Min To N1_Max
        For N2 = N2_Min To N2_Max
            For A1_1 = A1_Min To A1_Max
                For A1_2 = A1_Min To A1_Max
                    For A1_3 = A1_Min To A1_Max
                        For A1_4 = A1_Min To A1_Max
                            A1 = Chr(A1_1) & Chr(A1_2) & Chr(A1_3) & Chr(A1_4)
                            scode = N1 & N2 & A1 & "0" & N3
                            Write #iFNum, scode
                            DoEvents
                        Next A1_4
                    Next A1_3
                Next A1_2
            Next A1_1
        Next N2
        Debug.Print N1
    Next N1
Close iFNum

Exit Function
fin:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " " & Err.Description
End Select

End Function


