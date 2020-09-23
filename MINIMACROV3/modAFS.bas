Attribute VB_Name = "modAFS"

Public Const WS_EX_TRANSPARENT = &H20&
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const flags = SWP_NOMOVE + SWP_NOSIZE
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&     ' This setting is not in your API viewer, not sure why.
                                        ' If you use SC_MOVE then the mouse moves to the title bar
                                        ' and then moves the form, which makes forms with no title bar
                                        ' to not work.
Public Const WM_SYSCOMMAND = &H112

Dim CurRgn, TempRgn As Long  ' Region variables

Public Function AutoFormShape(bg As PictureBox)
Dim transColor As Long
Dim x, y As Integer
transColor = GetPixel(bg.hDC, x, y)
CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)  ' Create base region which is the current whole window
While y <= bg.ScaleHeight  ' Go through each column of pixels on form
    While x <= bg.ScaleWidth  ' Go through each line of pixels on form
        If GetPixel(bg.hDC, x, y) = transColor Then  ' If the pixels color is the transparency color (bright purple is a good one to use)
            TempRgn = CreateRectRgn(x, y, x + 1, y + 1)  ' Create a temporary pixel region for this pixel
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
            DeleteObject (TempRgn)  ' Delete the temporary pixel region and clear up very important resources
        End If
        x = x + 1
    Wend
        y = y + 1
        x = 0
Wend
success = SetWindowRgn(bg.hwnd, CurRgn, True)  ' Finally set the windows region to the final product
DeleteObject (CurRgn)  ' Delete the now un-needed base region and free resources

End Function

Public Function SetTransparent(hwndHandle As Long) As Boolean
SetWindowLong hwndHandle, GWL_EXSTYLE, WS_EX_TRANSPARENT
SetTransparent = True
End Function

Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim x As Long, y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1
            
            If GetPixel(hDC, x, y) = TransparentColor Or x = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function

