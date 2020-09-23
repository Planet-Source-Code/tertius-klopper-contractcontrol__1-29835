Attribute VB_Name = "PicDeclerations"
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Type POINTAPI
    X As Long
    y As Long
    End Type
Public picc As Integer
Public Max As Integer
Public phi As Integer
Public lhdc As Long
Public b As Boolean
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

'Dim pimp As Integer, pump As Integer, foop As String, loopy As Boolean
'Dim bMoveFrom As Boolean, LastPoint As POINTAPI, Pause As Boolean
'Dim ResultRegion As Long, HolderRegion As Long, ObjectRegion As Long, nRet As Long
'Dim PolyPoints() As POINTAPI

Public hRgn As Long

Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_PATHMUSTEXIST = &H800
Public Const CC_FULLOPEN = &H2
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_RGBINIT = &H1
Public Const CC_ANYCOLOR = &H100
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
    Dim hRgn As Long, tRgn As Long
    Dim X As Integer, y As Integer, X0 As Integer
    Dim hdc As Long, BM As BITMAP
    hdc = CreateCompatibleDC(0)
    If hdc Then
        SelectObject hdc, cPicture
        GetObject cPicture, Len(BM), BM
        hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        For y = 0 To BM.bmHeight
            For X = 0 To BM.bmWidth
                While X <= BM.bmWidth And GetPixel(hdc, X, y) <> cTransparent
                    X = X + 1
                Wend
                X0 = X
                While X <= BM.bmWidth And GetPixel(hdc, X, y) = cTransparent
                    X = X + 1
                Wend
                If X0 < X Then
                    tRgn = CreateRectRgn(X0, y, X, y + 1)
                    CombineRgn hRgn, hRgn, tRgn, 4
                    DeleteObject tRgn
                End If
            Next X
        Next y
        GetBitmapRegion = hRgn
        DeleteObject SelectObject(hdc, cPicture)
    End If
    DeleteDC hdc
End Function

