VERSION 5.00
Begin VB.Form frmMata 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1920
   ClientLeft      =   8265
   ClientTop       =   4650
   ClientWidth     =   2700
   ControlBox      =   0   'False
   Icon            =   "frmMata.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2700
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1800
      Picture         =   "frmMata.frx":030A
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   1845
      Left            =   1330
      Shape           =   2  'Oval
      Top             =   5
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   1840
      Left            =   5
      Shape           =   2  'Oval
      Top             =   5
      Width           =   1260
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnusp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
'***************************************************************************
' Nama Program  : Program Bola Mata
' Deskripsi     : Contoh penggunaan fungsi Windows API untuk :
'                 1. Membentuk tampilan form sesuai dengan keinginan
'                 2. Mendeteksi posisi pointer mouse
'                 3. Membuat form selalu di atas jendela windows yang lain
'                 4. Menyimpan posisi program
' Programer     : K4m4r82
' Sumber        : http://coding4ever.wordpress.com
'***************************************************************************

Private Const RGN_COPY = 5
Private ResultRegion As Long

Dim X As Long, Y As Long
Dim lX As Long, lY As Long
Const pi = 3.14159265358979
Const FileSetting As String = "FileSetting.ini"

Private Type PosisiProgram
    Left As Single
    Top As Single
End Type

Dim Posisi As PosisiProgram

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)


    ReDim PolyPoints(0 To 75)
    For Counter = 0 To 75
        PolyPoints(Counter).X = GP0X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP0Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 76, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 75)
    For Counter = 0 To 75
        PolyPoints(Counter).X = GP1X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP1Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 76, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 37
    Case 1
        GP0X = 50
    Case 2
        GP0X = 51
    Case 3
        GP0X = 56
    Case 4
        GP0X = 57
    Case 5
        GP0X = 58
    Case 6
        GP0X = 60
    Case 7
        GP0X = 61
    Case 8
        GP0X = 63
    Case 9
        GP0X = 64
    Case 10
        GP0X = 69
    Case 11
        GP0X = 69
    Case 12
        GP0X = 73
    Case 13
        GP0X = 74
    Case 14
        GP0X = 76
    Case 15
        GP0X = 78
    Case 16
        GP0X = 79
    Case 17
        GP0X = 80
    Case 18
        GP0X = 82
    Case 19
        GP0X = 83
    Case 20
        GP0X = 84
    Case 21
        GP0X = 84
    Case 22
        GP0X = 83
    Case 23
        GP0X = 83
    Case 24
        GP0X = 82
    Case 25
        GP0X = 82
    Case 26
        GP0X = 81
    Case 27
        GP0X = 80
    Case 28
        GP0X = 79
    Case 29
        GP0X = 78
    Case 30
        GP0X = 75
    Case 31
        GP0X = 73
    Case 32
        GP0X = 73
    Case 33
        GP0X = 71
    Case 34
        GP0X = 71
    Case 35
        GP0X = 62
    Case 36
        GP0X = 61
    Case 37
        GP0X = 59
    Case 38
        GP0X = 52
    Case 39
        GP0X = 49
    Case 40
        GP0X = 36
    Case 41
        GP0X = 35
    Case 42
        GP0X = 33
    Case 43
        GP0X = 32
    Case 44
        GP0X = 27
    Case 45
        GP0X = 24
    Case 46
        GP0X = 23
    Case 47
        GP0X = 16
    Case 48
        GP0X = 16
    Case 49
        GP0X = 13
    Case 50
        GP0X = 12
    Case 51
        GP0X = 10
    Case 52
        GP0X = 8
    Case 53
        GP0X = 7
    Case 54
        GP0X = 6
    Case 55
        GP0X = 5
    Case 56
        GP0X = 4
    Case 57
        GP0X = 3
    Case 58
        GP0X = 2
    Case 59
        GP0X = 2
    Case 60
        GP0X = 3
    Case 61
        GP0X = 3
    Case 62
        GP0X = 4
    Case 63
        GP0X = 4
    Case 64
        GP0X = 5
    Case 65
        GP0X = 7
    Case 66
        GP0X = 8
    Case 67
        GP0X = 12
    Case 68
        GP0X = 14
    Case 69
        GP0X = 14
    Case 70
        GP0X = 19
    Case 71
        GP0X = 19
    Case 72
        GP0X = 20
    Case 73
        GP0X = 21
    Case 74
        GP0X = 25
    Case 75
        GP0X = 34
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 2
    Case 1
        GP0Y = 2
    Case 2
        GP0Y = 3
    Case 3
        GP0Y = 4
    Case 4
        GP0Y = 5
    Case 5
        GP0Y = 5
    Case 6
        GP0Y = 7
    Case 7
        GP0Y = 7
    Case 8
        GP0Y = 9
    Case 9
        GP0Y = 9
    Case 10
        GP0Y = 14
    Case 11
        GP0Y = 15
    Case 12
        GP0Y = 19
    Case 13
        GP0Y = 22
    Case 14
        GP0Y = 24
    Case 15
        GP0Y = 29
    Case 16
        GP0Y = 32
    Case 17
        GP0Y = 35
    Case 18
        GP0Y = 40
    Case 19
        GP0Y = 47
    Case 20
        GP0Y = 50
    Case 21
        GP0Y = 75
    Case 22
        GP0Y = 76
    Case 23
        GP0Y = 81
    Case 24
        GP0Y = 82
    Case 25
        GP0Y = 85
    Case 26
        GP0Y = 86
    Case 27
        GP0Y = 91
    Case 28
        GP0Y = 92
    Case 29
        GP0Y = 95
    Case 30
        GP0Y = 102
    Case 31
        GP0Y = 104
    Case 32
        GP0Y = 105
    Case 33
        GP0Y = 107
    Case 34
        GP0Y = 108
    Case 35
        GP0Y = 117
    Case 36
        GP0Y = 117
    Case 37
        GP0Y = 119
    Case 38
        GP0Y = 122
    Case 39
        GP0Y = 123
    Case 40
        GP0Y = 123
    Case 41
        GP0Y = 122
    Case 42
        GP0Y = 122
    Case 43
        GP0Y = 121
    Case 44
        GP0Y = 119
    Case 45
        GP0Y = 116
    Case 46
        GP0Y = 116
    Case 47
        GP0Y = 109
    Case 48
        GP0Y = 108
    Case 49
        GP0Y = 105
    Case 50
        GP0Y = 102
    Case 51
        GP0Y = 100
    Case 52
        GP0Y = 95
    Case 53
        GP0Y = 92
    Case 54
        GP0Y = 89
    Case 55
        GP0Y = 86
    Case 56
        GP0Y = 83
    Case 57
        GP0Y = 78
    Case 58
        GP0Y = 71
    Case 59
        GP0Y = 50
    Case 60
        GP0Y = 49
    Case 61
        GP0Y = 44
    Case 62
        GP0Y = 43
    Case 63
        GP0Y = 40
    Case 64
        GP0Y = 39
    Case 65
        GP0Y = 31
    Case 66
        GP0Y = 30
    Case 67
        GP0Y = 21
    Case 68
        GP0Y = 19
    Case 69
        GP0Y = 18
    Case 70
        GP0Y = 13
    Case 71
        GP0Y = 12
    Case 72
        GP0Y = 11
    Case 73
        GP0Y = 11
    Case 74
        GP0Y = 7
    Case 75
        GP0Y = 3
    End Select
End Function
Private Function GP1X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP1X = 126
    Case 1
        GP1X = 139
    Case 2
        GP1X = 140
    Case 3
        GP1X = 145
    Case 4
        GP1X = 146
    Case 5
        GP1X = 147
    Case 6
        GP1X = 149
    Case 7
        GP1X = 150
    Case 8
        GP1X = 152
    Case 9
        GP1X = 153
    Case 10
        GP1X = 158
    Case 11
        GP1X = 158
    Case 12
        GP1X = 162
    Case 13
        GP1X = 163
    Case 14
        GP1X = 165
    Case 15
        GP1X = 167
    Case 16
        GP1X = 168
    Case 17
        GP1X = 169
    Case 18
        GP1X = 171
    Case 19
        GP1X = 172
    Case 20
        GP1X = 173
    Case 21
        GP1X = 173
    Case 22
        GP1X = 172
    Case 23
        GP1X = 172
    Case 24
        GP1X = 171
    Case 25
        GP1X = 171
    Case 26
        GP1X = 170
    Case 27
        GP1X = 169
    Case 28
        GP1X = 168
    Case 29
        GP1X = 167
    Case 30
        GP1X = 164
    Case 31
        GP1X = 162
    Case 32
        GP1X = 162
    Case 33
        GP1X = 160
    Case 34
        GP1X = 160
    Case 35
        GP1X = 151
    Case 36
        GP1X = 150
    Case 37
        GP1X = 148
    Case 38
        GP1X = 141
    Case 39
        GP1X = 138
    Case 40
        GP1X = 125
    Case 41
        GP1X = 124
    Case 42
        GP1X = 122
    Case 43
        GP1X = 121
    Case 44
        GP1X = 116
    Case 45
        GP1X = 113
    Case 46
        GP1X = 112
    Case 47
        GP1X = 105
    Case 48
        GP1X = 105
    Case 49
        GP1X = 102
    Case 50
        GP1X = 101
    Case 51
        GP1X = 99
    Case 52
        GP1X = 97
    Case 53
        GP1X = 96
    Case 54
        GP1X = 95
    Case 55
        GP1X = 94
    Case 56
        GP1X = 93
    Case 57
        GP1X = 92
    Case 58
        GP1X = 91
    Case 59
        GP1X = 91
    Case 60
        GP1X = 92
    Case 61
        GP1X = 92
    Case 62
        GP1X = 93
    Case 63
        GP1X = 93
    Case 64
        GP1X = 94
    Case 65
        GP1X = 96
    Case 66
        GP1X = 97
    Case 67
        GP1X = 101
    Case 68
        GP1X = 103
    Case 69
        GP1X = 103
    Case 70
        GP1X = 108
    Case 71
        GP1X = 108
    Case 72
        GP1X = 109
    Case 73
        GP1X = 110
    Case 74
        GP1X = 114
    Case 75
        GP1X = 123
    End Select
End Function
Private Function GP1Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP1Y = 2
    Case 1
        GP1Y = 2
    Case 2
        GP1Y = 3
    Case 3
        GP1Y = 4
    Case 4
        GP1Y = 5
    Case 5
        GP1Y = 5
    Case 6
        GP1Y = 7
    Case 7
        GP1Y = 7
    Case 8
        GP1Y = 9
    Case 9
        GP1Y = 9
    Case 10
        GP1Y = 14
    Case 11
        GP1Y = 15
    Case 12
        GP1Y = 19
    Case 13
        GP1Y = 22
    Case 14
        GP1Y = 24
    Case 15
        GP1Y = 29
    Case 16
        GP1Y = 32
    Case 17
        GP1Y = 35
    Case 18
        GP1Y = 40
    Case 19
        GP1Y = 47
    Case 20
        GP1Y = 50
    Case 21
        GP1Y = 75
    Case 22
        GP1Y = 76
    Case 23
        GP1Y = 81
    Case 24
        GP1Y = 82
    Case 25
        GP1Y = 85
    Case 26
        GP1Y = 86
    Case 27
        GP1Y = 91
    Case 28
        GP1Y = 92
    Case 29
        GP1Y = 95
    Case 30
        GP1Y = 102
    Case 31
        GP1Y = 104
    Case 32
        GP1Y = 105
    Case 33
        GP1Y = 107
    Case 34
        GP1Y = 108
    Case 35
        GP1Y = 117
    Case 36
        GP1Y = 117
    Case 37
        GP1Y = 119
    Case 38
        GP1Y = 122
    Case 39
        GP1Y = 123
    Case 40
        GP1Y = 123
    Case 41
        GP1Y = 122
    Case 42
        GP1Y = 122
    Case 43
        GP1Y = 121
    Case 44
        GP1Y = 119
    Case 45
        GP1Y = 116
    Case 46
        GP1Y = 116
    Case 47
        GP1Y = 109
    Case 48
        GP1Y = 108
    Case 49
        GP1Y = 105
    Case 50
        GP1Y = 102
    Case 51
        GP1Y = 100
    Case 52
        GP1Y = 95
    Case 53
        GP1Y = 92
    Case 54
        GP1Y = 89
    Case 55
        GP1Y = 86
    Case 56
        GP1Y = 83
    Case 57
        GP1Y = 78
    Case 58
        GP1Y = 71
    Case 59
        GP1Y = 50
    Case 60
        GP1Y = 49
    Case 61
        GP1Y = 44
    Case 62
        GP1Y = 43
    Case 63
        GP1Y = 40
    Case 64
        GP1Y = 39
    Case 65
        GP1Y = 31
    Case 66
        GP1Y = 30
    Case 67
        GP1Y = 21
    Case 68
        GP1Y = 19
    Case 69
        GP1Y = 18
    Case 70
        GP1Y = 13
    Case 71
        GP1Y = 12
    Case 72
        GP1Y = 11
    Case 73
        GP1Y = 11
    Case 74
        GP1Y = 7
    Case 75
        GP1Y = 3
    End Select
End Function


Sub PaintEye(mx As Long, my As Long, x1 As Long, y1 As Long, radius As Integer, pictureeye As Object)
    'On Error Resume Next

    Dim X2, Y2 As Long
    Dim MeLeft As Long, MeTop As Long
    Dim lkat1 As Long, lkat2 As Long
    Dim v As Double
    
    MeLeft = (Me.Left / Screen.TwipsPerPixelX)
    MeTop = (Me.Top / Screen.TwipsPerPixelY)
    ' this is the "calc"
    If my > (y1 + MeTop) Then
        lkat1 = mx - (x1 + MeLeft)
        lkat2 = my - (y1 + MeTop)
        v = Atn(lkat1 / lkat2)
        
    ElseIf my < (y1 + MeTop) Then
        lkat1 = mx - (x1 + MeLeft)
        lkat2 = my - (y1 + MeTop)
        v = Atn(lkat1 / lkat2) + 3.15
        
    ElseIf my = (y1 + MeTop) Then
        my = my + 1
        lkat1 = mx - (x1 + MeLeft)
        lkat2 = my - (y1 + MeTop)
        v = Atn(lkat1 / lkat2)
    End If
    
    X2 = (Sin(v) * radius) + x1
    Y2 = (Cos(v) * radius) + y1
    
    Call Paint_TransPictureAt2(pictureeye, X2 - (pictureeye.ScaleWidth / 2), Y2 - pictureeye.ScaleHeight / 2)
End Sub

Sub Get_Cursor_Pos()
    X = GetX
    Y = GetY
End Sub

Sub Paint_TransPictureAt2(picbox, xval, yval)
    Call BitBlt(Me.hDC, xval, yval, picbox.ScaleWidth, picbox.ScaleHeight, picbox.hDC, 0, 0, SRCAND)
End Sub

Sub Paint_TransPictureAt3(picbox, xval, yval)
    Call BitBlt(Me.hDC, xval, yval, picbox.ScaleWidth, picbox.ScaleHeight, picbox.hDC, 0, 0, SRCINVERT)
End Sub

Private Sub BacaPosisi()
    Posisi.Left = ReadINI("Posisi", "Left", App.Path & "\" & FileSetting)
    Posisi.Top = ReadINI("Posisi", "Top", App.Path & "\" & FileSetting)
End Sub

Private Sub SimpanPosisi(ByVal vLeft As String, ByVal vTop As String)
    Call WriteINI("Posisi", "Left", vLeft, App.Path & "\" & FileSetting)
    Call WriteINI("Posisi", "Top", vTop, App.Path & "\" & FileSetting)
End Sub

Private Sub Form_Activate()
    Do
        Sleep 1
        DoEvents
        lX = X
        lY = Y
        Get_Cursor_Pos
        If lX <> X Or lY <> Y Then
            Me.Cls
            Call PaintEye(X, Y, LEye.E_x, LEye.E_y, LEye.E_R, Picture1) 'mata kiri
            Call PaintEye(X, Y, REye.E_x, REye.E_y, REye.E_R, Picture1) 'mata kanan
        End If
    Loop

End Sub

Private Sub Form_Initialize()
    LEye.E_R = 27.5 'radius
    LEye.E_x = 40
    LEye.E_y = 65
    
    REye.E_R = 27.5 '18
    REye.E_x = 132
    REye.E_y = 65
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
    
    Call FormTopMost(Me.hWnd)
    
    Call BacaPosisi
    
    Me.Left = Posisi.Left
    Me.Top = Posisi.Top
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuFile
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    Call SimpanPosisi(Screen.ActiveForm.Left, Screen.ActiveForm.Top)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Call FormNoTopMost(Me.hWnd)
    Call SimpanPosisi(Screen.ActiveForm.Left, Screen.ActiveForm.Top)
    End
End Sub

