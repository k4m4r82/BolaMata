VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   4155
   ClientLeft      =   7005
   ClientTop       =   3915
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5175
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2220
      TabIndex        =   0
      Top             =   3300
      Width           =   735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RGN_COPY = 5
Private ResultRegion As Long

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

    ReDim PolyPoints(0 To 23)
    For Counter = 0 To 23
        PolyPoints(Counter).X = GP0X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP0Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 24, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 6)
    For Counter = 0 To 6
        PolyPoints(Counter).X = GP1X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP1Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 7, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 41)
    For Counter = 0 To 41
        PolyPoints(Counter).X = GP2X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP2Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 42, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 77)
    For Counter = 0 To 77
        PolyPoints(Counter).X = GP3X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP3Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 78, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 10)
    For Counter = 0 To 10
        PolyPoints(Counter).X = GP4X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP4Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 11, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 31)
    For Counter = 0 To 31
        PolyPoints(Counter).X = GP5X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP5Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 32, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 40)
    For Counter = 0 To 40
        PolyPoints(Counter).X = GP6X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP6Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 41, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 44)
    For Counter = 0 To 44
        PolyPoints(Counter).X = GP7X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP7Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 45, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 13)
    For Counter = 0 To 13
        PolyPoints(Counter).X = GP8X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP8Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 14, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 9)
    For Counter = 0 To 9
        PolyPoints(Counter).X = GP9X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP9Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 10, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 7)
    For Counter = 0 To 7
        PolyPoints(Counter).X = GP10X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP10Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 8, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 8)
    For Counter = 0 To 8
        PolyPoints(Counter).X = GP11X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP11Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 9, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 9)
    For Counter = 0 To 9
        PolyPoints(Counter).X = GP12X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP12Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 10, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 27)
    For Counter = 0 To 27
        PolyPoints(Counter).X = GP13X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP13Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 28, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 45)
    For Counter = 0 To 45
        PolyPoints(Counter).X = GP14X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP14Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 46, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 16)
    For Counter = 0 To 16
        PolyPoints(Counter).X = GP15X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP15Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 17, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 45)
    For Counter = 0 To 45
        PolyPoints(Counter).X = GP16X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP16Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 46, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 31)
    For Counter = 0 To 31
        PolyPoints(Counter).X = GP17X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP17Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 32, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 31)
    For Counter = 0 To 31
        PolyPoints(Counter).X = GP18X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP18Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 32, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 47)
    For Counter = 0 To 47
        PolyPoints(Counter).X = GP19X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP19Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 48, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 46)
    For Counter = 0 To 46
        PolyPoints(Counter).X = GP20X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP20Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 47, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 7)
    For Counter = 0 To 7
        PolyPoints(Counter).X = GP21X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP21Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 8, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 40)
    For Counter = 0 To 40
        PolyPoints(Counter).X = GP22X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP22Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 41, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 16)
    For Counter = 0 To 16
        PolyPoints(Counter).X = GP23X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP23Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 17, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 8)
    For Counter = 0 To 8
        PolyPoints(Counter).X = GP24X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP24Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 9, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 9)
    For Counter = 0 To 9
        PolyPoints(Counter).X = GP25X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP25Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 10, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 10)
    For Counter = 0 To 10
        PolyPoints(Counter).X = GP26X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP26Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 11, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    
    ReDim PolyPoints(0 To 24)
    For Counter = 0 To 24
        PolyPoints(Counter).X = GP27X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP27Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 25, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 3)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 33
    Case 1
        GP0X = 59
    Case 2
        GP0X = 60
    Case 3
        GP0X = 64
    Case 4
        GP0X = 65
    Case 5
        GP0X = 66
    Case 6
        GP0X = 71
    Case 7
        GP0X = 72
    Case 8
        GP0X = 71
    Case 9
        GP0X = 71
    Case 10
        GP0X = 67
    Case 11
        GP0X = 67
    Case 12
        GP0X = 68
    Case 13
        GP0X = 69
    Case 14
        GP0X = 73
    Case 15
        GP0X = 74
    Case 16
        GP0X = 74
    Case 17
        GP0X = 73
    Case 18
        GP0X = 73
    Case 19
        GP0X = 67
    Case 20
        GP0X = 64
    Case 21
        GP0X = 61
    Case 22
        GP0X = 34
    Case 23
        GP0X = 33
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 3
    Case 1
        GP0Y = 3
    Case 2
        GP0Y = 4
    Case 3
        GP0Y = 4
    Case 4
        GP0Y = 5
    Case 5
        GP0Y = 5
    Case 6
        GP0Y = 10
    Case 7
        GP0Y = 15
    Case 8
        GP0Y = 18
    Case 9
        GP0Y = 19
    Case 10
        GP0Y = 23
    Case 11
        GP0Y = 24
    Case 12
        GP0Y = 25
    Case 13
        GP0Y = 25
    Case 14
        GP0Y = 29
    Case 15
        GP0Y = 34
    Case 16
        GP0Y = 37
    Case 17
        GP0Y = 38
    Case 18
        GP0Y = 40
    Case 19
        GP0Y = 46
    Case 20
        GP0Y = 47
    Case 21
        GP0Y = 48
    Case 22
        GP0Y = 48
    Case 23
        GP0Y = 47
    End Select
End Function
Private Function GP1X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP1X = 120
    Case 1
        GP1X = 129
    Case 2
        GP1X = 130
    Case 3
        GP1X = 130
    Case 4
        GP1X = 129
    Case 5
        GP1X = 121
    Case 6
        GP1X = 120
    End Select
End Function
Private Function GP1Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP1Y = 3
    Case 1
        GP1Y = 3
    Case 2
        GP1Y = 4
    Case 3
        GP1Y = 47
    Case 4
        GP1Y = 48
    Case 5
        GP1Y = 48
    Case 6
        GP1Y = 47
    End Select
End Function
Private Function GP2X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP2X = 192
    Case 1
        GP2X = 207
    Case 2
        GP2X = 208
    Case 3
        GP2X = 210
    Case 4
        GP2X = 211
    Case 5
        GP2X = 212
    Case 6
        GP2X = 214
    Case 7
        GP2X = 215
    Case 8
        GP2X = 216
    Case 9
        GP2X = 217
    Case 10
        GP2X = 218
    Case 11
        GP2X = 220
    Case 12
        GP2X = 221
    Case 13
        GP2X = 222
    Case 14
        GP2X = 222
    Case 15
        GP2X = 223
    Case 16
        GP2X = 238
    Case 17
        GP2X = 239
    Case 18
        GP2X = 239
    Case 19
        GP2X = 238
    Case 20
        GP2X = 229
    Case 21
        GP2X = 228
    Case 22
        GP2X = 228
    Case 23
        GP2X = 227
    Case 24
        GP2X = 225
    Case 25
        GP2X = 224
    Case 26
        GP2X = 222
    Case 27
        GP2X = 221
    Case 28
        GP2X = 221
    Case 29
        GP2X = 220
    Case 30
        GP2X = 211
    Case 31
        GP2X = 209
    Case 32
        GP2X = 208
    Case 33
        GP2X = 207
    Case 34
        GP2X = 206
    Case 35
        GP2X = 205
    Case 36
        GP2X = 203
    Case 37
        GP2X = 202
    Case 38
        GP2X = 202
    Case 39
        GP2X = 201
    Case 40
        GP2X = 193
    Case 41
        GP2X = 192
    End Select
End Function
Private Function GP2Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP2Y = 3
    Case 1
        GP2Y = 3
    Case 2
        GP2Y = 4
    Case 3
        GP2Y = 12
    Case 4
        GP2Y = 15
    Case 5
        GP2Y = 18
    Case 6
        GP2Y = 26
    Case 7
        GP2Y = 27
    Case 8
        GP2Y = 26
    Case 9
        GP2Y = 23
    Case 10
        GP2Y = 20
    Case 11
        GP2Y = 12
    Case 12
        GP2Y = 9
    Case 13
        GP2Y = 6
    Case 14
        GP2Y = 4
    Case 15
        GP2Y = 3
    Case 16
        GP2Y = 3
    Case 17
        GP2Y = 4
    Case 18
        GP2Y = 47
    Case 19
        GP2Y = 48
    Case 20
        GP2Y = 48
    Case 21
        GP2Y = 47
    Case 22
        GP2Y = 22
    Case 23
        GP2Y = 23
    Case 24
        GP2Y = 31
    Case 25
        GP2Y = 34
    Case 26
        GP2Y = 42
    Case 27
        GP2Y = 45
    Case 28
        GP2Y = 47
    Case 29
        GP2Y = 48
    Case 30
        GP2Y = 48
    Case 31
        GP2Y = 46
    Case 32
        GP2Y = 39
    Case 33
        GP2Y = 38
    Case 34
        GP2Y = 33
    Case 35
        GP2Y = 30
    Case 36
        GP2Y = 22
    Case 37
        GP2Y = 23
    Case 38
        GP2Y = 47
    Case 39
        GP2Y = 48
    Case 40
        GP2Y = 48
    Case 41
        GP2Y = 47
    End Select
End Function
Private Function GP3X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP3X = 291
    Case 1
        GP3X = 292
    Case 2
        GP3X = 293
    Case 3
        GP3X = 293
    Case 4
        GP3X = 294
    Case 5
        GP3X = 298
    Case 6
        GP3X = 299
    Case 7
        GP3X = 299
    Case 8
        GP3X = 298
    Case 9
        GP3X = 293
    Case 10
        GP3X = 293
    Case 11
        GP3X = 295
    Case 12
        GP3X = 298
    Case 13
        GP3X = 300
    Case 14
        GP3X = 300
    Case 15
        GP3X = 304
    Case 16
        GP3X = 307
    Case 17
        GP3X = 312
    Case 18
        GP3X = 315
    Case 19
        GP3X = 318
    Case 20
        GP3X = 320
    Case 21
        GP3X = 320
    Case 22
        GP3X = 312
    Case 23
        GP3X = 311
    Case 24
        GP3X = 311
    Case 25
        GP3X = 310
    Case 26
        GP3X = 305
    Case 27
        GP3X = 304
    Case 28
        GP3X = 301
    Case 29
        GP3X = 300
    Case 30
        GP3X = 301
    Case 31
        GP3X = 302
    Case 32
        GP3X = 306
    Case 33
        GP3X = 309
    Case 34
        GP3X = 318
    Case 35
        GP3X = 321
    Case 36
        GP3X = 325
    Case 37
        GP3X = 326
    Case 38
        GP3X = 327
    Case 39
        GP3X = 330
    Case 40
        GP3X = 331
    Case 41
        GP3X = 331
    Case 42
        GP3X = 332
    Case 43
        GP3X = 332
    Case 44
        GP3X = 333
    Case 45
        GP3X = 333
    Case 46
        GP3X = 332
    Case 47
        GP3X = 323
    Case 48
        GP3X = 322
    Case 49
        GP3X = 322
    Case 50
        GP3X = 321
    Case 51
        GP3X = 319
    Case 52
        GP3X = 316
    Case 53
        GP3X = 311
    Case 54
        GP3X = 310
    Case 55
        GP3X = 307
    Case 56
        GP3X = 306
    Case 57
        GP3X = 303
    Case 58
        GP3X = 302
    Case 59
        GP3X = 302
    Case 60
        GP3X = 301
    Case 61
        GP3X = 300
    Case 62
        GP3X = 300
    Case 63
        GP3X = 299
    Case 64
        GP3X = 290
    Case 65
        GP3X = 289
    Case 66
        GP3X = 287
    Case 67
        GP3X = 283
    Case 68
        GP3X = 282
    Case 69
        GP3X = 282
    Case 70
        GP3X = 279
    Case 71
        GP3X = 278
    Case 72
        GP3X = 278
    Case 73
        GP3X = 281
    Case 74
        GP3X = 282
    Case 75
        GP3X = 282
    Case 76
        GP3X = 285
    Case 77
        GP3X = 287
    End Select
End Function
Private Function GP3Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP3Y = 4
    Case 1
        GP3Y = 4
    Case 2
        GP3Y = 5
    Case 3
        GP3Y = 14
    Case 4
        GP3Y = 15
    Case 5
        GP3Y = 15
    Case 6
        GP3Y = 16
    Case 7
        GP3Y = 22
    Case 8
        GP3Y = 23
    Case 9
        GP3Y = 23
    Case 10
        GP3Y = 38
    Case 11
        GP3Y = 40
    Case 12
        GP3Y = 39
    Case 13
        GP3Y = 37
    Case 14
        GP3Y = 34
    Case 15
        GP3Y = 30
    Case 16
        GP3Y = 29
    Case 17
        GP3Y = 28
    Case 18
        GP3Y = 27
    Case 19
        GP3Y = 27
    Case 20
        GP3Y = 25
    Case 21
        GP3Y = 23
    Case 22
        GP3Y = 23
    Case 23
        GP3Y = 24
    Case 24
        GP3Y = 25
    Case 25
        GP3Y = 26
    Case 26
        GP3Y = 26
    Case 27
        GP3Y = 25
    Case 28
        GP3Y = 25
    Case 29
        GP3Y = 24
    Case 30
        GP3Y = 23
    Case 31
        GP3Y = 20
    Case 32
        GP3Y = 16
    Case 33
        GP3Y = 15
    Case 34
        GP3Y = 14
    Case 35
        GP3Y = 15
    Case 36
        GP3Y = 15
    Case 37
        GP3Y = 16
    Case 38
        GP3Y = 16
    Case 39
        GP3Y = 19
    Case 40
        GP3Y = 22
    Case 41
        GP3Y = 42
    Case 42
        GP3Y = 43
    Case 43
        GP3Y = 45
    Case 44
        GP3Y = 46
    Case 45
        GP3Y = 47
    Case 46
        GP3Y = 48
    Case 47
        GP3Y = 48
    Case 48
        GP3Y = 47
    Case 49
        GP3Y = 45
    Case 50
        GP3Y = 45
    Case 51
        GP3Y = 47
    Case 52
        GP3Y = 48
    Case 53
        GP3Y = 49
    Case 54
        GP3Y = 48
    Case 55
        GP3Y = 48
    Case 56
        GP3Y = 47
    Case 57
        GP3Y = 46
    Case 58
        GP3Y = 45
    Case 59
        GP3Y = 44
    Case 60
        GP3Y = 43
    Case 61
        GP3Y = 44
    Case 62
        GP3Y = 47
    Case 63
        GP3Y = 48
    Case 64
        GP3Y = 49
    Case 65
        GP3Y = 48
    Case 66
        GP3Y = 48
    Case 67
        GP3Y = 44
    Case 68
        GP3Y = 35
    Case 69
        GP3Y = 23
    Case 70
        GP3Y = 23
    Case 71
        GP3Y = 22
    Case 72
        GP3Y = 15
    Case 73
        GP3Y = 15
    Case 74
        GP3Y = 14
    Case 75
        GP3Y = 9
    Case 76
        GP3Y = 8
    Case 77
        GP3Y = 6
    End Select
End Function
Private Function GP4X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP4X = 44
    Case 1
        GP4X = 57
    Case 2
        GP4X = 58
    Case 3
        GP4X = 60
    Case 4
        GP4X = 61
    Case 5
        GP4X = 60
    Case 6
        GP4X = 60
    Case 7
        GP4X = 59
    Case 8
        GP4X = 56
    Case 9
        GP4X = 45
    Case 10
        GP4X = 44
    End Select
End Function
Private Function GP4Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP4Y = 12
    Case 1
        GP4Y = 12
    Case 2
        GP4Y = 13
    Case 3
        GP4Y = 13
    Case 4
        GP4Y = 16
    Case 5
        GP4Y = 17
    Case 6
        GP4Y = 18
    Case 7
        GP4Y = 19
    Case 8
        GP4Y = 20
    Case 9
        GP4Y = 20
    Case 10
        GP4Y = 19
    End Select
End Function
Private Function GP5X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP5X = 95
    Case 1
        GP5X = 100
    Case 2
        GP5X = 103
    Case 3
        GP5X = 104
    Case 4
        GP5X = 107
    Case 5
        GP5X = 113
    Case 6
        GP5X = 114
    Case 7
        GP5X = 114
    Case 8
        GP5X = 113
    Case 9
        GP5X = 113
    Case 10
        GP5X = 111
    Case 11
        GP5X = 111
    Case 12
        GP5X = 108
    Case 13
        GP5X = 107
    Case 14
        GP5X = 105
    Case 15
        GP5X = 102
    Case 16
        GP5X = 95
    Case 17
        GP5X = 94
    Case 18
        GP5X = 90
    Case 19
        GP5X = 89
    Case 20
        GP5X = 84
    Case 21
        GP5X = 82
    Case 22
        GP5X = 82
    Case 23
        GP5X = 80
    Case 24
        GP5X = 78
    Case 25
        GP5X = 78
    Case 26
        GP5X = 79
    Case 27
        GP5X = 79
    Case 28
        GP5X = 80
    Case 29
        GP5X = 80
    Case 30
        GP5X = 85
    Case 31
        GP5X = 90
    End Select
End Function
Private Function GP5Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP5Y = 14
    Case 1
        GP5Y = 15
    Case 2
        GP5Y = 15
    Case 3
        GP5Y = 16
    Case 4
        GP5Y = 17
    Case 5
        GP5Y = 23
    Case 6
        GP5Y = 26
    Case 7
        GP5Y = 36
    Case 8
        GP5Y = 37
    Case 9
        GP5Y = 39
    Case 10
        GP5Y = 41
    Case 11
        GP5Y = 42
    Case 12
        GP5Y = 45
    Case 13
        GP5Y = 45
    Case 14
        GP5Y = 47
    Case 15
        GP5Y = 48
    Case 16
        GP5Y = 49
    Case 17
        GP5Y = 48
    Case 18
        GP5Y = 48
    Case 19
        GP5Y = 47
    Case 20
        GP5Y = 45
    Case 21
        GP5Y = 43
    Case 22
        GP5Y = 42
    Case 23
        GP5Y = 40
    Case 24
        GP5Y = 35
    Case 25
        GP5Y = 27
    Case 26
        GP5Y = 26
    Case 27
        GP5Y = 24
    Case 28
        GP5Y = 23
    Case 29
        GP5Y = 22
    Case 30
        GP5Y = 17
    Case 31
        GP5Y = 15
    End Select
End Function
Private Function GP6X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP6X = 150
    Case 1
        GP6X = 161
    Case 2
        GP6X = 163
    Case 3
        GP6X = 164
    Case 4
        GP6X = 166
    Case 5
        GP6X = 167
    Case 6
        GP6X = 167
    Case 7
        GP6X = 168
    Case 8
        GP6X = 168
    Case 9
        GP6X = 169
    Case 10
        GP6X = 169
    Case 11
        GP6X = 168
    Case 12
        GP6X = 159
    Case 13
        GP6X = 157
    Case 14
        GP6X = 152
    Case 15
        GP6X = 147
    Case 16
        GP6X = 146
    Case 17
        GP6X = 142
    Case 18
        GP6X = 140
    Case 19
        GP6X = 139
    Case 20
        GP6X = 137
    Case 21
        GP6X = 135
    Case 22
        GP6X = 136
    Case 23
        GP6X = 138
    Case 24
        GP6X = 138
    Case 25
        GP6X = 143
    Case 26
        GP6X = 148
    Case 27
        GP6X = 151
    Case 28
        GP6X = 154
    Case 29
        GP6X = 156
    Case 30
        GP6X = 156
    Case 31
        GP6X = 155
    Case 32
        GP6X = 148
    Case 33
        GP6X = 145
    Case 34
        GP6X = 141
    Case 35
        GP6X = 140
    Case 36
        GP6X = 137
    Case 37
        GP6X = 136
    Case 38
        GP6X = 138
    Case 39
        GP6X = 140
    Case 40
        GP6X = 145
    End Select
End Function
Private Function GP6Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP6Y = 14
    Case 1
        GP6Y = 15
    Case 2
        GP6Y = 17
    Case 3
        GP6Y = 17
    Case 4
        GP6Y = 19
    Case 5
        GP6Y = 24
    Case 6
        GP6Y = 42
    Case 7
        GP6Y = 43
    Case 8
        GP6Y = 45
    Case 9
        GP6Y = 46
    Case 10
        GP6Y = 47
    Case 11
        GP6Y = 48
    Case 12
        GP6Y = 48
    Case 13
        GP6Y = 46
    Case 14
        GP6Y = 48
    Case 15
        GP6Y = 49
    Case 16
        GP6Y = 48
    Case 17
        GP6Y = 48
    Case 18
        GP6Y = 46
    Case 19
        GP6Y = 46
    Case 20
        GP6Y = 44
    Case 21
        GP6Y = 39
    Case 22
        GP6Y = 34
    Case 23
        GP6Y = 32
    Case 24
        GP6Y = 31
    Case 25
        GP6Y = 29
    Case 26
        GP6Y = 28
    Case 27
        GP6Y = 27
    Case 28
        GP6Y = 27
    Case 29
        GP6Y = 25
    Case 30
        GP6Y = 24
    Case 31
        GP6Y = 23
    Case 32
        GP6Y = 23
    Case 33
        GP6Y = 26
    Case 34
        GP6Y = 26
    Case 35
        GP6Y = 25
    Case 36
        GP6Y = 25
    Case 37
        GP6Y = 24
    Case 38
        GP6Y = 19
    Case 39
        GP6Y = 17
    Case 40
        GP6Y = 15
    End Select
End Function
Private Function GP7X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP7X = 258
    Case 1
        GP7X = 269
    Case 2
        GP7X = 270
    Case 3
        GP7X = 271
    Case 4
        GP7X = 274
    Case 5
        GP7X = 275
    Case 6
        GP7X = 275
    Case 7
        GP7X = 276
    Case 8
        GP7X = 277
    Case 9
        GP7X = 277
    Case 10
        GP7X = 276
    Case 11
        GP7X = 267
    Case 12
        GP7X = 266
    Case 13
        GP7X = 266
    Case 14
        GP7X = 265
    Case 15
        GP7X = 263
    Case 16
        GP7X = 260
    Case 17
        GP7X = 255
    Case 18
        GP7X = 254
    Case 19
        GP7X = 251
    Case 20
        GP7X = 250
    Case 21
        GP7X = 247
    Case 22
        GP7X = 246
    Case 23
        GP7X = 246
    Case 24
        GP7X = 244
    Case 25
        GP7X = 244
    Case 26
        GP7X = 248
    Case 27
        GP7X = 251
    Case 28
        GP7X = 256
    Case 29
        GP7X = 261
    Case 30
        GP7X = 262
    Case 31
        GP7X = 264
    Case 32
        GP7X = 264
    Case 33
        GP7X = 256
    Case 34
        GP7X = 255
    Case 35
        GP7X = 255
    Case 36
        GP7X = 254
    Case 37
        GP7X = 249
    Case 38
        GP7X = 248
    Case 39
        GP7X = 245
    Case 40
        GP7X = 244
    Case 41
        GP7X = 245
    Case 42
        GP7X = 246
    Case 43
        GP7X = 250
    Case 44
        GP7X = 253
    End Select
End Function
Private Function GP7Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP7Y = 14
    Case 1
        GP7Y = 15
    Case 2
        GP7Y = 16
    Case 3
        GP7Y = 16
    Case 4
        GP7Y = 19
    Case 5
        GP7Y = 22
    Case 6
        GP7Y = 42
    Case 7
        GP7Y = 43
    Case 8
        GP7Y = 46
    Case 9
        GP7Y = 47
    Case 10
        GP7Y = 48
    Case 11
        GP7Y = 48
    Case 12
        GP7Y = 47
    Case 13
        GP7Y = 45
    Case 14
        GP7Y = 45
    Case 15
        GP7Y = 47
    Case 16
        GP7Y = 48
    Case 17
        GP7Y = 49
    Case 18
        GP7Y = 48
    Case 19
        GP7Y = 48
    Case 20
        GP7Y = 47
    Case 21
        GP7Y = 46
    Case 22
        GP7Y = 45
    Case 23
        GP7Y = 44
    Case 24
        GP7Y = 42
    Case 25
        GP7Y = 34
    Case 26
        GP7Y = 30
    Case 27
        GP7Y = 29
    Case 28
        GP7Y = 28
    Case 29
        GP7Y = 27
    Case 30
        GP7Y = 27
    Case 31
        GP7Y = 25
    Case 32
        GP7Y = 23
    Case 33
        GP7Y = 23
    Case 34
        GP7Y = 24
    Case 35
        GP7Y = 25
    Case 36
        GP7Y = 26
    Case 37
        GP7Y = 26
    Case 38
        GP7Y = 25
    Case 39
        GP7Y = 25
    Case 40
        GP7Y = 24
    Case 41
        GP7Y = 23
    Case 42
        GP7Y = 20
    Case 43
        GP7Y = 16
    Case 44
        GP7Y = 15
    End Select
End Function
Private Function GP8X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP8X = 93
    Case 1
        GP8X = 90
    Case 2
        GP8X = 89
    Case 3
        GP8X = 89
    Case 4
        GP8X = 90
    Case 5
        GP8X = 90
    Case 6
        GP8X = 94
    Case 7
        GP8X = 98
    Case 8
        GP8X = 99
    Case 9
        GP8X = 100
    Case 10
        GP8X = 102
    Case 11
        GP8X = 103
    Case 12
        GP8X = 103
    Case 13
        GP8X = 99
    End Select
End Function
Private Function GP8Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP8Y = 23
    Case 1
        GP8Y = 26
    Case 2
        GP8Y = 29
    Case 3
        GP8Y = 34
    Case 4
        GP8Y = 35
    Case 5
        GP8Y = 36
    Case 6
        GP8Y = 40
    Case 7
        GP8Y = 40
    Case 8
        GP8Y = 39
    Case 9
        GP8Y = 39
    Case 10
        GP8Y = 37
    Case 11
        GP8Y = 34
    Case 12
        GP8Y = 27
    Case 13
        GP8Y = 23
    End Select
End Function
Private Function GP9X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP9X = 44
    Case 1
        GP9X = 57
    Case 2
        GP9X = 58
    Case 3
        GP9X = 60
    Case 4
        GP9X = 62
    Case 5
        GP9X = 62
    Case 6
        GP9X = 60
    Case 7
        GP9X = 57
    Case 8
        GP9X = 45
    Case 9
        GP9X = 44
    End Select
End Function
Private Function GP9Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP9Y = 29
    Case 1
        GP9Y = 29
    Case 2
        GP9Y = 30
    Case 3
        GP9Y = 30
    Case 4
        GP9Y = 32
    Case 5
        GP9Y = 36
    Case 6
        GP9Y = 38
    Case 7
        GP9Y = 39
    Case 8
        GP9Y = 39
    Case 9
        GP9Y = 38
    End Select
End Function
Private Function GP10X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP10X = 151
    Case 1
        GP10X = 156
    Case 2
        GP10X = 156
    Case 3
        GP10X = 152
    Case 4
        GP10X = 149
    Case 5
        GP10X = 146
    Case 6
        GP10X = 146
    Case 7
        GP10X = 148
    End Select
End Function
Private Function GP10Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP10Y = 34
    Case 1
        GP10Y = 34
    Case 2
        GP10Y = 37
    Case 3
        GP10Y = 41
    Case 4
        GP10Y = 41
    Case 5
        GP10Y = 38
    Case 6
        GP10Y = 37
    Case 7
        GP10Y = 35
    End Select
End Function
Private Function GP11X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP11X = 260
    Case 1
        GP11X = 264
    Case 2
        GP11X = 264
    Case 3
        GP11X = 262
    Case 4
        GP11X = 259
    Case 5
        GP11X = 257
    Case 6
        GP11X = 255
    Case 7
        GP11X = 255
    Case 8
        GP11X = 256
    End Select
End Function
Private Function GP11Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP11Y = 34
    Case 1
        GP11Y = 34
    Case 2
        GP11Y = 38
    Case 3
        GP11Y = 40
    Case 4
        GP11Y = 41
    Case 5
        GP11Y = 41
    Case 6
        GP11Y = 39
    Case 7
        GP11Y = 36
    Case 8
        GP11Y = 35
    End Select
End Function
Private Function GP12X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP12X = 316
    Case 1
        GP12X = 320
    Case 2
        GP12X = 320
    Case 3
        GP12X = 318
    Case 4
        GP12X = 315
    Case 5
        GP12X = 313
    Case 6
        GP12X = 311
    Case 7
        GP12X = 311
    Case 8
        GP12X = 310
    Case 9
        GP12X = 312
    End Select
End Function
Private Function GP12Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP12Y = 34
    Case 1
        GP12Y = 34
    Case 2
        GP12Y = 38
    Case 3
        GP12Y = 40
    Case 4
        GP12Y = 41
    Case 5
        GP12Y = 41
    Case 6
        GP12Y = 39
    Case 7
        GP12Y = 38
    Case 8
        GP12Y = 37
    Case 9
        GP12Y = 35
    End Select
End Function
Private Function GP13X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP13X = 147
    Case 1
        GP13X = 157
    Case 2
        GP13X = 158
    Case 3
        GP13X = 158
    Case 4
        GP13X = 163
    Case 5
        GP13X = 172
    Case 6
        GP13X = 173
    Case 7
        GP13X = 176
    Case 8
        GP13X = 179
    Case 9
        GP13X = 181
    Case 10
        GP13X = 182
    Case 11
        GP13X = 182
    Case 12
        GP13X = 181
    Case 13
        GP13X = 181
    Case 14
        GP13X = 180
    Case 15
        GP13X = 179
    Case 16
        GP13X = 174
    Case 17
        GP13X = 171
    Case 18
        GP13X = 168
    Case 19
        GP13X = 164
    Case 20
        GP13X = 163
    Case 21
        GP13X = 160
    Case 22
        GP13X = 158
    Case 23
        GP13X = 157
    Case 24
        GP13X = 157
    Case 25
        GP13X = 156
    Case 26
        GP13X = 148
    Case 27
        GP13X = 147
    End Select
End Function
Private Function GP13Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP13Y = 70
    Case 1
        GP13Y = 70
    Case 2
        GP13Y = 71
    Case 3
        GP13Y = 83
    Case 4
        GP13Y = 82
    Case 5
        GP13Y = 82
    Case 6
        GP13Y = 83
    Case 7
        GP13Y = 84
    Case 8
        GP13Y = 87
    Case 9
        GP13Y = 92
    Case 10
        GP13Y = 97
    Case 11
        GP13Y = 101
    Case 12
        GP13Y = 102
    Case 13
        GP13Y = 105
    Case 14
        GP13Y = 106
    Case 15
        GP13Y = 109
    Case 16
        GP13Y = 114
    Case 17
        GP13Y = 115
    Case 18
        GP13Y = 116
    Case 19
        GP13Y = 116
    Case 20
        GP13Y = 115
    Case 21
        GP13Y = 114
    Case 22
        GP13Y = 112
    Case 23
        GP13Y = 113
    Case 24
        GP13Y = 114
    Case 25
        GP13Y = 115
    Case 26
        GP13Y = 115
    Case 27
        GP13Y = 114
    End Select
End Function
Private Function GP14X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP14X = 183
    Case 1
        GP14X = 193
    Case 2
        GP14X = 195
    Case 3
        GP14X = 195
    Case 4
        GP14X = 196
    Case 5
        GP14X = 197
    Case 6
        GP14X = 198
    Case 7
        GP14X = 199
    Case 8
        GP14X = 200
    Case 9
        GP14X = 201
    Case 10
        GP14X = 203
    Case 11
        GP14X = 206
    Case 12
        GP14X = 207
    Case 13
        GP14X = 208
    Case 14
        GP14X = 208
    Case 15
        GP14X = 218
    Case 16
        GP14X = 219
    Case 17
        GP14X = 218
    Case 18
        GP14X = 216
    Case 19
        GP14X = 214
    Case 20
        GP14X = 213
    Case 21
        GP14X = 211
    Case 22
        GP14X = 209
    Case 23
        GP14X = 207
    Case 24
        GP14X = 203
    Case 25
        GP14X = 201
    Case 26
        GP14X = 200
    Case 27
        GP14X = 198
    Case 28
        GP14X = 194
    Case 29
        GP14X = 193
    Case 30
        GP14X = 190
    Case 31
        GP14X = 189
    Case 32
        GP14X = 187
    Case 33
        GP14X = 186
    Case 34
        GP14X = 185
    Case 35
        GP14X = 185
    Case 36
        GP14X = 192
    Case 37
        GP14X = 194
    Case 38
        GP14X = 195
    Case 39
        GP14X = 195
    Case 40
        GP14X = 194
    Case 41
        GP14X = 193
    Case 42
        GP14X = 191
    Case 43
        GP14X = 189
    Case 44
        GP14X = 186
    Case 45
        GP14X = 184
    End Select
End Function
Private Function GP14Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP14Y = 82
    Case 1
        GP14Y = 82
    Case 2
        GP14Y = 84
    Case 3
        GP14Y = 86
    Case 4
        GP14Y = 87
    Case 5
        GP14Y = 90
    Case 6
        GP14Y = 93
    Case 7
        GP14Y = 96
    Case 8
        GP14Y = 99
    Case 9
        GP14Y = 100
    Case 10
        GP14Y = 98
    Case 11
        GP14Y = 87
    Case 12
        GP14Y = 86
    Case 13
        GP14Y = 83
    Case 14
        GP14Y = 82
    Case 15
        GP14Y = 82
    Case 16
        GP14Y = 83
    Case 17
        GP14Y = 86
    Case 18
        GP14Y = 91
    Case 19
        GP14Y = 96
    Case 20
        GP14Y = 99
    Case 21
        GP14Y = 104
    Case 22
        GP14Y = 109
    Case 23
        GP14Y = 114
    Case 24
        GP14Y = 123
    Case 25
        GP14Y = 125
    Case 26
        GP14Y = 125
    Case 27
        GP14Y = 127
    Case 28
        GP14Y = 127
    Case 29
        GP14Y = 128
    Case 30
        GP14Y = 128
    Case 31
        GP14Y = 127
    Case 32
        GP14Y = 127
    Case 33
        GP14Y = 126
    Case 34
        GP14Y = 123
    Case 35
        GP14Y = 119
    Case 36
        GP14Y = 119
    Case 37
        GP14Y = 117
    Case 38
        GP14Y = 114
    Case 39
        GP14Y = 112
    Case 40
        GP14Y = 111
    Case 41
        GP14Y = 108
    Case 42
        GP14Y = 103
    Case 43
        GP14Y = 98
    Case 44
        GP14Y = 91
    Case 45
        GP14Y = 86
    End Select
End Function
Private Function GP15X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP15X = 161
    Case 1
        GP15X = 167
    Case 2
        GP15X = 170
    Case 3
        GP15X = 170
    Case 4
        GP15X = 171
    Case 5
        GP15X = 170
    Case 6
        GP15X = 170
    Case 7
        GP15X = 169
    Case 8
        GP15X = 169
    Case 9
        GP15X = 168
    Case 10
        GP15X = 165
    Case 11
        GP15X = 162
    Case 12
        GP15X = 159
    Case 13
        GP15X = 158
    Case 14
        GP15X = 158
    Case 15
        GP15X = 159
    Case 16
        GP15X = 159
    End Select
End Function
Private Function GP15Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP15Y = 90
    Case 1
        GP15Y = 90
    Case 2
        GP15Y = 93
    Case 3
        GP15Y = 97
    Case 4
        GP15Y = 98
    Case 5
        GP15Y = 101
    Case 6
        GP15Y = 103
    Case 7
        GP15Y = 104
    Case 8
        GP15Y = 105
    Case 9
        GP15Y = 106
    Case 10
        GP15Y = 107
    Case 11
        GP15Y = 107
    Case 12
        GP15Y = 104
    Case 13
        GP15Y = 101
    Case 14
        GP15Y = 94
    Case 15
        GP15Y = 93
    Case 16
        GP15Y = 92
    End Select
End Function
Private Function GP16X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP16X = 49
    Case 1
        GP16X = 49
    Case 2
        GP16X = 50
    Case 3
        GP16X = 59
    Case 4
        GP16X = 60
    Case 5
        GP16X = 60
    Case 6
        GP16X = 66
    Case 7
        GP16X = 67
    Case 8
        GP16X = 68
    Case 9
        GP16X = 68
    Case 10
        GP16X = 70
    Case 11
        GP16X = 70
    Case 12
        GP16X = 72
    Case 13
        GP16X = 72
    Case 14
        GP16X = 74
    Case 15
        GP16X = 75
    Case 16
        GP16X = 77
    Case 17
        GP16X = 77
    Case 18
        GP16X = 80
    Case 19
        GP16X = 92
    Case 20
        GP16X = 93
    Case 21
        GP16X = 92
    Case 22
        GP16X = 92
    Case 23
        GP16X = 90
    Case 24
        GP16X = 90
    Case 25
        GP16X = 88
    Case 26
        GP16X = 88
    Case 27
        GP16X = 85
    Case 28
        GP16X = 85
    Case 29
        GP16X = 83
    Case 30
        GP16X = 83
    Case 31
        GP16X = 81
    Case 32
        GP16X = 81
    Case 33
        GP16X = 78
    Case 34
        GP16X = 78
    Case 35
        GP16X = 76
    Case 36
        GP16X = 75
    Case 37
        GP16X = 76
    Case 38
        GP16X = 77
    Case 39
        GP16X = 92
    Case 40
        GP16X = 91
    Case 41
        GP16X = 78
    Case 42
        GP16X = 61
    Case 43
        GP16X = 60
    Case 44
        GP16X = 60
    Case 45
        GP16X = 59
    End Select
End Function
Private Function GP16Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP16Y = 137
    Case 1
        GP16Y = 181
    Case 2
        GP16Y = 182
    Case 3
        GP16Y = 182
    Case 4
        GP16Y = 181
    Case 5
        GP16Y = 169
    Case 6
        GP16Y = 163
    Case 7
        GP16Y = 163
    Case 8
        GP16Y = 164
    Case 9
        GP16Y = 165
    Case 10
        GP16Y = 167
    Case 11
        GP16Y = 168
    Case 12
        GP16Y = 170
    Case 13
        GP16Y = 171
    Case 14
        GP16Y = 173
    Case 15
        GP16Y = 176
    Case 16
        GP16Y = 178
    Case 17
        GP16Y = 179
    Case 18
        GP16Y = 182
    Case 19
        GP16Y = 182
    Case 20
        GP16Y = 181
    Case 21
        GP16Y = 180
    Case 22
        GP16Y = 179
    Case 23
        GP16Y = 177
    Case 24
        GP16Y = 176
    Case 25
        GP16Y = 174
    Case 26
        GP16Y = 173
    Case 27
        GP16Y = 170
    Case 28
        GP16Y = 169
    Case 29
        GP16Y = 167
    Case 30
        GP16Y = 166
    Case 31
        GP16Y = 164
    Case 32
        GP16Y = 163
    Case 33
        GP16Y = 160
    Case 34
        GP16Y = 159
    Case 35
        GP16Y = 157
    Case 36
        GP16Y = 154
    Case 37
        GP16Y = 153
    Case 38
        GP16Y = 153
    Case 39
        GP16Y = 138
    Case 40
        GP16Y = 137
    Case 41
        GP16Y = 137
    Case 42
        GP16Y = 154
    Case 43
        GP16Y = 153
    Case 44
        GP16Y = 138
    Case 45
        GP16Y = 137
    End Select
End Function
Private Function GP17X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP17X = 113
    Case 1
        GP17X = 121
    Case 2
        GP17X = 122
    Case 3
        GP17X = 122
    Case 4
        GP17X = 123
    Case 5
        GP17X = 127
    Case 6
        GP17X = 128
    Case 7
        GP17X = 128
    Case 8
        GP17X = 126
    Case 9
        GP17X = 122
    Case 10
        GP17X = 122
    Case 11
        GP17X = 121
    Case 12
        GP17X = 112
    Case 13
        GP17X = 111
    Case 14
        GP17X = 111
    Case 15
        GP17X = 94
    Case 16
        GP17X = 92
    Case 17
        GP17X = 92
    Case 18
        GP17X = 93
    Case 19
        GP17X = 93
    Case 20
        GP17X = 96
    Case 21
        GP17X = 96
    Case 22
        GP17X = 99
    Case 23
        GP17X = 99
    Case 24
        GP17X = 102
    Case 25
        GP17X = 102
    Case 26
        GP17X = 105
    Case 27
        GP17X = 105
    Case 28
        GP17X = 108
    Case 29
        GP17X = 108
    Case 30
        GP17X = 110
    Case 31
        GP17X = 110
    End Select
End Function
Private Function GP17Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP17Y = 137
    Case 1
        GP17Y = 137
    Case 2
        GP17Y = 138
    Case 3
        GP17Y = 164
    Case 4
        GP17Y = 165
    Case 5
        GP17Y = 165
    Case 6
        GP17Y = 166
    Case 7
        GP17Y = 172
    Case 8
        GP17Y = 174
    Case 9
        GP17Y = 174
    Case 10
        GP17Y = 181
    Case 11
        GP17Y = 182
    Case 12
        GP17Y = 182
    Case 13
        GP17Y = 181
    Case 14
        GP17Y = 174
    Case 15
        GP17Y = 174
    Case 16
        GP17Y = 172
    Case 17
        GP17Y = 165
    Case 18
        GP17Y = 164
    Case 19
        GP17Y = 163
    Case 20
        GP17Y = 160
    Case 21
        GP17Y = 159
    Case 22
        GP17Y = 156
    Case 23
        GP17Y = 155
    Case 24
        GP17Y = 152
    Case 25
        GP17Y = 151
    Case 26
        GP17Y = 148
    Case 27
        GP17Y = 147
    Case 28
        GP17Y = 144
    Case 29
        GP17Y = 143
    Case 30
        GP17Y = 141
    Case 31
        GP17Y = 140
    End Select
End Function
Private Function GP18X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP18X = 207
    Case 1
        GP18X = 215
    Case 2
        GP18X = 216
    Case 3
        GP18X = 216
    Case 4
        GP18X = 217
    Case 5
        GP18X = 220
    Case 6
        GP18X = 221
    Case 7
        GP18X = 221
    Case 8
        GP18X = 220
    Case 9
        GP18X = 216
    Case 10
        GP18X = 216
    Case 11
        GP18X = 215
    Case 12
        GP18X = 206
    Case 13
        GP18X = 205
    Case 14
        GP18X = 205
    Case 15
        GP18X = 188
    Case 16
        GP18X = 186
    Case 17
        GP18X = 186
    Case 18
        GP18X = 187
    Case 19
        GP18X = 187
    Case 20
        GP18X = 190
    Case 21
        GP18X = 190
    Case 22
        GP18X = 192
    Case 23
        GP18X = 192
    Case 24
        GP18X = 195
    Case 25
        GP18X = 195
    Case 26
        GP18X = 198
    Case 27
        GP18X = 198
    Case 28
        GP18X = 201
    Case 29
        GP18X = 201
    Case 30
        GP18X = 204
    Case 31
        GP18X = 204
    End Select
End Function
Private Function GP18Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP18Y = 137
    Case 1
        GP18Y = 137
    Case 2
        GP18Y = 138
    Case 3
        GP18Y = 164
    Case 4
        GP18Y = 165
    Case 5
        GP18Y = 165
    Case 6
        GP18Y = 166
    Case 7
        GP18Y = 173
    Case 8
        GP18Y = 174
    Case 9
        GP18Y = 174
    Case 10
        GP18Y = 181
    Case 11
        GP18Y = 182
    Case 12
        GP18Y = 182
    Case 13
        GP18Y = 181
    Case 14
        GP18Y = 174
    Case 15
        GP18Y = 174
    Case 16
        GP18Y = 172
    Case 17
        GP18Y = 165
    Case 18
        GP18Y = 164
    Case 19
        GP18Y = 163
    Case 20
        GP18Y = 160
    Case 21
        GP18Y = 159
    Case 22
        GP18Y = 157
    Case 23
        GP18Y = 156
    Case 24
        GP18Y = 153
    Case 25
        GP18Y = 152
    Case 26
        GP18Y = 149
    Case 27
        GP18Y = 148
    Case 28
        GP18Y = 145
    Case 29
        GP18Y = 144
    Case 30
        GP18Y = 141
    Case 31
        GP18Y = 140
    End Select
End Function
Private Function GP19X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP19X = 261
    Case 1
        GP19X = 269
    Case 2
        GP19X = 270
    Case 3
        GP19X = 272
    Case 4
        GP19X = 273
    Case 5
        GP19X = 276
    Case 6
        GP19X = 278
    Case 7
        GP19X = 280
    Case 8
        GP19X = 280
    Case 9
        GP19X = 279
    Case 10
        GP19X = 278
    Case 11
        GP19X = 275
    Case 12
        GP19X = 276
    Case 13
        GP19X = 277
    Case 14
        GP19X = 280
    Case 15
        GP19X = 281
    Case 16
        GP19X = 281
    Case 17
        GP19X = 280
    Case 18
        GP19X = 280
    Case 19
        GP19X = 278
    Case 20
        GP19X = 278
    Case 21
        GP19X = 277
    Case 22
        GP19X = 276
    Case 23
        GP19X = 274
    Case 24
        GP19X = 271
    Case 25
        GP19X = 268
    Case 26
        GP19X = 261
    Case 27
        GP19X = 260
    Case 28
        GP19X = 258
    Case 29
        GP19X = 257
    Case 30
        GP19X = 256
    Case 31
        GP19X = 250
    Case 32
        GP19X = 249
    Case 33
        GP19X = 249
    Case 34
        GP19X = 250
    Case 35
        GP19X = 250
    Case 36
        GP19X = 252
    Case 37
        GP19X = 252
    Case 38
        GP19X = 253
    Case 39
        GP19X = 254
    Case 40
        GP19X = 255
    Case 41
        GP19X = 252
    Case 42
        GP19X = 250
    Case 43
        GP19X = 250
    Case 44
        GP19X = 251
    Case 45
        GP19X = 252
    Case 46
        GP19X = 254
    Case 47
        GP19X = 259
    End Select
End Function
Private Function GP19Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP19Y = 137
    Case 1
        GP19Y = 137
    Case 2
        GP19Y = 138
    Case 3
        GP19Y = 138
    Case 4
        GP19Y = 139
    Case 5
        GP19Y = 140
    Case 6
        GP19Y = 142
    Case 7
        GP19Y = 147
    Case 8
        GP19Y = 151
    Case 9
        GP19Y = 152
    Case 10
        GP19Y = 155
    Case 11
        GP19Y = 158
    Case 12
        GP19Y = 159
    Case 13
        GP19Y = 159
    Case 14
        GP19Y = 162
    Case 15
        GP19Y = 165
    Case 16
        GP19Y = 173
    Case 17
        GP19Y = 174
    Case 18
        GP19Y = 175
    Case 19
        GP19Y = 177
    Case 20
        GP19Y = 178
    Case 21
        GP19Y = 179
    Case 22
        GP19Y = 179
    Case 23
        GP19Y = 181
    Case 24
        GP19Y = 182
    Case 25
        GP19Y = 183
    Case 26
        GP19Y = 183
    Case 27
        GP19Y = 182
    Case 28
        GP19Y = 182
    Case 29
        GP19Y = 181
    Case 30
        GP19Y = 181
    Case 31
        GP19Y = 175
    Case 32
        GP19Y = 170
    Case 33
        GP19Y = 165
    Case 34
        GP19Y = 164
    Case 35
        GP19Y = 163
    Case 36
        GP19Y = 161
    Case 37
        GP19Y = 160
    Case 38
        GP19Y = 159
    Case 39
        GP19Y = 159
    Case 40
        GP19Y = 158
    Case 41
        GP19Y = 155
    Case 42
        GP19Y = 150
    Case 43
        GP19Y = 146
    Case 44
        GP19Y = 145
    Case 45
        GP19Y = 142
    Case 46
        GP19Y = 140
    Case 47
        GP19Y = 138
    End Select
End Function
Private Function GP20X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP20X = 300
    Case 1
        GP20X = 305
    Case 2
        GP20X = 306
    Case 3
        GP20X = 309
    Case 4
        GP20X = 310
    Case 5
        GP20X = 311
    Case 6
        GP20X = 316
    Case 7
        GP20X = 317
    Case 8
        GP20X = 317
    Case 9
        GP20X = 316
    Case 10
        GP20X = 315
    Case 11
        GP20X = 312
    Case 12
        GP20X = 312
    Case 13
        GP20X = 309
    Case 14
        GP20X = 308
    Case 15
        GP20X = 301
    Case 16
        GP20X = 302
    Case 17
        GP20X = 316
    Case 18
        GP20X = 317
    Case 19
        GP20X = 317
    Case 20
        GP20X = 316
    Case 21
        GP20X = 285
    Case 22
        GP20X = 284
    Case 23
        GP20X = 285
    Case 24
        GP20X = 286
    Case 25
        GP20X = 287
    Case 26
        GP20X = 301
    Case 27
        GP20X = 302
    Case 28
        GP20X = 304
    Case 29
        GP20X = 304
    Case 30
        GP20X = 306
    Case 31
        GP20X = 306
    Case 32
        GP20X = 304
    Case 33
        GP20X = 299
    Case 34
        GP20X = 296
    Case 35
        GP20X = 296
    Case 36
        GP20X = 294
    Case 37
        GP20X = 293
    Case 38
        GP20X = 286
    Case 39
        GP20X = 285
    Case 40
        GP20X = 285
    Case 41
        GP20X = 286
    Case 42
        GP20X = 286
    Case 43
        GP20X = 288
    Case 44
        GP20X = 288
    Case 45
        GP20X = 290
    Case 46
        GP20X = 295
    End Select
End Function
Private Function GP20Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP20Y = 137
    Case 1
        GP20Y = 137
    Case 2
        GP20Y = 138
    Case 3
        GP20Y = 138
    Case 4
        GP20Y = 139
    Case 5
        GP20Y = 139
    Case 6
        GP20Y = 144
    Case 7
        GP20Y = 147
    Case 8
        GP20Y = 154
    Case 9
        GP20Y = 155
    Case 10
        GP20Y = 158
    Case 11
        GP20Y = 161
    Case 12
        GP20Y = 162
    Case 13
        GP20Y = 165
    Case 14
        GP20Y = 165
    Case 15
        GP20Y = 172
    Case 16
        GP20Y = 173
    Case 17
        GP20Y = 173
    Case 18
        GP20Y = 174
    Case 19
        GP20Y = 181
    Case 20
        GP20Y = 182
    Case 21
        GP20Y = 182
    Case 22
        GP20Y = 181
    Case 23
        GP20Y = 176
    Case 24
        GP20Y = 175
    Case 25
        GP20Y = 172
    Case 26
        GP20Y = 158
    Case 27
        GP20Y = 158
    Case 28
        GP20Y = 156
    Case 29
        GP20Y = 155
    Case 30
        GP20Y = 153
    Case 31
        GP20Y = 148
    Case 32
        GP20Y = 146
    Case 33
        GP20Y = 146
    Case 34
        GP20Y = 149
    Case 35
        GP20Y = 151
    Case 36
        GP20Y = 153
    Case 37
        GP20Y = 152
    Case 38
        GP20Y = 152
    Case 39
        GP20Y = 151
    Case 40
        GP20Y = 148
    Case 41
        GP20Y = 147
    Case 42
        GP20Y = 145
    Case 43
        GP20Y = 143
    Case 44
        GP20Y = 142
    Case 45
        GP20Y = 140
    Case 46
        GP20Y = 138
    End Select
End Function
Private Function GP21X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP21X = 262
    Case 1
        GP21X = 268
    Case 2
        GP21X = 269
    Case 3
        GP21X = 269
    Case 4
        GP21X = 268
    Case 5
        GP21X = 265
    Case 6
        GP21X = 262
    Case 7
        GP21X = 261
    End Select
End Function
Private Function GP21Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP21Y = 146
    Case 1
        GP21Y = 146
    Case 2
        GP21Y = 147
    Case 3
        GP21Y = 152
    Case 4
        GP21Y = 153
    Case 5
        GP21Y = 154
    Case 6
        GP21Y = 153
    Case 7
        GP21Y = 152
    End Select
End Function
Private Function GP22X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP22X = 132
    Case 1
        GP22X = 131
    Case 2
        GP22X = 131
    Case 3
        GP22X = 132
    Case 4
        GP22X = 141
    Case 5
        GP22X = 142
    Case 6
        GP22X = 142
    Case 7
        GP22X = 143
    Case 8
        GP22X = 143
    Case 9
        GP22X = 145
    Case 10
        GP22X = 150
    Case 11
        GP22X = 151
    Case 12
        GP22X = 151
    Case 13
        GP22X = 152
    Case 14
        GP22X = 161
    Case 15
        GP22X = 162
    Case 16
        GP22X = 162
    Case 17
        GP22X = 163
    Case 18
        GP22X = 163
    Case 19
        GP22X = 165
    Case 20
        GP22X = 170
    Case 21
        GP22X = 171
    Case 22
        GP22X = 172
    Case 23
        GP22X = 172
    Case 24
        GP22X = 173
    Case 25
        GP22X = 182
    Case 26
        GP22X = 183
    Case 27
        GP22X = 183
    Case 28
        GP22X = 182
    Case 29
        GP22X = 182
    Case 30
        GP22X = 178
    Case 31
        GP22X = 175
    Case 32
        GP22X = 166
    Case 33
        GP22X = 165
    Case 34
        GP22X = 160
    Case 35
        GP22X = 158
    Case 36
        GP22X = 155
    Case 37
        GP22X = 146
    Case 38
        GP22X = 145
    Case 39
        GP22X = 142
    Case 40
        GP22X = 140
    End Select
End Function
Private Function GP22Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP22Y = 149
    Case 1
        GP22Y = 150
    Case 2
        GP22Y = 181
    Case 3
        GP22Y = 182
    Case 4
        GP22Y = 182
    Case 5
        GP22Y = 181
    Case 6
        GP22Y = 161
    Case 7
        GP22Y = 160
    Case 8
        GP22Y = 159
    Case 9
        GP22Y = 157
    Case 10
        GP22Y = 157
    Case 11
        GP22Y = 158
    Case 12
        GP22Y = 181
    Case 13
        GP22Y = 182
    Case 14
        GP22Y = 182
    Case 15
        GP22Y = 181
    Case 16
        GP22Y = 163
    Case 17
        GP22Y = 162
    Case 18
        GP22Y = 159
    Case 19
        GP22Y = 157
    Case 20
        GP22Y = 157
    Case 21
        GP22Y = 158
    Case 22
        GP22Y = 161
    Case 23
        GP22Y = 181
    Case 24
        GP22Y = 182
    Case 25
        GP22Y = 182
    Case 26
        GP22Y = 181
    Case 27
        GP22Y = 160
    Case 28
        GP22Y = 159
    Case 29
        GP22Y = 154
    Case 30
        GP22Y = 150
    Case 31
        GP22Y = 149
    Case 32
        GP22Y = 149
    Case 33
        GP22Y = 150
    Case 34
        GP22Y = 152
    Case 35
        GP22Y = 150
    Case 36
        GP22Y = 149
    Case 37
        GP22Y = 149
    Case 38
        GP22Y = 150
    Case 39
        GP22Y = 151
    Case 40
        GP22Y = 149
    End Select
End Function
Private Function GP23X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP23X = 226
    Case 1
        GP23X = 234
    Case 2
        GP23X = 236
    Case 3
        GP23X = 238
    Case 4
        GP23X = 247
    Case 5
        GP23X = 249
    Case 6
        GP23X = 247
    Case 7
        GP23X = 247
    Case 8
        GP23X = 245
    Case 9
        GP23X = 244
    Case 10
        GP23X = 239
    Case 11
        GP23X = 237
    Case 12
        GP23X = 236
    Case 13
        GP23X = 236
    Case 14
        GP23X = 235
    Case 15
        GP23X = 226
    Case 16
        GP23X = 225
    End Select
End Function
Private Function GP23Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP23Y = 149
    Case 1
        GP23Y = 149
    Case 2
        GP23Y = 151
    Case 3
        GP23Y = 149
    Case 4
        GP23Y = 149
    Case 5
        GP23Y = 151
    Case 6
        GP23Y = 156
    Case 7
        GP23Y = 157
    Case 8
        GP23Y = 159
    Case 9
        GP23Y = 158
    Case 10
        GP23Y = 158
    Case 11
        GP23Y = 160
    Case 12
        GP23Y = 167
    Case 13
        GP23Y = 181
    Case 14
        GP23Y = 182
    Case 15
        GP23Y = 182
    Case 16
        GP23Y = 181
    End Select
End Function
Private Function GP24X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP24X = 110
    Case 1
        GP24X = 111
    Case 2
        GP24X = 111
    Case 3
        GP24X = 110
    Case 4
        GP24X = 104
    Case 5
        GP24X = 103
    Case 6
        GP24X = 106
    Case 7
        GP24X = 106
    Case 8
        GP24X = 109
    End Select
End Function
Private Function GP24Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP24Y = 155
    Case 1
        GP24Y = 156
    Case 2
        GP24Y = 164
    Case 3
        GP24Y = 165
    Case 4
        GP24Y = 165
    Case 5
        GP24Y = 164
    Case 6
        GP24Y = 161
    Case 7
        GP24Y = 160
    Case 8
        GP24Y = 157
    End Select
End Function
Private Function GP25X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP25X = 204
    Case 1
        GP25X = 205
    Case 2
        GP25X = 205
    Case 3
        GP25X = 204
    Case 4
        GP25X = 198
    Case 5
        GP25X = 197
    Case 6
        GP25X = 199
    Case 7
        GP25X = 199
    Case 8
        GP25X = 202
    Case 9
        GP25X = 202
    End Select
End Function
Private Function GP25Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP25Y = 155
    Case 1
        GP25Y = 155
    Case 2
        GP25Y = 164
    Case 3
        GP25Y = 165
    Case 4
        GP25Y = 165
    Case 5
        GP25Y = 164
    Case 6
        GP25Y = 162
    Case 7
        GP25Y = 161
    Case 8
        GP25Y = 158
    Case 9
        GP25Y = 157
    End Select
End Function
Private Function GP26X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP26X = 262
    Case 1
        GP26X = 268
    Case 2
        GP26X = 270
    Case 3
        GP26X = 270
    Case 4
        GP26X = 271
    Case 5
        GP26X = 270
    Case 6
        GP26X = 270
    Case 7
        GP26X = 267
    Case 8
        GP26X = 263
    Case 9
        GP26X = 260
    Case 10
        GP26X = 260
    End Select
End Function
Private Function GP26Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP26Y = 163
    Case 1
        GP26Y = 163
    Case 2
        GP26Y = 165
    Case 3
        GP26Y = 167
    Case 4
        GP26Y = 168
    Case 5
        GP26Y = 169
    Case 6
        GP26Y = 171
    Case 7
        GP26Y = 174
    Case 8
        GP26Y = 174
    Case 9
        GP26Y = 171
    Case 10
        GP26Y = 165
    End Select
End Function
Private Function GP27X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP27X = 46
    Case 1
        GP27X = 47
    Case 2
        GP27X = 48
    Case 3
        GP27X = 48
    Case 4
        GP27X = 49
    Case 5
        GP27X = 298
    Case 6
        GP27X = 299
    Case 7
        GP27X = 299
    Case 8
        GP27X = 301
    Case 9
        GP27X = 323
    Case 10
        GP27X = 324
    Case 11
        GP27X = 345
    Case 12
        GP27X = 315
    Case 13
        GP27X = 314
    Case 14
        GP27X = 300
    Case 15
        GP27X = 299
    Case 16
        GP27X = 299
    Case 17
        GP27X = 48
    Case 18
        GP27X = 48
    Case 19
        GP27X = 47
    Case 20
        GP27X = 29
    Case 21
        GP27X = 28
    Case 22
        GP27X = 2
    Case 23
        GP27X = 20
    Case 24
        GP27X = 21
    End Select
End Function
Private Function GP27Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP27Y = 190
    Case 1
        GP27Y = 190
    Case 2
        GP27Y = 191
    Case 3
        GP27Y = 209
    Case 4
        GP27Y = 210
    Case 5
        GP27Y = 210
    Case 6
        GP27Y = 209
    Case 7
        GP27Y = 190
    Case 8
        GP27Y = 190
    Case 9
        GP27Y = 212
    Case 10
        GP27Y = 212
    Case 11
        GP27Y = 233
    Case 12
        GP27Y = 263
    Case 13
        GP27Y = 263
    Case 14
        GP27Y = 277
    Case 15
        GP27Y = 276
    Case 16
        GP27Y = 257
    Case 17
        GP27Y = 257
    Case 18
        GP27Y = 276
    Case 19
        GP27Y = 277
    Case 20
        GP27Y = 259
    Case 21
        GP27Y = 259
    Case 22
        GP27Y = 233
    Case 23
        GP27Y = 215
    Case 24
        GP27Y = 215
    End Select
End Function

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
    
    Call FormTopMost(Me.hWnd)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub
