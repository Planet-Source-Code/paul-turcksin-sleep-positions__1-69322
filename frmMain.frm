VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tell me how you sleep ..."
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9180
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Want to know more? Click here."
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   9255
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Move the mouse over your sleeping position."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   3600
      TabIndex        =   0
      Top             =   3360
      Width           =   5295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'............................ DC
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private OriginalDC As Long
Private HighlightedDC As Long

'............................ BITMAP
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'............................ OBJECT
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private oldOriginalObj As Long
Private oldHighlightedObj As Long

'............................ RECT
Private Type RECT
  iLeft As Long
  iTop As Long
  iRight As Long
  iBottom As Long
End Type
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long

Private arRect(5) As RECT
Private oldSelected As Integer
Private arWidth(5) As Long

Private stdPict101 As StdPicture
Private stdPict102 As StdPicture

Private Sub Form_Load()
   Dim lWidth As Long
   Dim lHeight As Long

' get "original" picture from resource file and show it in the form
   OriginalDC = CreateCompatibleDC(Me.hdc)
   Set stdPict101 = LoadResPicture(101, vbResBitmap)
   oldOriginalObj = SelectObject(OriginalDC, stdPict101.Handle)
   lWidth = Round(ScaleX(stdPict101.Width, vbHimetric, vbPixels))
   lHeight = Round(ScaleY(stdPict101.Height, vbHimetric, vbPixels))
   BitBlt Me.hdc, 0, 0, lWidth, lHeight, OriginalDC, 0, 0, vbSrcCopy
   Me.Refresh
   
' get "highlighted" picture
   HighlightedDC = CreateCompatibleDC(Me.hdc)
   Set stdPict102 = LoadResPicture(102, vbResBitmap)
   oldHighlightedObj = SelectObject(HighlightedDC, stdPict102.Handle)
   
' set up rectangles defining individual pictures
   With arRect(0): .iLeft = 0:   .iTop = 0: .iBottom = 217: .iRight = 114: End With
   With arRect(1): .iLeft = 115: .iTop = 0: .iBottom = 217: .iRight = 220: End With
   With arRect(2): .iLeft = 221: .iTop = 0: .iBottom = 217: .iRight = 307: End With
   With arRect(3): .iLeft = 308: .iTop = 0: .iBottom = 217: .iRight = 410: End With
   With arRect(4): .iLeft = 411: .iTop = 0: .iBottom = 217: .iRight = 517: End With
   With arRect(5): .iLeft = 518: .iTop = 0: .iBottom = 217: .iRight = 612: End With

' set up width of individual pictures
   arWidth(0) = 114
   arWidth(1) = 116
   arWidth(2) = 87
   arWidth(3) = 103
   arWidth(4) = 107
   arWidth(5) = 105
   
  oldSelected = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim iSelected As Integer
   Dim i As Integer
   
' is the mouse over a rectangle?
   iSelected = -1
   For i = 0 To 5
      If PtInRect(arRect(i), CLng(X), CLng(Y)) <> 0 Then
         iSelected = i
         Exit For
      End If
   Next i
   
' over same as previous rectangle?
   If iSelected = oldSelected Then
      Exit Sub
   End If
   
' remove highligth
   If oldSelected <> -1 Then
      BitBlt Me.hdc, arRect(oldSelected).iLeft, 0, arWidth(oldSelected), 217, OriginalDC, arRect(oldSelected).iLeft, 0, vbSrcCopy
      Me.Refresh
      oldSelected = -1
      lblType = ""
      lblMessage.ForeColor = &HE0E0E0
      lblMessage = "Move the mouse over your sleeping position."
      End If
      
 ' new highligth
    If iSelected <> -1 Then
      BitBlt Me.hdc, arRect(iSelected).iLeft, 0, arWidth(iSelected), 217, HighlightedDC, arRect(iSelected).iLeft, 0, vbSrcCopy
      Me.Refresh
      
      lblMessage.ForeColor = vbYellow
      Select Case iSelected
         Case 0
            lblType = "The Yearner"
            lblMessage = "A suspicious person with a very rational approach to life."
         Case 1
            lblType = "The Starfish"
            lblMessage = "A good listener who likes to help whenever needed."
         Case 2
            lblType = "The Log"
            lblMessage = "Easy going and social, but can be seen as too gulible."
         Case 3
            lblType = "The Soldier"
            lblMessage = "Quiet and reserved who loathes noisy social scenes."
         Case 4
            lblType = "The Freefaller"
            lblMessage = "Appears brash but cannot cope with personal criticism."
         Case 5
            lblType = "The Foetus"
            lblMessage = "Seems tough but is really a sensitive, shy person."
      End Select
      oldSelected = iSelected
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SelectObject OriginalDC, oldOriginalObj
   DeleteDC OriginalDC
   Set stdPict101 = Nothing
   SelectObject HighlightedDC, oldHighlightedObj
   DeleteDC HighlightedDC
   Set stdPict102 = Nothing
   Set frmMain = Nothing
   
End Sub

Private Sub lblAbout_Click()
   frmAbout.Show vbModal
End Sub
