VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Paint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   Caption         =   "Spaint v1.0a"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   6
      Left            =   120
      Picture         =   "Paint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Eraser"
      Top             =   120
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   5
      Left            =   120
      Picture         =   "Paint.frx":05C2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Close Loop Pen"
      Top             =   1200
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdlPaint 
      Left            =   120
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "Paint.frx":0AB4
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Filled Rectangle"
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "Paint.frx":0FA6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Hollow Rectangle"
      Top             =   1560
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "Paint.frx":1498
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Pencil"
      Top             =   480
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "Paint.frx":198A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Line"
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Paint.frx":1E7C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Eraser"
      Top             =   2760
      Width           =   375
   End
   Begin VB.PictureBox picPaint 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   720
      ScaleHeight     =   7755
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   0
      Width           =   11175
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   5400
      Shape           =   2  'Oval
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   23
      Top             =   6315
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   22
      Top             =   6315
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   6060
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   16
      Top             =   6060
      Width           =   255
   End
   Begin VB.Label lblCurrentColour 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Height          =   135
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   4245
      Width           =   375
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Height          =   105
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   4065
      Width           =   375
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3930
      Width           =   375
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3810
      Width           =   375
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblColour 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   5550
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   5295
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   5295
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   5
      Top             =   5805
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Top             =   5805
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label lblColour 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   5550
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu subLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu subSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu subExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu subColour 
         Caption         =   "Colours"
         Begin VB.Menu mnuColour 
            Caption         =   "Red"
            Index           =   0
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "Blue"
            Index           =   1
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "Black"
            Index           =   2
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "White"
            Index           =   3
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "Cyan"
            Index           =   4
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "Magenta"
            Index           =   5
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "Green"
            Index           =   6
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuColour 
            Caption         =   "Yellow"
            Index           =   7
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu subLineWidth 
         Caption         =   "Line Width"
         Begin VB.Menu mnuLine 
            Caption         =   "1"
            Index           =   0
         End
         Begin VB.Menu mnuLine 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu mnuLine 
            Caption         =   "4"
            Index           =   2
         End
         Begin VB.Menu mnuLine 
            Caption         =   "8"
            Index           =   3
         End
         Begin VB.Menu mnuLine 
            Caption         =   "10"
            Index           =   4
         End
         Begin VB.Menu mnuLine 
            Caption         =   "16"
            Index           =   5
         End
         Begin VB.Menu mnuLine 
            Caption         =   "20"
            Index           =   6
         End
         Begin VB.Menu mnuLine 
            Caption         =   "32"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu subIndex 
         Caption         =   "Help Index"
         Shortcut        =   {F1}
      End
      Begin VB.Menu subAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Paint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X1 As Single
Dim Y1 As Single
Private MyColour     As Long
Private MyTool       As Long
Enum Tools
TEraser
Tpencil
TLine
TBox
TFilledBox
tloop
TFan
TCircle
End Enum
Private MyLineWidth       As Long
Private OldX         As Integer
Private OldY         As Integer


Private Sub DisplayColour(ByVal intIndex As Integer)

  Select Case intIndex
   Case 0
    MyColour = vbBlack
   Case 1
    MyColour = vbBlue
   Case 2
    MyColour = vbGreen
   Case 3
    MyColour = vbMagenta
   Case 4
    MyColour = vbRed
   Case 5
    MyColour = vbYellow
   Case 6
    MyColour = vbWhite
   Case 7
    MyColour = vbCyan
   Case 8
    MyColour = &H800000
   Case 9
    MyColour = &HC0C0FF
   Case 10
    MyColour = &H80FF&
   Case 11
    MyColour = &HC0FFFF
  End Select
  lblCurrentColour.BackColor = MyColour

End Sub

Private Sub Form_Load()
Paint.WindowState = vbMaximized
  DisplayColour 0
  SetTool 2
  SetLine 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Paint.Caption = "Spaint v1.0a*" Then
 With cdlPaint
    .Filter = "BMP Files|*.bmp"
    .FilterIndex = 1
    .ShowSave
    If LenB(.FileName) Then
      SavePicture picPaint.Image, .FileName
    End If
  End With
End If
End Sub

Private Sub lblColour_Click(index As Integer)

  DisplayColour index

End Sub

Private Sub lblWidth_Click(index As Integer)

  SetLine index

End Sub

Private Sub mnuColour_Click(index As Integer)

  Select Case index
   Case 0
    DisplayColour 4
   Case 1
    DisplayColour 1
   Case 2
    DisplayColour 0
   Case 3
    DisplayColour 6
   Case 4
    DisplayColour 7
   Case 5
    DisplayColour 3
   Case 6
    DisplayColour 2
   Case 7
    DisplayColour 5
  End Select

End Sub

Private Sub mnuLine_Click(index As Integer)

  SetLine index

End Sub

Private Sub optTool_Click(index As Integer)

  SetTool index

End Sub





Private Sub picPaint_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
X1 = X
Y1 = Y
OldX = X
OldY = Y
  picPaint.CurrentX = X
  picPaint.CurrentY = Y
  Paint.Caption = "Spaint v1.0a*"
End Sub

Private Sub picPaint_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

  If Button = 1 Then
    Select Case MyTool
     Case TLine, tloop
      picPaint.Line (picPaint.CurrentX, picPaint.CurrentY)-(X, Y), MyColour
     Case TEraser
      picPaint.DrawWidth = 32
      picPaint.Line (picPaint.CurrentX, picPaint.CurrentY)-(X, Y), vbWhite
      picPaint.DrawWidth = MyLineWidth
      Case TFan
      picPaint.Line (X, Y)-(picPaint.CurrentX, picPaint.CurrentY), MyColour

    End Select
  End If

End Sub

Private Sub picPaint_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

  If Button = 1 Then
    Select Case MyTool
    Case tloop
      picPaint.Line (X, Y)-(OldX, OldY), MyColour
     Case Tpencil
      picPaint.Line (OldX, OldY)-(X, Y), MyColour
     Case TBox
      picPaint.Line (OldX, OldY)-(X, Y), MyColour, B
     Case TFilledBox
      picPaint.Line (OldX, OldY)-(X, Y), MyColour, BF

    End Select
  End If

End Sub

Private Sub SetLine(ByVal intIndex As Integer)

  Dim I As Long
  For I = 0 To 5
   lblWidth(I).BorderStyle = IIf(intIndex = I, 1, 0)
  Next
  Select Case intIndex
   Case 0
    MyLineWidth = 1
   Case 1
    MyLineWidth = 2
   Case 2
    MyLineWidth = 4
   Case 3
    MyLineWidth = 8
   Case 4
    MyLineWidth = 10
   Case 5
    MyLineWidth = 12
  End Select
picPaint.DrawWidth = MyLineWidth
End Sub

Private Sub SetTool(ByVal intIndex As Integer)

  MyTool = intIndex

End Sub

Private Sub subAbout_Click()

  frmAbout.Show

End Sub

Private Sub subExit_Click()
If Paint.Caption = "Spaint v1.0a*" Then
With cdlPaint
    .Filter = "BMP Files|*.bmp"
    .FilterIndex = 1
    .ShowSave
    If LenB(.FileName) Then
      SavePicture picPaint.Image, .FileName
    End If
    End
End With
Else
  End
  End If
End Sub

Private Sub subIndex_Click()

  Help.Show

End Sub

Private Sub subLoad_Click()

  With cdlPaint
    .Filter = "BMP Files|*.bmp"
    .FilterIndex = 1
    .ShowOpen
    If LenB(.FileName) Then
      picPaint = LoadPicture(.FileName)
    End If
  End With

End Sub

Private Sub subSave_Click()

  With cdlPaint
    .Filter = "BMP Files|*.bmp"
    .FilterIndex = 1
    .ShowSave
    If LenB(.FileName) Then
      SavePicture picPaint.Image, .FileName
    End If
  End With

End Sub



