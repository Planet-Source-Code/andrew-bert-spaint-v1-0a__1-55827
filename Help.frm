VERSION 5.00
Begin VB.Form Help 
   Caption         =   "Spaint v1.0a Help"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstHelp 
      BackColor       =   &H00C0FFFF&
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblHelp_2 
      Caption         =   "Spaint v1.0a Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label lblHelpDisplay 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00404040&
      Height          =   2775
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  With lstHelp
    .AddItem "Drawing"
    .AddItem "Pecil Tool"
    .AddItem "Line Tool"
    .AddItem "Eraser Tool"
    .AddItem "Colours"
    .AddItem "How to Get Started with Spaint"
    .AddItem "Short Cut Keys"
    .AddItem "About Spaint v1.0"
    .AddItem "MODIFICATIONS"
  End With 'List1
  'Help.WindowState = vbMaximized

End Sub

Private Sub lstHelp_Click()
Dim strMsg As String
  Select Case lstHelp.ListIndex
   Case 0
    strMsg = "To Draw using Spaint v1.0 just choose which tool you would like to use (see tools) and then Draw on the White Box. You can use a few different colours by clicking on the colour you would like to use.(see Colours)"
   Case 1
    strMsg = "This Tool is just like a normal pencil, just move it around the box and you will be able to free draw. You must have the left mouse button pressed down to be able to draw."
   Case 2
    strMsg = "This tool allows the user to just draw a straight line. Press the left mouse button down for when you want to start the line, and release it when you want the line to end."
   Case 3
    strMsg = "This tool just erases where you click by drawing white over it."
   Case 4
    strMsg = "In Spaint v1.0 you are able to use 8 different colours to draw with, these are, Red, Blue, Black, White, Green, Yellow, Cyan and Magenta."
   Case 5
    strMsg = "To get Started just choose what you want to draw and draw it."
   Case 6
    strMsg = "In Spaint there are some short cut keys to make things quicker. Red = F1, Blue = F2, Black = F3, White = F4, Cyan = F5, Magenta = F6, Green = F7, Yellow = F8 and Exit = Ctrl + X."
   Case 7
    strMsg = "Spaint v1.0 was created by Andrew Bertuch it took him about three weeks to make and had over 1000 lines of code in it. It was made using Microsoft Visual Basic 6.0."
    Case 8
    strMsg = "Modifications by Roger Gilchrist. Converted to use Control arrays and simplified variables. Added Hollow/Filled Box Tools (replaced Box Tool with seperate setting for hollow/filled) and the Close Loop Pen. Added Commondialogbox and Tool ToolTips"
  End Select
lblHelpDisplay.Caption = strMsg
End Sub

