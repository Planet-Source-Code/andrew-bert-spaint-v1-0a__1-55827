VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spaint"
   ClientHeight    =   3045
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2101.713
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Line linAbout 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.582
      Y2              =   1687.582
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":030A
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Spaint"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line linAbout 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:1"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Reg Key Security Options...
Private Const READ_CONTROL                 As Long = &H20000
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_SET_VALUE                As Long = &H2
Private Const KEY_CREATE_SUB_KEY           As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_CREATE_LINK              As Long = &H20

''Private Const KEY_ALL_ACCESS               As Double = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
' Reg Key ROOT Types...
''Private Const HKEY_LOCAL_MACHINE           As Long = &H80000002
''Private Const ERROR_SUCCESS                As Integer = 0
''Private Const REG_SZ                       As Integer = 1    ' Unicode nul terminated string
''Private Const REG_DWORD                    As Integer = 4    ' 32-bit number
''Private Const gREGKEYSYSINFOLOC            As String = "SOFTWARE\Microsoft\Shared Tools Location"
''Private Const gREGVALSYSINFOLOC            As String = "MSINFO"
''Private Const gREGKEYSYSINFO               As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
''Private Const gREGVALSYSINFO               As String = "PATH"
''Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
''Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
''Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Sub cmdOK_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  Me.Caption = "About Spaint "
  lblVersion.Caption = "1.0"
  lblTitle.Caption = "Spaint"

End Sub

