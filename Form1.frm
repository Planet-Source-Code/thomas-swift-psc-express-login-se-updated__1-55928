VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PSC Express Login SE"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   735
      TabIndex        =   21
      Top             =   1305
      Width           =   2880
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go To My Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   675
      TabIndex        =   20
      Top             =   3615
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log Into Desired PSC Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   675
      TabIndex        =   19
      Top             =   2865
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Log In And Upload Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   675
      TabIndex        =   18
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1065
      Left            =   158
      TabIndex        =   5
      Top             =   1740
      Width           =   4035
      Begin VB.OptionButton Option1 
         Caption         =   "LISP"
         Height          =   195
         Index           =   11
         Left            =   2550
         TabIndex        =   17
         Top             =   735
         Width           =   660
      End
      Begin VB.OptionButton Option1 
         Caption         =   ".Net"
         Height          =   195
         Index           =   10
         Left            =   1920
         TabIndex        =   16
         Top             =   720
         Width           =   660
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cold Fusion"
         Height          =   195
         Index           =   9
         Left            =   780
         TabIndex        =   15
         Top             =   705
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PHP"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   14
         Top             =   705
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Front Door"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   195
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Visual Basic"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   12
         Top             =   195
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Java/Javascript"
         Height          =   195
         Index           =   2
         Left            =   2460
         TabIndex        =   11
         Top             =   195
         Width           =   1485
      End
      Begin VB.OptionButton Option1 
         Caption         =   "C/C++"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   450
         Width           =   810
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ASP"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   9
         Top             =   450
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SQL"
         Height          =   195
         Index           =   5
         Left            =   1620
         TabIndex        =   8
         Top             =   450
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Perl"
         Height          =   195
         Index           =   6
         Left            =   2295
         TabIndex        =   7
         Top             =   450
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Delphi"
         Height          =   195
         Index           =   7
         Left            =   2910
         TabIndex        =   6
         Top             =   450
         Width           =   780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   525
      Left            =   0
      TabIndex        =   4
      Top             =   3975
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   926
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Picture         =   "Form1.frx":08CA
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1413
            MinWidth        =   1413
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "8/31/2004"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1657
            MinWidth        =   1657
            TextSave        =   "3:38 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   750
      TabIndex        =   0
      Top             =   240
      Width           =   2790
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   735
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   780
      Width           =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1815
      TabIndex        =   22
      Top             =   1095
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Email Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1373
      TabIndex        =   3
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1635
      TabIndex        =   2
      Top             =   570
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private MyIndex As Integer
Private Sub Command1_Click()
Text1.SetFocus
If Text1.Text = "" Or Text2.Text = "" Or Text2.Text = "" Then MsgBox "You havent filled in all the required feilds !": Exit Sub
If MyIndex = 0 Then
ShellExecute Me.hwnd, "open", "https://www.rentacoder.com/ads/authentication/LoginAction.asp?txtReturnURL=http://planet-source-code.com/vb/authentication/Login.asp?txtReturnURL=http://planet-source-code.com&txtEmailAddress=" & Text1.Text & "&txtPassword=" & Text2.Text, vbNullString, vbNullString, 1
Else
ShellExecute Me.hwnd, "open", "https://www.rentacoder.com/ads/authentication/LoginAction.asp?txtReturnURL=http://planet-source-code.com/vb/authentication/Login.asp?txtReturnURL=http://planet-source-code.com/vb/default.asp%26lngWId%3D" & MyIndex & "&lngWId=1&blnOutsideOfVBSubWeb=False&txtEmailAddress=" & Text1.Text & "&txtPassword=" & Text2.Text, vbNullString, vbNullString, 1
End If
Call SaveSettings
End
End Sub
Private Sub Command2_Click()
Text1.SetFocus
If Text1.Text = "" Or Text2.Text = "" Or Text2.Text = "" Then MsgBox "You havent filled in all the required feilds !": Exit Sub
ShellExecute Me.hwnd, "open", "https://www.rentacoder.com/ads/authentication/LoginAction.asp?txtReturnURL=http://www.planet-source-code.com/vb/authors/determine_author_type.asp?lngWId=1&lngWId=1&blnOutsideOfVBSubWeb=False&strPassKey=&txtEmailAddress=" & Text1.Text & "&txtPassword=" & Text2.Text, vbNullString, vbNullString, 1
Call SaveSettings
End
End Sub
Private Sub Command3_Click()
Text1.SetFocus
If Text1.Text = "" Or Text2.Text = "" Or Text2.Text = "" Then MsgBox "You havent filled in all the required feilds !": Exit Sub
ShellExecute Me.hwnd, "open", "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&txtCriteria=" & fConvert(Text3.Text) & "&lngWId=1&B1=Quick+Search", vbNullString, vbNullString, 1
Call SaveSettings
End
End Sub
Private Sub Form_Load()
If App.PrevInstance Then End
btnFlat Command1
btnFlat Command2
btnFlat Command3
Text1.Text = GetSetting("PSC Express Login", "Settings", "User ID", "")
Text2.Text = GetSetting("PSC Express Login", "Settings", "Password", "")
Text3.Text = GetSetting("PSC Express Login", "Settings", "Name", "")
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub SaveSettings()
SaveSetting "PSC Express Login", "Settings", "User ID", Text1.Text
SaveSetting "PSC Express Login", "Settings", "Password", Text2.Text
SaveSetting "PSC Express Login", "Settings", "Name", Text3.Text
End Sub
Function btnFlat(Button As CommandButton)
SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
Private Sub Option1_Click(Index As Integer)
Text1.SetFocus
On Error Resume Next
Dim GetMyIndex As Integer
Dim x As Integer
For x = 0 To 13
If Option1(x).Value = True Then
If x = 11 Then x = 13
MyIndex = x
Exit Sub
End If
Next x
End Sub
Public Function fConvert(ByVal sStr As String) As String
    Dim I As Integer
    Dim sBadChar As String
    Dim sNewChar As String
    Dim NewString As String
    Dim NewString2 As String
    'List all illegal / unwanted characters
    sBadChar = Chr(32)
    'Loop through all the characters of the
    '     string
    'checking whether each is an illegal cha
    '     racter
    sNewChar = "+"
    Dim m As Integer
    Dim GetChr0 As String
    Dim GetChr1 As String
    For m = 1 To Len(sStr)
        GetChr0 = Left(sStr, m)
        GetChr1 = Right(GetChr0, 1)
        If GetChr1 = sBadChar Then
        fConvert = fConvert & sNewChar
        Else
        fConvert = fConvert & GetChr1
        End If
    Next m

End Function
