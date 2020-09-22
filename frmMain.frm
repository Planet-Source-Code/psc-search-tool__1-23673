VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Search Tool"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSort 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtHits 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox cmbApp 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Created by Mike Rossi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   2520
      Picture         =   "frmMain.frx":0000
      Top             =   960
      Width           =   1965
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Sort Options:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Number of hits to return:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Application:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Data to search for:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reg As New clsRegistry

Private Sub cmdExit_Click()

  Unload Me
  
End Sub

Private Sub cmdSearch_Click()

Dim strURL As String
  
  If cmbApp.Text = "" Then cmbApp.Text = "Visual Basic"
  If cmbSort.Text = "" Then cmbSort.Text = "Alphabetical"
  If txtHits.Text = "" Then txtHits.Text = "20"
  
  If txtSearch.Text = "" Then
    MsgBox "Enter criteria to search for!", vbCritical, "Error"
    Exit Sub
  ElseIf IsNumeric(txtHits.Text) = False Then
    MsgBox "Enter a valid number in the 'Hits' text box!", vbCritical, "Error"
    Exit Sub
  End If

  strURL = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?blnResetAllVariables=TRUE&cmSearch=Search"
  strURL = strURL & "&lngWId=" & appID(cmbApp.Text)
  strURL = strURL & "&txtMaxNumberOfEntriesPerPage=" & txtHits.Text
  strURL = strURL & "&optSort=" & sortID(cmbSort.Text)
  strURL = strURL & "&txtCriteria=" & FormatSearchText(txtSearch.Text)
  
  hypURL frmMain, strURL
  
  reg.SaveSettingString Local_Machine, "Software\Rossi\PSCSearch", "App", cmbApp.Text
  reg.SaveSettingString Local_Machine, "Software\Rossi\PSCSearch", "Sort", cmbSort.Text
  reg.SaveSettingString Local_Machine, "Software\Rossi\PSCSearch", "Hits", txtHits.Text
  
End Sub

Public Function FormatSearchText(SearchString As String) As String

Dim strSearch As String
Dim x As Integer

  strSearch = SearchString
  
  strSearch = Replace(strSearch, " ", "+")
  strSearch = Replace(strSearch, Chr(34), "%22")
  FormatSearchText = strSearch

End Function

Public Function appID(appName As String) As Long

  Select Case appName
    Case "ASP/VBScript"
      appID = 4
    Case "Visual Basic"
      appID = 1
    Case "C++/C"
      appID = 3
    Case "Javascript/Java"
      appID = 2
    Case "Perl"
      appID = 6
    Case "Delphi"
      appID = 7
    Case "PHP"
      appID = 8
    Case "SQL"
      appID = 5
  End Select

End Function

Public Function sortID(sortName As String) As String

  Select Case sortName
    Case "Alphabetical"
      sortID = "Alphabetical"
    Case "Newest"
      sortID = "DateDescending"
    Case "Oldest"
      sortID = "DateAscending"
    Case "Most Popular"
      sortID = "CountDescending"
  End Select

End Function
Private Sub Form_Load()

  With cmbApp
    .AddItem "Visual Basic"
    .AddItem "ASP/VBScript"
    .AddItem "C++/C"
    .AddItem "Javascript/Java"
    .AddItem "Perl"
    .AddItem "Delphi"
    .AddItem "PHP"
    .AddItem "SQL"
  End With
  
  With cmbSort
    .AddItem "Alphabetical"
    .AddItem "Newest"
    .AddItem "Oldest"
    .AddItem "Most Popular"
  End With
  
  cmbApp.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\PSCSearch", "App", "Visual Basic")
  cmbSort.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\PSCSearch", "Sort", "Alphabetical")
  txtHits.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\PSCSearch", "Hits", "10")
  
End Sub

