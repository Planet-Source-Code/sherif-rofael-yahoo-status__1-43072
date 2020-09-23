VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo!"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   1680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Yahoo! Id Here"
      Top             =   720
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      ExtentX         =   2990
      ExtentY         =   1296
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Online Or Offline"
      Height          =   615
      Left            =   0
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M As Boolean

Private Sub Command1_Click()
url = "http://opi.yahoo.com/online?u=" & Text1.Text
wb.Navigate url
End Sub

Private Sub Form_Load()
M = True
End Sub

Private Sub Text1_Click()
If M = True Then
Text1.Text = ""
M = False
End If
End Sub
