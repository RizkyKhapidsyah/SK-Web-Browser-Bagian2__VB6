VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmSettings 
   Caption         =   "You Internet Settings"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdGetProp 
      Caption         =   "Properties"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser wbStatus 
      Height          =   1815
      Left            =   5040
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
      ExtentX         =   873
      ExtentY         =   3201
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
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
      Location        =   "http:///"
   End
   Begin VB.TextBox txtHomepage 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Homepage:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetProp_Click()
'wbStatus.GetProperty
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    wbStatus.GoHome
End Sub

Private Sub wbStatus_StatusTextChange(ByVal Text As String)
    txtHomepage.Text = wbStatus.LocationURL
    'txtName.Text = wbStatus.QueryStatusWB
End Sub
