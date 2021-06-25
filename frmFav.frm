VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmFav 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add a Bookmark"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser wbFav 
      Height          =   2415
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   4260
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtFavURL 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtFavName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Web Site Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Add new bookmark"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Dim strFav As String
 strFav = "f:\classwork\vb browser\links.htm"
 wbFav.Navigate strFav
End Sub

Public Sub populateBookmarks()


    
End Sub
