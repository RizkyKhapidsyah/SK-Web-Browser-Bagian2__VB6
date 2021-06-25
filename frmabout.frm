VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Browser created using Microsoft Visual Basic."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please visit my website at www.altenageimpact.com"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2001 Keith Miller/AltenageImpact Design Group"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Good Old Fashion Browser created by Keith Miller"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4080
      Left            =   0
      Picture         =   "frmabout.frx":0000
      Top             =   0
      Width           =   5430
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
