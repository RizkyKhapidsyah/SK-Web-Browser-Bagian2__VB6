VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Browser 
   Caption         =   "VB Web Browser"
   ClientHeight    =   8925
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10920
   Icon            =   "Browser5.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1482
      ButtonWidth     =   1588
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "H&ome"
            Key             =   "tlbHome"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "B&ack"
            Key             =   "tlbBack"
            Object.ToolTipText     =   "Back in history"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "For&ward"
            Key             =   "tlbForward"
            Object.ToolTipText     =   "Forward in history"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reload"
            Key             =   "tlbReload"
            Object.ToolTipText     =   "Reload current page"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "tlbStop"
            Object.ToolTipText     =   "Stop current download"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sear&ch"
            Key             =   "tlbSearch"
            Object.ToolTipText     =   "Search the web"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7200
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":19C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":269A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":3374
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":404E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":4D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":5A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":66DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":73B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":8090
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":8D6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8550
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   12356
            MinWidth        =   12349
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   8775
   End
   Begin SHDocVwCtl.WebBrowser wbBrowser 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   10695
      ExtentX         =   18865
      ExtentY         =   12515
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4860
      Top             =   4215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser5.frx":9A44
            Key             =   "homeIcon_off2"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgGo 
      Height          =   480
      Left            =   9840
      MouseIcon       =   "Browser5.frx":9D5E
      MousePointer    =   99  'Custom
      Picture         =   "Browser5.frx":A1A0
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuNewBrowser 
            Caption         =   "New Browser"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuHR3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuHR4 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu nmuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuHL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find on page"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuHTMLEdit 
         Caption         =   "View Source"
      End
      Begin VB.Menu mnuToolbars 
         Caption         =   "Toolbars"
         Begin VB.Menu mnuStandardTools 
            Caption         =   "Standard Tools"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLocationBar 
            Caption         =   "Location Bar"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuStatusBar 
            Caption         =   "Status Bar"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuHR5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDefault 
            Caption         =   "Default"
         End
      End
      Begin VB.Menu mnuISettings 
         Caption         =   "Internet Settings"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add a Bookmark"
      End
      Begin VB.Menu mnuViewBook 
         Caption         =   "V&iew my Bookmarks"
         Begin VB.Menu mnuMicro 
            Caption         =   "Microsoft"
         End
         Begin VB.Menu mnuNet 
            Caption         =   "Netscape"
         End
         Begin VB.Menu mnuMacro 
            Caption         =   "Macromedia"
         End
         Begin VB.Menu mnuAltenage 
            Caption         =   "AltenageImpact"
         End
         Begin VB.Menu mnuHL2 
            Caption         =   "-"
         End
         Begin VB.Menu searchengines 
            Caption         =   "Search Engines"
            Begin VB.Menu mnuGoogle 
               Caption         =   "Google"
            End
            Begin VB.Menu mnuGoNet 
               Caption         =   "GO Network"
            End
            Begin VB.Menu mnuISeek 
               Caption         =   "InfoSeek"
            End
            Begin VB.Menu mnuLycos 
               Caption         =   "Lycos"
            End
            Begin VB.Menu mnuYahoo 
               Caption         =   "Yahoo"
            End
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About GOF"
      End
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const conBrowser As String = " "
Dim shlHelper As ShellUIHelper
Dim strFavUrl As String
Dim strLocName As String, cap As String
Dim bolDefault As Boolean

Private Sub Form_Load()
Dim strHome As String

    wbBrowser.GoHome 'Loads the first page of the wbBrowser control to your Homepage
    wbBrowser.StatusBar = True
    wbBrowser.Height = (Browser.Height) - 2250
    wbBrowser.Width = (Browser.Width) - 100
    Set shlHelper = New ShellUIHelper
    

    If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
        Browser.mnuDefault.Enabled = True
    Else
        Browser.mnuDefault.Enabled = False
    End If

End Sub

Private Sub Form_Resize()
On Error GoTo resizeerror
'Scales the wbBrowser control to fit the screen when you maximize the browser
wbBrowser.Height = (Browser.Height) - 2250 'Set wbBrowser to browser form height - 2250
wbBrowser.Width = (Browser.Width) - 140 'Set wbBrowser to browser form width - 140
    If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
        bolDefault = True
    Else
        bolDefault = False
    End If
    If bolDefault = True Then
        mnuDefault.Enabled = False
    Else
        mnuDefault.Enabled = True
    End If
resizeerror:
End Sub

Private Sub mnuAbout_Click() 'Opens the "About GOF Browser" splash screen
 frmAbout.Visible = True
 frmAbout.Icon = Browser.Icon
'**************************
'** Begin dropdown menus **
'**************************
End Sub

Private Sub mnuCopy_Click()
wbBrowser.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuCut_Click()
wbBrowser.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuDefault_Click()
Dim bolDefault As Boolean
If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
    bolDefault = True
Else
    bolDefault = False
End If
If bolDefault = True Then
    mnuDefault.Enabled = False
Else
    mnuDefault.Enabled = True
    lblLocation.Visible = True
    txtAddress.Visible = True
    imgGo.Visible = True
    Toolbar1.Visible = True
    StatusBar1.Visible = True
    mnuStatusBar.Checked = True
    mnuLocationBar.Checked = True
    mnuStandardTools.Checked = True
End If
If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
    Browser.mnuDefault.Enabled = False
Else
    Browser.mnuDefault.Enabled = True
End If
End Sub

Private Sub mnuLocationBar_Click()
    If lblLocation.Visible = True Then
        lblLocation.Visible = False
        mnuLocationBar.Checked = False
    Else
        lblLocation.Visible = True
        mnuLocationBar.Checked = True
    End If
    If txtAddress.Visible = True Then
        txtAddress.Visible = False
    Else
        txtAddress.Visible = True
    End If
    If imgGo.Visible = True Then
        imgGo.Visible = False
    Else
        imgGo.Visible = True
    End If
    If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
        Browser.mnuDefault.Enabled = False
    Else
        Browser.mnuDefault.Enabled = True
    End If
End Sub

Private Sub mnuOpen_Click()
CommonDialog1.ShowOpen
wbBrowser.Navigate (CommonDialog1.FileName)
End Sub

Private Sub mnuPage_Click()
Dim eQuery As OLECMDF
    
    On Error Resume Next
    eQuery = wbBrowser.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            wbBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
        Else
            MsgBox "The Page Setup For Print Command is currently disabled."
        End If
    End If
End Sub

Private Sub mnuPrint_Click()
On Error GoTo printerror
wbBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
printerror:
End Sub

Private Sub mnuSave_Click()
wbBrowser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuSaveAs_Click()
wbBrowser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuSelectAll_Click()
wbBrowser.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuStandardTools_Click()
    If Toolbar1.Visible = True Then
        Toolbar1.Visible = False
        mnuStandardTools.Checked = False
    Else
        Toolbar1.Visible = True
        mnuStandardTools.Checked = True
    End If
    If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
        Browser.mnuDefault.Enabled = False
    Else
        Browser.mnuDefault.Enabled = True
    End If
End Sub

Private Sub mnuStatusBar_Click()
    If StatusBar1.Visible = True Then
        StatusBar1.Visible = False
        mnuStatusBar.Checked = False
    Else
        StatusBar1.Visible = True
        mnuStatusBar.Checked = True
    End If
    If Toolbar1.Visible = True And StatusBar1.Visible = True And txtAddress.Visible = True Then
        Browser.mnuDefault.Enabled = False
    Else
        Browser.mnuDefault.Enabled = True
    End If
End Sub

Private Sub nmuPaste_Click()
wbBrowser.ExecWB OLECMDID_PASTESPECIAL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuGoNet_Click()
Dim strGoNet As String
    strGoNet = "www.go.com"
    wbBrowser.Navigate strGoNet
End Sub

Private Sub mnuGoogle_Click()
Dim strGoogle As String
    strGoogle = "www.google.com"
    wbBrowser.Navigate strGoogle
End Sub

Private Sub mnuHTMLEdit_Click()
Dim strEditTitleNew As String
Const strEditTitle As String = " Source for: "

strEditTitleNew = wbBrowser.LocationName
On Error GoTo editerror
    frmEdit.txtEdit.Text = wbBrowser.Document.documentElement.outerHTML
    frmEdit.Caption = cap & "-Design Web Explorer"
    frmEdit.lblEdit.Caption = strEditTitle & strEditTitleNew
    frmEdit.Icon = Browser.Icon
    frmEdit.Show
editerror:
End Sub

Private Sub mnuISeek_Click()
Dim strISeek As String
    strISeek = "www.infoseek.com"
    wbBrowser.Navigate strISeek
End Sub

Private Sub mnuISettings_Click()
'frmSettings.Visible = True
End Sub

Private Sub mnuLycos_Click()
Dim strLycos As String
    strLycos = "www.lycos.com"
    wbBrowser.Navigate strLycos
End Sub

Private Sub mnuYahoo_Click()
Dim strYahoo As String
    strYahoo = "www.yahoo.com"
    wbBrowser.Navigate strYahoo
End Sub
Private Sub mnuAdd_Click()
On Error GoTo eaddfav
    strLocName = Browser.Caption
    strFavUrl = txtAddress.Text
    shlHelper.AddFavorite strFavUrl, strLocName
eaddfav:
    
End Sub

Private Sub mnuAltenage_Click()
Const conAltURL As String = "www.altenageimpact.com"
    wbBrowser.Navigate conAltURL
End Sub


Private Sub mnuMacro_Click()
Const conMacURL As String = "www.macromedia.com"
    wbBrowser.Navigate conMacURL
End Sub


Private Sub mnuMicro_Click()
Const conMicURL As String = "www.microsoft.com"
    wbBrowser.Navigate conMicURL
End Sub


Private Sub mnuNet_Click()
Const conNetURL As String = "www.netscape.com"
    wbBrowser.Navigate conNetURL
End Sub

Private Sub exit_Click()
 End
 '************************
'** End dropdown menus **
'************************
End Sub

'Handles the toolbar buttons
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
Dim strSearch As String
    On Error Resume Next
    Select Case Button.Key
        Case "tlbHome"
        wbBrowser.GoHome
        
        Case "tlbBack"
        wbBrowser.GoBack
        
        Case "tlbForward"
        wbBrowser.GoForward
        
        Case "tlbReload"
        wbBrowser.Refresh
        
        Case "tlbStop"
        wbBrowser.Stop
        
        Case "tlbSearch"
            strSearch = "www.yahoo.com"
            wbBrowser.Navigate strSearch
    End Select
End Sub
'Handles the GO button next to the address/location bar
Private Sub imgGo_Click()
Dim strString As String
 strString = txtAddress.Text
 wbBrowser.Navigate strString
 
End Sub
'This lets the user press ENTER after entering a URL into the txtAddress textbox
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
Dim strSearch1 As String
Dim strLocName As String

    If KeyAscii = 13 Then 'The enter button is 13 in Ascii
        strSearch1 = txtAddress.Text
        wbBrowser.Navigate strSearch1
    End If
    strLocName = wbBrowser.LocationName 'Set object strLocName to the value of <TITLE> in a HTML doc
        Browser.Caption = strLocName & conBrowser 'Changes the Browsers Form caption to strLocName and adds to Const conBrowser
End Sub

Private Sub wbBrowser_StatusTextChange(ByVal Text As String)
'Changes the text in the txtAddress to the current URL of page
'Changes the caption of the Browser Form to the current value of <TITLE>
    If Text <> "Done" Then
        StatusBar1.Panels(1).Text = Text
    End If
    
    strLocName = wbBrowser.LocationName
        Browser.Caption = strLocName & conBrowser
        
    txtAddress.Text = wbBrowser.LocationURL
    

End Sub
