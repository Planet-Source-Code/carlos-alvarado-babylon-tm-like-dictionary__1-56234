VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmConsulta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Baby Dic"
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClipboard 
      Interval        =   2
      Left            =   120
      Top             =   750
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Translate"
      Default         =   -1  'True
      Height          =   375
      Left            =   3750
      TabIndex        =   2
      Top             =   30
      Width           =   975
   End
   Begin VB.TextBox txtWord 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   3585
   End
   Begin SHDocVwCtl.WebBrowser webBabylon 
      Height          =   4275
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   4635
      ExtentX         =   8176
      ExtentY         =   7541
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
      Location        =   "http:///"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnusettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnusettingsCapture 
         Caption         =   "&Capture clipboard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSettingsTranslateTo 
         Caption         =   "Translate &to"
         Begin VB.Menu mnuLanguage 
            Caption         =   "All translations"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintDestLanguage As Integer
Private mblnCaptureClipboard As Boolean

Private Sub cmdGo_Click()
    Dim strURL As String
    Dim strTextToTrans As String
    Dim intIndex As Integer
    
   
    strTextToTrans = Trim(txtWord)
    
    If mintDestLanguage = -1 Then
        strURL = "http://info.babylon.com/cgi-bin/info.cgi"

        strURL = strURL & "?word=" & strTextToTrans
        strURL = strURL & "&lang=" & ""
        strURL = strURL & "&ot=2&layout=combo2.html&n=10&keeplang=4"
    Else
        strURL = "http://info.babylon.com/cgi-bin/info.cgi"
        strURL = strURL & "?word=" & strTextToTrans
        strURL = strURL & "&lang=" & mintDestLanguage
        strURL = strURL & "&type=hp&layout=combo.html&n=10&list="
    End If
    webBabylon.Navigate strURL
                

End Sub

Private Sub Form_Load()
    Dim intIndex As Integer
    
    webBabylon.Navigate "about:blank"
    
    For intIndex = 1 To 18
        Load mnuLanguage(intIndex)
    Next intIndex
    
    mnuLanguage(0).Caption = "All translations"
    mnuLanguage(0).Tag = -1
    mnuLanguage(1).Caption = "Arabic"
    mnuLanguage(1).Tag = 15
    mnuLanguage(2).Caption = "Chinese"
    mnuLanguage(2).Tag = 9
    mnuLanguage(3).Caption = "Chinese (S)"
    mnuLanguage(3).Tag = 10
    mnuLanguage(4).Caption = "Dutch"
    mnuLanguage(4).Tag = 4
    mnuLanguage(5).Caption = "English"
    mnuLanguage(5).Tag = 0
    mnuLanguage(6).Caption = "Esperanto"
    mnuLanguage(6).Tag = 17
    mnuLanguage(7).Caption = "French"
    mnuLanguage(7).Tag = 1
    mnuLanguage(8).Caption = "German"
    mnuLanguage(8).Tag = 6
    mnuLanguage(9).Caption = "Greek"
    mnuLanguage(9).Tag = 11
    mnuLanguage(10).Caption = "Hebrew"
    mnuLanguage(10).Tag = 14
    mnuLanguage(11).Caption = "Italian"
    mnuLanguage(11).Tag = 2
    mnuLanguage(12).Caption = "Japanese"
    mnuLanguage(12).Tag = 8
    mnuLanguage(13).Caption = "Korean"
    mnuLanguage(13).Tag = 12
    mnuLanguage(14).Caption = "Portuguese"
    mnuLanguage(14).Tag = 5
    mnuLanguage(15).Caption = "Russian"
    mnuLanguage(15).Tag = 7
    mnuLanguage(16).Caption = "Spanish"
    mnuLanguage(16).Tag = 3
    mnuLanguage(17).Caption = "Thai"
    mnuLanguage(17).Tag = 16
    mnuLanguage(18).Caption = "Turkish"
    mnuLanguage(18).Tag = 13
    
    RestoreSettings
End Sub

Private Sub Form_Resize()
    
    If WindowState <> vbMinimized Then
        With webBabylon
            webBabylon.Move .Left, .Top, ScaleWidth - .Left * 2, ScaleHeight - .Top - .Left
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuLanguage_Click(Index As Integer)
    Dim intIndex As Integer
    
    If mnuLanguage(Index).Checked Then
        Exit Sub
    End If
    
    For intIndex = 0 To mnuLanguage.Count - 1
        mnuLanguage(intIndex).Checked = False
    Next intIndex
    
    mnuLanguage(Index).Checked = True
    mintDestLanguage = CInt(mnuLanguage(Index).Tag)
    
    
End Sub

Private Sub mnusettingsCapture_Click()
    mnusettingsCapture.Checked = Not mnusettingsCapture.Checked
    mblnCaptureClipboard = mnusettingsCapture.Checked
    tmrClipboard.Enabled = mblnCaptureClipboard
End Sub

Private Sub tmrClipboard_Timer()
    ReadClipBoard
End Sub

Private Sub webBabylon_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    txtWord.SetFocus
End Sub

Private Sub ReadClipBoard()
Static lastClip As String
Static ctm As Integer
Dim currentClip As String
'take only 1 out of 10 reads
ctm = ctm + 1: If ctm > 10 Then ctm = 0
If ctm > 0 Then Exit Sub
'
On Error GoTo noClipRead

currentClip = Clipboard.GetText
If currentClip <> lastClip Then
    lastClip = currentClip
    txtWord.Text = currentClip
    Call cmdGo_Click
    End If
noClipRead:
'
End Sub

Private Sub SaveSettings()
    SaveSetting App.ProductName, "Settings", "Language", CStr(mintDestLanguage)
    SaveSetting App.ProductName, "Settings", "Clipboard", CStr(mblnCaptureClipboard)
End Sub

Private Sub RestoreSettings()
    Dim intIndex As Integer
    
    mintDestLanguage = CInt(GetSetting(App.ProductName, "Settings", "Language", "-1"))
    mblnCaptureClipboard = CBool(GetSetting(App.ProductName, "Settings", "Clipboard", "False"))
    
    mnusettingsCapture.Checked = mblnCaptureClipboard
    tmrClipboard.Enabled = mblnCaptureClipboard
    
    For intIndex = 0 To mnuLanguage.Count - 1
        If CInt(mnuLanguage(intIndex).Tag) = mintDestLanguage Then
            mnuLanguage(intIndex).Checked = mintDestLanguage
        End If
    Next intIndex
End Sub
