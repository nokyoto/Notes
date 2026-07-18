VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNotes 
   Caption         =   "Note"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "notes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "notes.frx":0442
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuViewRTL 
         Caption         =   "Right-to-Left"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActiveFile As String
Dim ActiveFilePath As String
Private Sub Form_Load()
    ActiveFilePath = "Untitled"
    Text1.Text = ""
    SetWinTitle
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CompareTexts
End Sub
Private Sub Form_Resize()
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
Private Sub SaveActiveFile()
    Dim f As Integer
    f = FreeFile
    Open ActiveFilePath For Output As #f
    Print #f, Text1.Text
    Close #f
End Sub
Private Sub SetWinTitle()
    Me.Caption = ActiveFilePath & " in Notes"
End Sub
Private Sub SaveAsActiveFile()
    CommonDialog1.CancelError = True
    On Error GoTo cancelled
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    ActiveFilePath = CommonDialog1.FileName
    GoTo saveNow

cancelled:
    Exit Sub
saveNow:
    SaveActiveFile
    ActiveFile = Text1.Text
End Sub
Private Sub CompareTexts()
    If Text1.Text <> ActiveFile Then
        Dim SaveModified As VbMsgBoxResult
        SaveModified = MsgBox("File is modified. Save it?", vbYesNo, "Save?")
        If SaveModified = vbYes Then
            If ActiveFilePath = "Untitled" Then
                SaveAsActiveFile
            Else
                SaveActiveFile
            End If
        End If
    End If
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileSaveAs_Click()
    SaveAsActiveFile
    SetWinTitle
End Sub
Private Sub mnuFileSave_Click()
    If ActiveFilePath = "Untitled" Then
        SaveAsActiveFile
        SetWinTitle
    Else
        SaveActiveFile
    End If
End Sub
Private Sub mnuFileNew_Click()
    CompareTexts
    Text1.Text = ""
End Sub
Private Sub mnuFileOpen_Click()
    CompareTexts
    CommonDialog1.CancelError = True
    On Error GoTo cancelled
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    ActiveFilePath = CommonDialog1.FileName
    GoTo loadFile
cancelled:
    Exit Sub
loadFile:
    Dim f As Integer
    f = FreeFile
    Open ActiveFilePath For Input As #f
    ActiveFile = Input$(LOF(f), f)
    Close #f
    Text1.Text = ActiveFile
    SetWinTitle
End Sub
Private Sub mnuViewFont_Click()
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlCFScreenFonts
    On Error GoTo cancelled
    CommonDialog1.ShowFont
    GoTo setFont
cancelled:
    Exit Sub
setFont:
    Text1.FontName = CommonDialog1.FontName
    Text1.FontSize = CommonDialog1.FontSize
    Text1.FontBold = CommonDialog1.FontBold
    Text1.FontItalic = CommonDialog1.FontItalic
End Sub
Private Sub mnuViewRTL_Click()
    If mnuViewRTL.Checked = False Then
    mnuViewRTL.Checked = True
    Text1.Alignment = 1
    Else
    mnuViewRTL.Checked = False
    Text1.Alignment = 0
    End If
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub
