VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmINI 
   Caption         =   "VisualINI"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView tvwSections 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   9340
      _Version        =   327682
      Indentation     =   443
      Style           =   5
      ImageList       =   "ImageList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.StatusBar sbBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   5880
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8361
            Text            =   "Untitled.ini"
            TextSave        =   "Untitled.ini"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "0 Section(s)"
            TextSave        =   "0 Section(s)"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "SCRL"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView lstTemp 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Key"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   4940
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialogs 
      Left            =   600
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.ListView lstKeys 
      Height          =   5325
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   9393
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Key"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   4940
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":624C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save As"
      End
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSection 
      Caption         =   "&Section"
      Begin VB.Menu mnuSectionNew 
         Caption         =   "&New Section"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSectionDelete 
         Caption         =   "&Delete Section"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSectionRename 
         Caption         =   "&Rename Section"
         Enabled         =   0   'False
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSectionCopyName 
         Caption         =   "&Copy Name"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSectionCopyPath 
         Caption         =   "Copy &Path"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuKey 
      Caption         =   "&Key"
      Begin VB.Menu mnuKeyNew 
         Caption         =   "&New Key"
      End
      Begin VB.Menu mnuKeyChangeData 
         Caption         =   "&Change Data..."
         Enabled         =   0   'False
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeyDelete 
         Caption         =   "&Delete Key"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuKeyRename 
         Caption         =   "&Rename Key"
         Enabled         =   0   'False
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeyCopyName 
         Caption         =   "C&opy Name"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuKeyCopyPath 
         Caption         =   "Copy &Path"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuKeyCopyValue 
         Caption         =   "Copy &Value"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "&Statusbar"
         Checked         =   -1  'True
         Shortcut        =   ^I
      End
      Begin VB.Menu sep45 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewProperties 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LoadedFile As String
Private Function StripPath(t As String) As String
Dim X As Integer
Dim ct As Integer
    StripPath = t
    X = InStr(t, "\")
    Do While X
        ct = X
        X = InStr(ct + 1, t, "\")
    Loop
    If ct > 0 Then StripPath = Mid(t, ct + 1)
End Function
Private Sub LoadData(File As String)
Dim strSection As String
Dim strTemp As String
Dim strKey As String
Dim strValue As String
Dim intKey As Integer
Dim i As Integer

tvwSections.Nodes.Clear
For i = 1 To lstKeys.Count - 1
    Unload lstKeys(i + 1)
Next

LoadedFile = File

If intKey = 0 Then intKey = 1
tvwSections.Nodes.Add , , "Root", StripPath(File), 3

Open File For Input As #1
    Do Until EOF(1)
        Input #1, strSection
        If Left$(strSection, 1) = "[" And Right$(strSection, 1) = "]" Then
            strTemp = Mid$(strSection, 2, Len(strSection) - 2)
            tvwSections.Nodes(1).Expanded = True
            tvwSections.Nodes.Add "Root", tvwChild, , strTemp, 1, 4
            intKey = intKey + 1
            Load lstKeys(intKey)
            lstKeys(intKey).Visible = True
        Else
            strTemp = strSection
            strKey = InStr(1, strTemp, "=")
            If strKey <> 0 Then
                strValue = Mid$(strTemp, strKey + 1)
                strKey = Left$(strTemp, strKey - 1)
                lstKeys(intKey).ListItems.Add 1, , strKey, , 2
                If strValue = "" Then strValue = "(Empty)"
                lstKeys(intKey).ListItems(1).SubItems(1) = strValue
                Set lstKeys(intKey).SelectedItem = lstKeys(intKey).ListItems(1)
            End If
        End If
    Loop
Close #1
sbBar.Panels(1).Text = File
sbBar.Panels(2).Text = tvwSections.Nodes.Count - 1 & " Section(s)"
tvwSections.Nodes(1).Selected = True
End Sub
Private Sub lstSections_ItemClick(ByVal Item As ComctlLib.ListItem)
   lstKeys(Item.Index + 1).ZOrder
End Sub


Private Sub Form_Load()
    tvwSections.Nodes.Add , , "Root", "Untitled.ini", 3
    tvwSections.Nodes(1).Expanded = True
    tvwSections.Nodes(1).Selected = True
    Dialogs.InitDir = App.Path
    LoadedFile = "Untitled"
    
End Sub


Private Sub Form_Resize()
On Error Resume Next
Dim i As Integer
If sbBar.Visible = True Then
    tvwSections.Height = Me.Height - 1075
    For i = 0 To lstKeys.Count
        lstKeys(i).Height = Me.Height - 1075
        lstKeys(i).Width = Me.Width - 3150
    Next
Else
    tvwSections.Height = Me.Height - 830
    For i = 0 To lstKeys.Count
        lstKeys(i).Height = Me.Height - 830
        lstKeys(i).Width = Me.Width - 3150
    Next
End If
End Sub

Private Sub lstKeys_DblClick(Index As Integer)
mnuKeyChangeData_Click
End Sub

Private Sub lstKeys_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If tvwSections.Nodes(1).Selected = False Then
        If lstKeys(Index).ListItems.Count <> 0 Then
            If lstKeys(Index).SelectedItem.Selected = False Then
                mnuKeyDelete.Enabled = False
                mnuKeyRename.Enabled = False
                mnuKeyCopyName.Enabled = False
                mnuKeyCopyPath.Enabled = False
                mnuKeyChangeData.Enabled = False
                mnuKeyCopyValue.Enabled = False
            Else
                mnuKeyDelete.Enabled = True
                mnuKeyRename.Enabled = True
                mnuKeyCopyName.Enabled = True
                mnuKeyCopyPath.Enabled = True
                mnuKeyChangeData.Enabled = True
                mnuKeyCopyValue.Enabled = True
            End If
        End If
        mnuKeyNew.Enabled = True
    Else
        mnuKeyNew.Enabled = False
        mnuKeyDelete.Enabled = False
        mnuKeyRename.Enabled = False
        mnuKeyCopyName.Enabled = False
        mnuKeyCopyPath.Enabled = False
        mnuKeyChangeData.Enabled = False
        mnuKeyCopyValue.Enabled = False
    End If
    PopupMenu mnuKey
End If
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileNew_Click()
Dim i As Integer
tvwSections.Nodes.Clear
For i = 2 To lstKeys.Count
    Unload lstKeys(i)
Next
tvwSections.Nodes.Add , , "Root", "Untitled.ini", 3
sbBar.Panels(1).Text = "Untitled.ini"
sbBar.Panels(2).Text = "0 Section(s)"
LoadedFile = "Untitled"
lstKeys(0).SetFocus
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo ErrHandler

Dialogs.DialogTitle = "Open INI file"
Dialogs.Flags = &H4
Dialogs.Filter = "Configration Settings (*.ini)|*.ini|All Files (*.*)|*.*"
Dialogs.ShowOpen
LoadData Dialogs.FileName
Dialogs.FileName = ""

ErrHandler:
    Exit Sub
End Sub
Private Sub SaveData(Where As String)
Dim i As Integer, j As Integer
Open Where For Output As #1
For i = 2 To tvwSections.Nodes.Count
        Print #1, "[" & tvwSections.Nodes(i) & "]"
        If lstKeys(i).ListItems.Count > 0 Then
            For j = 1 To lstKeys(i).ListItems.Count
                Print #1, lstKeys(i).ListItems(j).Text & "=" & lstKeys(i).ListItems(j).SubItems(1)
            Next
        End If
Next
Close #1
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo ErrHandler
If LoadedFile = "Untitled" Then
    Dialogs.DialogTitle = "Save INI file"
    Dialogs.Flags = &H4
    Dialogs.Filter = "Configration Settings (*.ini)|*.ini|All Files (*.*)|*.*"
    Dialogs.ShowSave
    SaveData Dialogs.FileName
    LoadData Dialogs.FileName
    Dialogs.FileName = ""
Else
    SaveData LoadedFile
End If
ErrHandler:
    Exit Sub
End Sub

Private Sub mnuFileSaveAs_Click()
    Dialogs.DialogTitle = "Save INI file"
    Dialogs.Flags = &H4
    Dialogs.Filter = "Configration Settings (*.ini)|*.ini|All Files (*.*)|*.*"
    Dialogs.ShowSave
    SaveData Dialogs.FileName
    LoadData Dialogs.FileName
    Dialogs.FileName = ""
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "VisualINI - Visual Editor for Configuration files" & vbCrLf & "Written by Agam Saran"
End Sub

Private Sub mnuKeyChangeData_Click()
AnalyseType lstKeys(tvwSections.SelectedItem.Index).SelectedItem.SubItems(1)
ValueName = lstKeys(tvwSections.SelectedItem.Index).SelectedItem.Text
frmChangeData.Show vbModal, Me
End Sub
Private Function AnalyseType(Text As String) As TextTypes
ValueData = Text
If IsNumeric(Text) = True Then
    AnalyseType = Number
End If
If Text = "True" Or Text = "False" Then
    AnalyseType = TrueFalse
End If
Select Case Left(Text, 3)
    Case "A:\", "B:\", "C:\", "D:\", "E:\", "F:\", "G:\", "H:\", "K:\", "L:\", "M:\", "N:\", "O:\", "P:\", "Q:\", "R:\", "S:\", "T:\", "U:\", "V:\", "W:\", "X:\", "Y:\", "Z:\"
        AnalyseType = Path
End Select
If AnalyseType = 0 Then
    AnalyseType = JustText
End If
ValueType = AnalyseType
End Function
Private Sub mnuKeyCopyName_Click()
Clipboard.Clear
Clipboard.SetText lstKeys(tvwSections.SelectedItem.Index).SelectedItem.Text
End Sub

Private Sub mnuKeyCopyPath_Click()
Clipboard.Clear
Clipboard.SetText tvwSections.SelectedItem.FullPath & "\" & lstKeys(tvwSections.SelectedItem.Index).SelectedItem.Text
End Sub

Private Sub mnuKeyCopyValue_Click()
Clipboard.Clear
Clipboard.SetText lstKeys(tvwSections.SelectedItem.Index).SelectedItem.SubItems(1)
End Sub

Private Sub mnuKeyDelete_Click()
Dim CurSec As Integer
Dim MsgAns As Integer
MsgAns = MsgBox("Are you sure that you want to delete this key?", vbYesNo)
If MsgAns = vbYes Then
    CurSec = tvwSections.SelectedItem.Index
    lstKeys(CurSec).ListItems.Remove lstKeys(CurSec).SelectedItem.Index
End If
End Sub

Private Sub mnuKeyNew_Click()
Dim CurSec As Integer
CurSec = tvwSections.SelectedItem.Index
lstKeys(CurSec).Sorted = False
lstKeys(CurSec).ListItems.Add lstKeys(CurSec).ListItems.Count + 1, , "New Key", , 2
lstKeys(CurSec).ListItems(lstKeys(CurSec).ListItems.Count).Selected = True
lstKeys(CurSec).ZOrder
lstKeys(CurSec).SetFocus
For CurSec = 0 To 1
    lstKeys(tvwSections.SelectedItem.Index).StartLabelEdit
Next
End Sub

Private Sub mnuKeyRename_Click()
lstKeys(tvwSections.SelectedItem.Index).StartLabelEdit
End Sub

Private Sub mnuSectionCopyName_Click()
Clipboard.Clear
Clipboard.SetText tvwSections.SelectedItem.Text
End Sub

Private Sub mnuSectionCopyPath_Click()
Clipboard.Clear
Clipboard.SetText tvwSections.SelectedItem.FullPath
End Sub

Private Sub mnuSectionDelete_Click()
Dim bolAns As VbMsgBoxResult
Dim i As Integer, j As Integer
bolAns = MsgBox("Are you sure that you want to delete the section """ & tvwSections.SelectedItem.Text & """", vbYesNo + vbQuestion, "Delete Section")
If bolAns = vbYes Then
    If tvwSections.SelectedItem.Index = tvwSections.Nodes.Count Then
        Unload lstKeys(tvwSections.Nodes.Count)
        tvwSections.Nodes.Remove tvwSections.SelectedItem.Index
        tvwSections.Nodes(1).Selected = True
        lstKeys(0).ZOrder
        Exit Sub
    End If
    For j = tvwSections.SelectedItem.Index To lstKeys.Count - 1
        For i = 1 To lstKeys(j + 1).ListItems.Count
            lstTemp.ListItems.Add , , lstKeys(j + 1).ListItems(i).Text, , 2
            lstTemp.ListItems(i).SubItems(1) = lstKeys(j + 1).ListItems(i).SubItems(1)
        Next
        Unload lstKeys(j)
        Load lstKeys(j)
        lstKeys(j).Visible = True
        For i = 1 To lstTemp.ListItems.Count
            lstKeys(j).ListItems.Add , , lstTemp.ListItems(i).Text, , 2
            lstKeys(j).ListItems(i).SubItems(1) = lstTemp.ListItems(i).SubItems(1)
        Next
        lstKeys(tvwSections.SelectedItem.Index).ZOrder
        lstTemp.ListItems.Clear
    Next
    Unload lstKeys(tvwSections.Nodes.Count)
    tvwSections.Nodes.Remove tvwSections.SelectedItem.Index
    sbBar.Panels(2).Text = tvwSections.Nodes.Count - 1 & " Section(s)"
End If
End Sub

Private Sub mnuSectionNew_Click()
    tvwSections.Nodes.Add "Root", tvwChild, , "New Section", 1, 4
    Load lstKeys(lstKeys.Count + 1)
    lstKeys(lstKeys.Count).Visible = True
    lstKeys(lstKeys.Count).ZOrder
    Set tvwSections.SelectedItem = tvwSections.Nodes(tvwSections.Nodes.Count)
    tvwSections.SetFocus
    tvwSections.StartLabelEdit
    lstKeys(lstKeys.Count).ListItems.Clear
    sbBar.Panels(2).Text = tvwSections.Nodes.Count - 1 & " Section(s)"

End Sub



Private Sub mnuSectionRename_Click()
tvwSections.StartLabelEdit
End Sub

Private Sub mnuViewProperties_Click()
Dim i As Integer, KeyCount As Integer
With frmProperties
    If LoadedFile = "Untitled" Then
        .lblApp.Caption = "Untitled.ini"
        .lblShadow.Caption = "Untitled.ini"
        .lblPath = "Not Saved"
        .lblStatus = "Waiting to Save the Configuration File"
    Else
        .lblApp.Caption = StripPath(LoadedFile)
        .lblShadow.Caption = StripPath(LoadedFile)
        .lblPath = LoadedFile
        .lblStatus = "File saved to Hard Disk"
    End If
    KeyCount = 0
    For i = 1 To tvwSections.Nodes.Count - 1
        KeyCount = KeyCount + lstKeys(i + 1).ListItems.Count
    Next
    .lblKeys.Caption = KeyCount
    .lblSections.Caption = tvwSections.Nodes.Count - 1
    .Show vbModal, Me
End With
End Sub

Private Sub mnuViewRefresh_Click()
If LoadedFile <> "Untitled" Then
    LoadData LoadedFile
End If
End Sub

Private Sub mnuViewStatusbar_Click()
mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
sbBar.Visible = mnuViewStatusbar.Checked
Form_Resize
End Sub

Private Sub tvwSections_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If tvwSections.Nodes(1).Selected = True Then
        mnuSectionDelete.Enabled = False
        mnuSectionRename.Enabled = False
        mnuSectionCopyName.Enabled = False
        mnuSectionCopyPath.Enabled = False
    Else
        mnuSectionDelete.Enabled = True
        mnuSectionRename.Enabled = True
        mnuSectionCopyName.Enabled = True
        mnuSectionCopyPath.Enabled = True
    End If
    PopupMenu mnuSection
End If
End Sub

Private Sub tvwSections_NodeClick(ByVal Node As ComctlLib.Node)
If Node.Index = 1 Then
    lstKeys.Item(0).ZOrder
Else
    If lstKeys(Node.Index).Sorted = False Then
        lstKeys(Node.Index).SortOrder = lvwAscending
        lstKeys(Node.Index).Sorted = True
    End If
    lstKeys(Node.Index).ZOrder
End If
End Sub
