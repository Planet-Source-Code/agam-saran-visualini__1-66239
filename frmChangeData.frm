VERSION 5.00
Begin VB.Form frmChangeData 
   BackColor       =   &H00B24801&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Data"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmChangeData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2180
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   5055
   End
   Begin VB.ComboBox cmbBoolean 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   2670
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Value :"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name :"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD6234&
      Height          =   510
      Left            =   390
      TabIndex        =   9
      Top             =   270
      Width           =   2670
   End
   Begin VB.Shape shpName 
      BackColor       =   &H00CC8859&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00CC8859&
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00CC8859&
      Height          =   1815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "frmChangeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
On Error GoTo ErrHandler
With frmINI.Dialogs
    .DialogTitle = "Browse..."
    .Flags = &H4
    .Filter = "All Files (*.*)|*.*"
    .ShowOpen
    txtData.Text = .FileName
    .FileName = ""
End With
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If ValueType = TrueFalse Then
    txtData.Text = cmbBoolean.List(cmbBoolean.ListIndex)
End If
With frmINI
    .lstKeys(.tvwSections.SelectedItem.Index).SelectedItem.Text = txtName.Text
    .lstKeys(.tvwSections.SelectedItem.Index).SelectedItem.SubItems(1) = txtData.Text
End With
Unload Me
End Sub

Private Sub Form_Load()
FlatBorder txtName.hwnd
FlatBorder txtData.hwnd
FlatBorder cmdBrowse.hwnd
cmbBoolean.AddItem "True"
cmbBoolean.AddItem "False"
txtName.Text = ValueName
txtData.Text = ValueData
Select Case ValueType
    Case JustText, Number
        txtData.Visible = True
        txtData.Width = 5055
        cmbBoolean.Visible = False
        cmdBrowse.Visible = False
    Case Path
        txtData.Visible = True
        txtData.Width = 4575
        cmbBoolean.Visible = False
        cmdBrowse.Visible = True
    Case TrueFalse
        txtData.Visible = False
        cmbBoolean.Visible = True
        cmdBrowse.Visible = False
        If ValueData = "True" Then
            cmbBoolean.ListIndex = 0
        Else
            cmbBoolean.ListIndex = 1
        End If
End Select
txtData.SelLength = Len(txtData.Text)
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
If ValueType = Number Then
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End If
End Sub
