VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Waiting to save the configuration file"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   2580
   End
   Begin VB.Label lblKeys 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label lblSections 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Not Saved"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   1440
      Width           =   765
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   6840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label4 
      Caption         =   "Status :"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   "Location :"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label Label2 
      Caption         =   "Total Keys :"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Total Sections :"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   1110
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   6840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   240
      Picture         =   "frmProperties.frx":000C
      Top             =   240
      Width           =   630
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Untitled.ini"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   510
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2265
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Untitled.ini"
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
      Left            =   1110
      TabIndex        =   1
      Top             =   390
      Width           =   2265
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload Me
End Sub
