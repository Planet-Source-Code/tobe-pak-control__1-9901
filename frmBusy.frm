VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusy 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1980
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgFile 
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmBusy.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      Caption         =   "Adding ABC.DEF"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1530
      Width           =   3345
   End
   Begin VB.Label lblBusy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label lblBusy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   180
      Width           =   1290
   End
End
Attribute VB_Name = "frmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
frmMain.ZOrder
End Sub
