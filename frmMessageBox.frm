VERSION 5.00
Begin VB.Form frmMessageBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "frmMessageBox.frx":0000
   ScaleHeight     =   1125
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctBtnNormal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   30
      Picture         =   "frmMessageBox.frx":632E
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   8
      Top             =   690
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pctBtnDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   30
      Picture         =   "frmMessageBox.frx":6D8C
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pctButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   1
      Left            =   2655
      Picture         =   "frmMessageBox.frx":77EA
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   5
      Top             =   675
      Visible         =   0   'False
      Width           =   1020
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   70
         Width           =   915
      End
   End
   Begin VB.PictureBox pctButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   3690
      Picture         =   "frmMessageBox.frx":8248
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   3
      Top             =   675
      Visible         =   0   'False
      Width           =   1020
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   70
         Width           =   915
      End
   End
   Begin VB.PictureBox pctButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   2
      Left            =   1080
      Picture         =   "frmMessageBox.frx":8CA6
      ScaleHeight     =   345
      ScaleWidth      =   1020
      TabIndex        =   0
      Top             =   675
      Visible         =   0   'False
      Width           =   1020
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   2
         Top             =   70
         Width           =   915
      End
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   810
      TabIndex        =   1
      Top             =   90
      Width           =   3885
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   360
      Picture         =   "frmMessageBox.frx":9704
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "frmMessageBox.frx":9FCE
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "frmMessageBox.frx":A898
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   360
      Picture         =   "frmMessageBox.frx":B162
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
    'Release capture
    Call ReleaseCapture
    'Send a 'left mouse button down on caption'-message to our form
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
    'Release capture
    Call ReleaseCapture
    'Send a 'left mouse button down on caption'-message to our form
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub lblCaption_Click(Index As Integer)
pctButton_Click (Index)
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pctButton(Index).Picture = pctBtnDown.Picture
lblCaption(Index).ForeColor = RGB(255, 255, 255)
End Sub

Private Sub lblCaption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pctButton(Index).Picture = pctBtnNormal.Picture
lblCaption(Index).ForeColor = RGB(0, 0, 0)
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
    'Release capture
    Call ReleaseCapture
    'Send a 'left mouse button down on caption'-message to our form
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub pctButton_Click(Index As Integer)
If lblCaption(Index).Caption = "OK" Then
    Result = 0
ElseIf lblCaption(Index).Caption = "Yes" Then
    Result = 1
ElseIf lblCaption(Index).Caption = "No" Then
    Result = 2
ElseIf lblCaption(Index).Caption = "Cancel" Then
    Result = 3
ElseIf lblCaption(Index).Caption = "Retry" Then
    Result = 4
ElseIf lblCaption(Index).Caption = "Ignore" Then
    Result = 5
ElseIf lblCaption(Index).Caption = "Abort" Then
    Result = 6
Else
    Result = 7
End If

Unload Me
End Sub

Private Sub pctButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pctButton(Index).Picture = pctBtnDown.Picture
lblCaption(Index).ForeColor = RGB(255, 255, 255)
End Sub

Private Sub pctButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
pctButton(Index).Picture = pctBtnNormal.Picture
lblCaption(Index).ForeColor = RGB(0, 0, 0)
End Sub
