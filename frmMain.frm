VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "PAK Control"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3960
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.pak"
      DialogTitle     =   "Open PAK"
      Filter          =   "PAK File (*.pak)|*.pak"
      Flags           =   38930
   End
   Begin MSComctlLib.ListView ListFiles 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageListGray 
      Left            =   2610
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":275E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":303A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   1058
      ButtonWidth     =   2170
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageListGray"
      DisabledImageList=   "ImageListGray"
      HotImageList    =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Open"
            Key             =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Add"
            Key             =   "Add"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Extract"
            Key             =   "Extract"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2070
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3916
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":476A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5046
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B76
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PAKFile As String
Dim FileListStart As Long
Dim Header As String
Dim FileList As String

Private Sub Form_Load()
ListFiles.ColumnHeaders.Add , , "Filename", Me.Width - 2000
ListFiles.ColumnHeaders.Add , , "Offset", 1000
ListFiles.ColumnHeaders.Add , , "Size", 1000
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    ListFiles.Width = Me.Width - 120
    ListFiles.Height = Me.Height - 1035
    ListFiles.ColumnHeaders(1).Width = Me.Width - 2150
    ListFiles.ColumnHeaders(2).Width = 1000
    ListFiles.ColumnHeaders(3).Width = 1000
End If
End Sub

Private Sub ListFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Erro
If PAKFile = "" Then
    MessageBox "You don't have opened any file, to you want create a new file ?", YesNo, Question
    If Result = 1 Then
        CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
        CommonDialog.DialogTitle = "Save PAK file"
        CommonDialog.Filter = "PAK File (*.pak)|*.pak"
        CommonDialog.ShowSave
        If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
        PAKCreate CommonDialog.FileName
    ElseIf Result = 2 Then
        Exit Sub
    End If
End If

frmBusy.Visible = True
Me.Enabled = False
Me.MousePointer = 11
For Files = 1 To Data.Files.Count
    PAKAdd PAKFile, Data.Files(Files), RemoveBackSlash(Data.Files(Files))
    Me.Refresh
Next Files
Unload frmBusy
Me.Enabled = True
Me.MousePointer = 0

Exit Sub
Erro:
If Err = 32755 Then
    PAKFile = ""
    CommonDialog.FileName = ""
    Toolbar.Buttons(4).Enabled = False
    ListFiles.ListItems.Clear
    Exit Sub
Else
    MessageBox "A unknown error occur!", OKOnly, Critical
    End
End If

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Erro

If Button.Key = "Open" Then
    CommonDialog.Flags = &H1000 + &H4 + &H8 + &H800
    CommonDialog.Filter = "PAK File (*.pak)|*.pak"
    CommonDialog.DialogTitle = "Open PAK file"
    CommonDialog.DefaultExt = ""
    CommonDialog.ShowOpen
    If FileExist(CommonDialog.FileName) = True Then
        PAKOpen CommonDialog.FileName
    End If

ElseIf Button.Key = "Add" Then
    If CommonDialog.FileName = "" Then
        MessageBox "You don't have opened any file, to you want create a new file ?", YesNo, Question
        If Result = 1 Then
            CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
            CommonDialog.DialogTitle = "Save PAK file"
            CommonDialog.Filter = "PAK File (*.pak)|*.pak"
            CommonDialog.DefaultExt = ".pak"
            CommonDialog.ShowSave
            If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
            If PAKCreate(CommonDialog.FileName) = False Then Err.Raise 1
        ElseIf Result = 2 Then
            Exit Sub
        End If
    End If
    CommonDialog.Flags = &H1000 + &H4 + &H8 + &H800
    CommonDialog.DialogTitle = "ADD files to PAK"
    CommonDialog.Filter = "All files (*.*)|*.*"
    CommonDialog.DefaultExt = ""
    CommonDialog.ShowOpen
    frmBusy.Show
    Me.Enabled = False
    Me.MousePointer = 11
    PAKAdd PAKFile, CommonDialog.FileName, CommonDialog.FileTitle
    Unload frmBusy
    Me.MousePointer = 0
    Me.Enabled = True
    
ElseIf Button.Key = "Extract" Then
    If ListFiles.SelectedItem = "" Then
        Exit Sub
    Else
        CommonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
        CommonDialog.DialogTitle = "Save file"
        CommonDialog.Filter = "All files (*.*)|*.*"
        CommonDialog.DefaultExt = ""
        CommonDialog.FileName = ListFiles.SelectedItem
        CommonDialog.ShowSave
        If FileExist(CommonDialog.FileName) = True Then Kill CommonDialog.FileName
        frmBusy.Show
        Me.Enabled = False
        Me.MousePointer = 11
        If PAKExtract(PAKFile, ListFiles.SelectedItem, CommonDialog.FileName) = False Then MessageBox "A error occur when try to extract the file!", OKOnly, Critical
        Me.Enabled = True
        Me.MousePointer = 0
        Unload frmBusy
    End If

ElseIf Button.Key = "About" Then
    frmAbout.Show 1, Me

End If

Exit Sub

Erro:
If Err = 32755 Then
    'PAKFile = ""
    CommonDialog.FileName = ""
    Toolbar.Buttons(4).Enabled = False
    ListFiles.ListItems.Clear
    Exit Sub
Else
    MessageBox "A unknown error occur!", OKOnly, Critical
    End
End If
End Sub

Function PAKOpen(FileName As String) As Boolean
Dim FileList As String
Dim OffSet As Long
Dim Size As Long
Dim Name As String
Dim LF As ListItem
Dim LFS As ListSubItem

On Error GoTo Erro

ListFiles.ListItems.Clear

'Check if is a valid PAK file
If PAKValid(FileName) = True Then
    PAKOpen = True
    FileNumber = FreeFile
    Open FileName For Binary As FileNumber
        'Is a valid PAK file
        'Get the FileList
        Get FileNumber, 7, FileListStart
        
        If FileListStart = 0 Then
            MessageBox "Empy file!", OKOnly, Information
            Close FileNumber
            Exit Function
        Else
            PAKFile = FileName
            Toolbar.Buttons(4).Enabled = True
            
            'Add the FileName, OffSet and Size in the ListView control
            ListFiles.ListItems.Clear
            
            Do
                Get FileNumber, FileListStart, OffSet
                FileListStart = FileListStart + 4
                
                Get FileNumber, FileListStart, Size
                FileListStart = FileListStart + 4
                
                Name = String$(255, Chr$(0))
                Get FileNumber, FileListStart, Name
                Name = Mid(Name, 1, InStr(1, Name, Chr$(0)) - 1)
                FileListStart = FileListStart + Len(Name) + 1
                
                If Name = "" Or OffSet = 0 Or Size = 0 Then
                    MessageBox "Empy file!", OKOnly, Information
                    Close FileNumber
                    Exit Function
                End If
            
                Set LF = ListFiles.ListItems.Add(, , Name)
                Set LFS = LF.ListSubItems.Add(, , OffSet)
                Set LFS = LF.ListSubItems.Add(, , Size)

                
            Loop Until FileListStart > LOF(FileNumber)
        End If
    Else
        'Is a invalid PAK file
        PAKOpen = False
        CommonDialog.FileName = ""
        PAKFile = ""
        Toolbar.Buttons(4).Enabled = False
        MessageBox "The specified filename is not a valid file!", OKOnly, Critical
        Close FileNumber
        Exit Function
    End If
Close FileNumber
Exit Function

Erro:
If Err = 5 Then
    PAKOpen = False
    PAKFile = ""
    CommonDialog.FileName = ""
    ListFiles.ListItems.Clear
    Toolbar.Buttons(4).Enabled = False
    MessageBox "A error occur when trying to read the file!", OKOnly, Critical
    Close FileNumber
    Exit Function
End If
End Function

Function PAKCreate(FileName As String) As Boolean
On Error GoTo Erro
Dim FileList As String

Header = "TPF2.0"
FileListStart = 0

If FileExist(FileName) = True Then
    PAKCreate = False
    Exit Function
Else
    FileNumber = FreeFile
    Open FileName For Binary As FileNumber
        Put FileNumber, 1, Header
        Put FileNumber, Len(Header) + 1, FileListStart
    Close FileNumber
End If
PAKFile = FileName
Toolbar.Buttons(4).Enabled = True
PAKCreate = True
Exit Function

Erro:
If Err <> 0 Then
    PAKCreate = False
    Exit Function
End If
End Function

Function PAKAdd(FilePAK As String, FileADD As String, NameADD As String) As Boolean
On Error GoTo Erro
Dim BytesADD As String
Dim OffSetADD As Long
Dim SizeADD As Long
Dim LF As ListItem
Dim LFS As ListSubItem

NameADD = NameADD & Chr$(0)

If FileExist(FilePAK) = False Or FileExist(FileADD) = False Then
    PAKAdd = False
    Exit Function
Else
    'Check if is a valid PAK file
    If PAKValid(FilePAK) = True Then
        'Is a valid PAK file
        
        FileNumberPAK = FreeFile
        Open FilePAK For Binary As FileNumberPAK

        
        'Get the FileList
        Get FileNumberPAK, 7, FileListStart

        'Get the FileList and put in the memory
        If FileListStart = 0 Then
            FileListStart = LOF(FileNumberPAK) + 1
            FileList = ""
        Else
            FileList = String(LOF(FileNumberPAK) - FileListStart + 1, Chr$(0))
            Get FileNumberPAK, FileListStart, FileList
        End If

        OffSetADD = FileListStart
        SizeADD = FileLen(FileADD)
            
        'Put the file inside of the PAK
        FileNumberADD = FreeFile
        frmBusy.lblFile = "Adding " & RemoveBackSlash(FileADD)
        frmBusy.Refresh
        Open FileADD For Binary As FileNumberADD
            If LOF(FileNumberADD) > 1000000 Then 'Divid the file in parts to use less memory and make less swap
                'BytesADD = String(LOF(FileNumberADD) / 100, Chr$(0))
                'For Position = 1 To LOF(FileNumberADD) Step Len(BytesADD)
                    'Get FileNumberADD, Position, BytesADD
                    'Put FileNumberPAK, FileListStart, BytesADD
                    'FileListStart = FileListStart + Len(BytesADD)
                'Next Position
                
                Position = -999999
                frmBusy.prgFile.Max = LOF(FileNumberADD)
                Do
                    Position = Position + 1000000
                    If Position + 999999 > LOF(FileNumberADD) Then
                        frmBusy.prgFile.Value = frmBusy.prgFile.Max
                        frmBusy.Refresh
                        BytesADD = String(LOF(FileNumberADD) - Position + 1, Chr$(0))
                    Else
                        frmBusy.prgFile.Value = Position
                        frmBusy.Refresh
                        BytesADD = String(1000000, Chr$(0))
                    End If
                    Get FileNumberADD, Position, BytesADD
                    Put FileNumberPAK, FileListStart, BytesADD
                    FileListStart = FileListStart + Len(BytesADD)
                Loop Until Position + 999999 > LOF(FileNumberADD)
                
            Else
                frmBusy.prgFile.Max = 1
                frmBusy.prgFile.Value = 0
                BytesADD = String(LOF(FileNumberADD), Chr$(0))
                Get FileNumberADD, 1, BytesADD
                Put FileNumberPAK, FileListStart, BytesADD
                FileListStart = FileListStart + Len(BytesADD)
                frmBusy.prgFile.Value = 1
            End If
        Close FileNumberADD
        
        'Add the new file in the FileList
        Put FileNumberPAK, 7, FileListStart
        Put FileNumberPAK, FileListStart, FileList
        Put FileNumberPAK, FileListStart + Len(FileList), OffSetADD
        Put FileNumberPAK, FileListStart + Len(FileList) + 4, SizeADD
        Put FileNumberPAK, FileListStart + Len(FileList) + 8, NameADD
        Close FileNumberPAK
    Else
        PAKAdd = False
        Exit Function
    End If
End If
PAKAdd = True
Set LF = ListFiles.ListItems.Add(, , Left(NameADD, Len(NameADD) - 1))
Set LFS = LF.ListSubItems.Add(, , OffSetADD)
Set LFS = LF.ListSubItems.Add(, , SizeADD)
Exit Function

Erro:
PAKAdd = False
Exit Function
End Function

Function PAKValid(PAKFileName As String) As Boolean
Dim Header As String
Header = String$(6, Chr$(0))

If FileExist(PAKFileName) = False Then
    PAKValid = False
    Exit Function
Else
    FileNumber = FreeFile
    Open PAKFileName For Binary As FileNumber
        Get FileNumber, 1, Header
        If Header = "TPF2.0" Then
            PAKValid = True
        Else
            PAKValid = False
        End If
    Close FileNumber
End If
End Function

Function PAKExtract(PAKFile As String, FileToExtract As String, DestinationFile As String) As Boolean
Dim BytesExtract As String
Dim OffSet As Long
Dim Size As Long
Dim Name As String

If FileExist(PAKFile) = False Or FileExist(DestinationFile) = True Then
    PAKExtract = False
    Exit Function
Else
    If PAKValid(PAKFile) = True Then
    
        FileNumber = FreeFile
        Open PAKFile For Binary As FileNumber
            'Get the FileList
            Get FileNumber, 7, FileListStart
        
            If FileListStart = 0 Then
                PAKExtract = False
                Close FileNumber
                Exit Function
            Else
                

                Do
                    Get FileNumber, FileListStart, OffSet
                    FileListStart = FileListStart + 4
                
                    Get FileNumber, FileListStart, Size
                    FileListStart = FileListStart + 4
                
                    Name = String$(255, Chr$(0))
                    Get FileNumber, FileListStart, Name
                    Name = Mid(Name, 1, InStr(1, Name, Chr$(0)) - 1)
                    FileListStart = FileListStart + Len(Name) + 1
                
                    If Name = "" Or OffSet = 0 Or Size = 0 Then
                        PAKExtract = False
                        Close FileNumber
                        Exit Function
                    ElseIf LCase(Name) = LCase(FileToExtract) Then
                        frmBusy.lblFile = "Extracting " & FileToExtract
                        DestinationNumber = FreeFile
                        Open DestinationFile For Binary As DestinationNumber
                            If Size > 100000 Then 'Divid the file in parts to use less memory and make less swap
                                'BytesExtract = String(Size / 100, Chr$(0))
                                'For Position = 1 To Size Step Len(BytesExtract)
                                    'Get FileNumber, Position + OffSet, BytesExtract
                                    'Put DestinationNumber, Position, BytesExtract
                                'Next Position
                                
                                Position = -1000000
                                frmBusy.prgFile.Max = Size
                                Do
                                    
                                    Position = Position + 1000000
                                    If Position + 999999 > Size Then
                                        BytesExtract = String(Size - Position, Chr$(0))
                                        frmBusy.prgFile.Value = frmBusy.prgFile.Max
                                        frmBusy.Refresh
                                    Else
                                        BytesExtract = String(1000000, Chr$(0))
                                        frmBusy.prgFile.Value = Position
                                        frmBusy.Refresh
                                    End If
                                    Get FileNumber, Position + OffSet, BytesExtract
                                    Put DestinationNumber, Position + 1, BytesExtract
                                Loop Until Position + 999999 >= Size
                            Else
                                BytesExtract = String(Size, Chr$(0))
                                Get FileNumber, OffSet, BytesExtract
                                Put DestinationNumber, 1, BytesExtract
                            End If
                        Close DestinationNumber
                        Close FileNumber
                        PAKExtract = True
                        Exit Function
                    End If
                Loop Until FileListStart > LOF(FileNumber)
            End If
        Close FileNumber
        PAKExtract = False
    Else
        PAKExtract = False
        Exit Function
    End If
End If

End Function

Function RemoveBackSlash(FileName As String) As String
Dim Temp As Integer

Do
    Temp = Slash
    Slash = InStr(Slash + 1, FileName, "\")
    If Slash = 0 Then
        RemoveBackSlash = Mid(FileName, Temp + 1)
        Exit Function
    End If
Loop
End Function
