VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frexplorer"
   ClientHeight    =   5355
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgGrosse 
      Left            =   8190
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDir.frx":0000
            Key             =   "filelarge"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   7515
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDir.frx":031A
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDir.frx":0474
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDir.frx":0A0E
            Key             =   "selfolder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDir.frx":0FA8
            Key             =   "filesmall"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbView 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   635
      ButtonWidth     =   2064
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Big icons"
            Key             =   "icons"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Small icons"
            Key             =   "sicons"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "list"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Details"
            Key             =   "details"
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lsvFiles 
      Height          =   4515
      Left            =   3150
      TabIndex        =   2
      Top             =   765
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   7964
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgGrosse"
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "nom"
         Text            =   "Name"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "taille"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView trvDir 
      Height          =   4515
      Left            =   90
      TabIndex        =   1
      Top             =   765
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   7964
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgIcons"
      Appearance      =   1
   End
   Begin VB.DriveListBox drvHDs 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   2985
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnuFiles 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Open with notepad"
         Index           =   1
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Cut"
         Index           =   3
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Copy"
         Index           =   4
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Delete"
         Index           =   6
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Properties"
         Index           =   8
      End
   End
   Begin VB.Menu mnuDir1 
      Caption         =   "RightDir"
      Visible         =   0   'False
      Begin VB.Menu mnuRightDir 
         Caption         =   "View"
         Index           =   0
         Begin VB.Menu mnuView 
            Caption         =   "Big icons"
            Index           =   0
         End
         Begin VB.Menu mnuView 
            Caption         =   "Small icons"
            Index           =   1
         End
         Begin VB.Menu mnuView 
            Caption         =   "List"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu mnuView 
            Caption         =   "Details"
            Index           =   3
         End
      End
      Begin VB.Menu mnuRightDir 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuRightDir 
         Caption         =   "Refresh"
         Index           =   2
      End
      Begin VB.Menu mnuRightDir 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuRightDir 
         Caption         =   "Paste"
         Index           =   4
      End
      Begin VB.Menu mnuRightDir 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuRightDir 
         Caption         =   "Properties"
         Index           =   6
      End
   End
   Begin VB.Menu mnuDir2 
      Caption         =   "LeftDir"
      Visible         =   0   'False
      Begin VB.Menu mnuLeftDir 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "Cut"
         Index           =   2
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "Copy"
         Index           =   3
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "Paste"
         Index           =   4
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "Delete"
         Index           =   6
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuLeftDir 
         Caption         =   "Properties"
         Index           =   8
      End
   End
   Begin VB.Menu mnuDir3 
      Caption         =   "DirRecycle"
      Visible         =   0   'False
      Begin VB.Menu mnuDirRecycle 
         Caption         =   "Empty RecycleBin"
      End
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 261
    cAlternate As String * 14
End Type

Private Const SEE_MASK_INVOKEIDLIST As Long = &HC
Private Const SEE_MASK_FLAG_NO_UI As Long = &H400
Private Const FO_DELETE As Long = &H3
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const SW_SHOWNORMAL As Long = 1
Private Const VK_SHIFT As Long = &H10
Private Const FO_COPY As Long = &H2
Private Const FO_MOVE As Long = &H1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private mstrPath As String
Private mblnRefresh As Boolean
Private mstrOldName As String
Private mblnOnAFile As Boolean

Private Sub drvHDs_Change()

    lsvFiles.ListItems.Clear
    RefreshTreeView

End Sub

Private Sub EraseFile(ByVal FileName As String)

Dim intResult As Integer
Dim filop As SHFILEOPSTRUCT

    'simulate behavior of windows explorer
    'if shift is pressed, don't send to recycle bin
    If GetKeyState(VK_SHIFT) < 0 Then
        With filop
            .wFunc = FO_DELETE
            .pFrom = FileName
        End With
        SHFileOperation filop
    Else
        With filop
            .fFlags = FOF_ALLOWUNDO  'send to recycle bin
            .wFunc = FO_DELETE
            .pFrom = FileName
        End With
        SHFileOperation filop
    End If

End Sub

Private Sub Form_Load()

    tbView.Buttons("list").Value = tbrPressed
    mblnRefresh = False
    drvHDs.Drive = "c:"
    'on winXP, the above command doesn't trigger the change event of drivelistbox
    If trvDir.Nodes.Count = 0 Then
        RefreshTreeView
    End If

End Sub

Private Sub GetFiles(ByVal nodSelected As Node)

Dim objlstFile As ListItem
Dim hFind As Long
Dim w32FindData As WIN32_FIND_DATA
Dim blnFileExists As Boolean
Dim strFileName As String

    lsvFiles.ListItems.Clear
    frmDir.MousePointer = vbHourglass
    'disable refresh of the control
    LockWindowUpdate lsvFiles.hwnd
    hFind = FindFirstFile(nodSelected.Key & "\*.*", w32FindData)
    If hFind = INVALID_HANDLE_VALUE Then
        MsgBox "Invalid directory, refreshing...", vbCritical, frmDir.Caption
        LockWindowUpdate 0
        mstrPath = mstrOldName
        RefreshTreeView
        SearchDir
        frmDir.MousePointer = vbDefault
        Exit Sub
    End If
    On Error GoTo errPermission
    blnFileExists = True
    While blnFileExists
        strFileName = Left$(w32FindData.cFileName, InStr(1, w32FindData.cFileName, vbNullChar) - 1)
        If LenB(strFileName) Then
            If AscW(strFileName) <> 46 Then
                If (w32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                Else
                    Set objlstFile = lsvFiles.ListItems.Add(, strFileName, strFileName, "filelarge", "filesmall")
                    If w32FindData.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN Then
                        objlstFile.Ghosted = True
                    End If
                    Select Case w32FindData.nFileSizeLow
                    Case 0
                        objlstFile.ListSubItems.Add 1, "size", "0 Kb"
                    Case 1 To 1024
                        objlstFile.ListSubItems.Add 1, "size", "1 Kb"
                    Case Else
                        objlstFile.ListSubItems.Add 1, "size", Round(w32FindData.nFileSizeLow / 1024, 0) & " Kb"
                    End Select
                End If
            End If
        End If
        blnFileExists = FindNextFile(hFind, w32FindData)
    Wend
    FindClose hFind
    'restore refresh of the control
    LockWindowUpdate 0
    frmDir.MousePointer = vbDefault
    Exit Sub
errPermission:
    If Err.Number = 70 Then
        MsgBox "Access denied", vbInformation, frmDir.Caption
    End If
    LockWindowUpdate 0

End Sub

Private Sub GetSubDirs(ByVal nodParent As Node)

Dim nodTemp As Node
Dim hFind As Long
Dim w32FindData As WIN32_FIND_DATA
Dim blnFileExists As Boolean
Dim strFileName As String

    On Error GoTo errSubDirs
    hFind = FindFirstFile(nodParent.Key & "\*.*", w32FindData)
    blnFileExists = True
    While blnFileExists
        strFileName = Left$(w32FindData.cFileName, InStr(1, w32FindData.cFileName, vbNullChar) - 1)
        If LenB(strFileName) Then
            If AscW(strFileName) <> 46 Then
                If (w32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    Set nodTemp = trvDir.Nodes.Add(nodParent.Key, tvwChild, UCase$(nodParent.Key & "\" & strFileName), strFileName, "folder", "selfolder")
                    nodTemp.Sorted = True
                End If
            End If
        End If
        blnFileExists = FindNextFile(hFind, w32FindData)
    Wend
    FindClose hFind
errSubDirs:

End Sub

Private Sub lsvFiles_AfterLabelEdit(Cancel As Integer, NewString As String)

    On Error GoTo errName
    Name trvDir.SelectedItem.Key & "\" & mstrOldName As trvDir.SelectedItem.Key & "\" & NewString
    Call GetFiles(trvDir.SelectedItem)

Exit Sub

errName:
    MsgBox Err.Description, vbCritical, "Error #" & CStr(Err.Number)
    Cancel = 1

End Sub

Private Sub lsvFiles_BeforeLabelEdit(Cancel As Integer)

    mstrOldName = lsvFiles.SelectedItem.Text

End Sub

Private Sub lsvFiles_DblClick()

    If mblnOnAFile Then
        Call mnuFiles_Click(0)
    End If

End Sub

Private Sub lsvFiles_KeyDown(KeyCode As Integer, Shift As Integer)

    'refresh files
    If KeyCode = vbKeyF5 Then
        Call GetFiles(trvDir.SelectedItem)
    ElseIf KeyCode = vbKeyF2 Then 'edit name of file
        lsvFiles.StartLabelEdit
    End If

End Sub

Private Sub lsvFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        'if mouse is not on a file, open a directory menu
        If TypeName(lsvFiles.HitTest(x, y)) = "Nothing" Then
            PopupMenu mnuDir1
        Else
            PopupMenu mnuFile
        End If
    Else
        If TypeName(lsvFiles.HitTest(x, y)) = "Nothing" Then
            If TypeName(lsvFiles.SelectedItem) <> "Nothing" Then
                lsvFiles.SelectedItem.Selected = False 'remove the highlight on the file
                Set lsvFiles.SelectedItem = Nothing 'remove the selection "rectangle" around the file
            End If
            mblnOnAFile = False
        Else
            mblnOnAFile = True
        End If
    End If

End Sub

Private Sub mnuDirRecycle_Click()

    SHEmptyRecycleBin Me.hwnd, vbNullString, 0

End Sub

Private Sub mnuFiles_Click(Index As Integer)

    Select Case Index
    Case 0
        Call ShellExecute(Me.hwnd, "open", trvDir.SelectedItem.Key & "\" & lsvFiles.SelectedItem.Text, "", trvDir.SelectedItem.Key, SW_SHOWNORMAL)
    Case 1
        Call ShellExecute(Me.hwnd, "open", "notepad.exe", trvDir.SelectedItem.Key & "\" & lsvFiles.SelectedItem.Key, trvDir.SelectedItem.Key, SW_SHOWNORMAL)
    Case 3
        Clipboard.SetText "frexplorer;" & trvDir.SelectedItem.Key & ";" & lsvFiles.SelectedItem.Text & ";cut"
    Case 4
        Clipboard.SetText "frexplorer;" & trvDir.SelectedItem.Key & ";" & lsvFiles.SelectedItem.Text & ";copy"
    Case 6
        EraseFile trvDir.SelectedItem.Key & "\" & lsvFiles.SelectedItem.Text
        Call GetFiles(trvDir.SelectedItem)
    Case 8
        ShowProps trvDir.SelectedItem.Key & "\" & lsvFiles.SelectedItem.Text
    End Select

End Sub

Private Sub mnuLeftDir_Click(Index As Integer)

    Select Case Index
    Case 0
        trvDir_KeyDown vbKeyF5, 0
    Case 2
        Clipboard.SetText "frexplorer;" & trvDir.SelectedItem.Key & ";;cut"
    Case 3
        Clipboard.SetText "frexplorer;" & trvDir.SelectedItem.Key & ";;copy"
    Case 4
        PasteFile
    Case 6
        If Len(trvDir.SelectedItem.Key) = 3 Then
            MsgBox "You can't delete root", vbCritical, "Delete"
            Exit Sub
        End If
        EraseFile trvDir.SelectedItem.Key
        trvDir.SelectedItem.Parent.Selected = True
        trvDir_KeyDown vbKeyF5, 0
    Case 8
        ShowProps trvDir.SelectedItem.Key
    End Select

End Sub

Private Sub mnuRightDir_Click(Index As Integer)

    Select Case Index
    Case 2
        Call lsvFiles_KeyDown(vbKeyF5, 0)
    Case 4
        PasteFile
    Case 6
        ShowProps trvDir.SelectedItem.Key
    End Select

End Sub

Private Sub mnuView_Click(Index As Integer)

    Select Case Index
    Case 0
        Call tbView_ButtonClick(tbView.Buttons("icons"))
        tbView.Buttons("icons").Value = tbrPressed
    Case 1
        Call tbView_ButtonClick(tbView.Buttons("sicons"))
        tbView.Buttons("sicons").Value = tbrPressed
    Case 2
        Call tbView_ButtonClick(tbView.Buttons("list"))
        tbView.Buttons("list").Value = tbrPressed
    Case 3
        Call tbView_ButtonClick(tbView.Buttons("details"))
        tbView.Buttons("details").Value = tbrPressed
    End Select

End Sub

Private Sub PasteFile()

Dim strData As String
Dim strDir As String
Dim strFile As String
Dim intPos As Integer
Dim intPos2 As Integer
Dim filop As SHFILEOPSTRUCT

    If Clipboard.GetFormat(vbCFText) Then
        strData = Clipboard.GetText
        'verify if we copied from this program
        If Left$(strData, 10) = "frexplorer" Then
            intPos = InStr(12, strData, ";", vbTextCompare)
            intPos2 = InStr(intPos + 1, strData, ";", vbTextCompare)
            strDir = Mid$(strData, 12, intPos - 12)
            strFile = Mid$(strData, intPos + 1, intPos2 - intPos - 1)
            With filop
                'if there's no filename, it's a complete dir
                If intPos = intPos2 - 1 Then
                    .pFrom = strDir
                Else
                    .pFrom = strDir & "\" & strFile
                End If
                .pTo = trvDir.SelectedItem.Key
                If Mid$(strData, intPos2 + 1) = "copy" Then
                    .wFunc = FO_COPY
                Else
                    .wFunc = FO_MOVE
                End If
            End With
            SHFileOperation filop
            'if not canceled
            If filop.fAnyOperationsAborted = False Then
                'if it was a file, refresh listview, else treeview
                If strFile <> "" Then
                    lsvFiles_KeyDown vbKeyF5, 0
                Else
                    Clipboard.SetText ""
                    trvDir_KeyDown vbKeyF5, 0
                End If
            End If
        End If
    End If

End Sub

Private Sub RefreshTreeView()

Dim nodTemp As Node
Dim hFind As Long
Dim w32FindData As WIN32_FIND_DATA
Dim blnFileExists As Boolean
Dim strFileName As String

    trvDir.Nodes.Clear
    frmDir.MousePointer = vbHourglass
    hFind = FindFirstFile(Left$(drvHDs.Drive, 2) & "\*.*", w32FindData)
    If hFind = INVALID_HANDLE_VALUE Then
        drvHDs.Drive = "C:"
        hFind = FindFirstFile(Left$(drvHDs.Drive, 2) & "\*.*", w32FindData)
    End If
    Set nodTemp = trvDir.Nodes.Add(, , UCase$(Left$(drvHDs.Drive, 2) & "\"), UCase$(Left$(drvHDs.Drive, 2) & "\"), "drive", "drive")
    nodTemp.Sorted = True
    nodTemp.Tag = 1
    blnFileExists = True
    While blnFileExists
        strFileName = Left$(w32FindData.cFileName, InStr(1, w32FindData.cFileName, vbNullChar) - 1)
        If LenB(strFileName) Then
            If AscW(strFileName) <> 46 Then
                If (w32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    Set nodTemp = trvDir.Nodes.Add(UCase$(Left$(drvHDs.Drive, 2) & "\"), tvwChild, UCase$(Left$(drvHDs.Drive, 2) & "\" & strFileName), strFileName, "folder", "selfolder")
                    nodTemp.Sorted = True
                    GetSubDirs nodTemp
                End If
            End If
        End If
        blnFileExists = FindNextFile(hFind, w32FindData)
    Wend
    FindClose hFind
    trvDir.Nodes(UCase$(Left$(drvHDs.Drive, 2) & "\")).Expanded = True
    trvDir.SelectedItem = trvDir.Nodes(UCase$(Left$(drvHDs.Drive, 2) & "\"))
    If Not mblnRefresh Then
        Call GetFiles(trvDir.Nodes(UCase$(Left$(drvHDs.Drive, 2) & "\")))
    End If
    frmDir.MousePointer = vbDefault

End Sub

Private Sub SearchDir()

Dim intPos As Integer, intPos2 As Integer
Dim strDir As String, strPath As String

    If Dir(mstrPath, vbDirectory) = "" Then Exit Sub
    intPos = InStr(1, mstrPath, "\", vbTextCompare)
    intPos2 = 1
    strPath = trvDir.SelectedItem.Key
    While intPos2 <> 0
        intPos2 = InStr(intPos + 1, mstrPath, "\", vbTextCompare)
        If intPos2 <> 0 Then
            strDir = Mid$(mstrPath, intPos + 1, intPos2 - intPos - 1)
            If Right$(strPath, 1) = "\" Then
                strPath = strPath & strDir
            Else
                strPath = strPath & "\" & strDir
            End If
            trvDir.Nodes(strPath).Selected = True
            intPos = intPos2
        Else
            If intPos <> Len(mstrPath) Then
                If Right$(strPath, 1) = "\" Then
                    strPath = strPath & Mid$(mstrPath, intPos + 1)
                Else
                    strPath = strPath & "\" & Mid$(mstrPath, intPos + 1)
                End If
                trvDir.Nodes(strPath).Selected = True
            End If
        End If
    Wend
    trvDir.SelectedItem.EnsureVisible
    GetFiles trvDir.SelectedItem

End Sub

Private Sub ShowProps(ByVal FileName As String)

Dim SEI As SHELLEXECUTEINFO
Dim R As Long

    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .lpVerb = "properties"
        .lpFile = FileName
        .nShow = 5
    End With
    R = ShellExecuteEX(SEI)

End Sub

Private Sub tbView_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "icons"
        lsvFiles.View = lvwIcon
        mnuView(0).Checked = True
        mnuView(1).Checked = False
        mnuView(2).Checked = False
        mnuView(3).Checked = False
    Case "sicons"
        lsvFiles.View = lvwSmallIcon
        mnuView(0).Checked = False
        mnuView(1).Checked = True
        mnuView(2).Checked = False
        mnuView(3).Checked = False
    Case "list"
        lsvFiles.View = lvwList
        mnuView(0).Checked = False
        mnuView(1).Checked = False
        mnuView(2).Checked = True
        mnuView(3).Checked = False
    Case Else
        lsvFiles.View = lvwReport
        mnuView(0).Checked = False
        mnuView(1).Checked = False
        mnuView(2).Checked = False
        mnuView(3).Checked = True
    End Select

End Sub

Private Sub trvDir_Collapse(ByVal Node As MSComctlLib.Node)

    If Node.Key = trvDir.SelectedItem.Key Then
        Call GetFiles(Node)
    End If

End Sub

Private Sub trvDir_Expand(ByVal Node As MSComctlLib.Node)

Dim hFind As Long
Dim w32FindData As WIN32_FIND_DATA
Dim blnFileExists As Boolean
Dim strFileName As String

    If Node.Tag = 1 Then Exit Sub  'don't search if already done
    frmDir.MousePointer = vbHourglass
    hFind = FindFirstFile(Node.Key & "\*.*", w32FindData)
    If hFind = INVALID_HANDLE_VALUE Then
        MsgBox "Invalid directory", vbCritical, "Refreshing..."
        trvDir_KeyDown vbKeyF5, 0
        Exit Sub
    End If
    blnFileExists = True
    While blnFileExists
        strFileName = Left$(w32FindData.cFileName, InStr(1, w32FindData.cFileName, vbNullChar) - 1)
        If LenB(strFileName) Then
            If AscW(strFileName) <> 46 Then
                If (w32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    GetSubDirs trvDir.Nodes(UCase$(Node.Key & "\" & strFileName))
                End If
            End If
        End If
        blnFileExists = FindNextFile(hFind, w32FindData)
    Wend
    FindClose hFind
    Node.Tag = 1
    frmDir.MousePointer = vbDefault

End Sub

Private Sub trvDir_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        mstrPath = trvDir.SelectedItem.Key
        mblnRefresh = True
        RefreshTreeView
        mblnRefresh = False
        SearchDir
    End If

End Sub

Private Sub trvDir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        If TypeName(trvDir.HitTest(x, y)) <> "Nothing" Then
            mstrOldName = trvDir.SelectedItem.Key
        End If
    End If

End Sub

Private Sub trvDir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If TypeName(trvDir.HitTest(x, y)) <> "Nothing" Then
            If UCase$(Left$(trvDir.SelectedItem.Text, 7)) = "RECYCLE" Then
                PopupMenu mnuDir3
            Else
                PopupMenu mnuDir2
            End If
        End If
    End If

End Sub

Private Sub trvDir_NodeClick(ByVal Node As MSComctlLib.Node)

    If mstrOldName <> Node.Key Then
        Call GetFiles(Node)
    End If

End Sub
