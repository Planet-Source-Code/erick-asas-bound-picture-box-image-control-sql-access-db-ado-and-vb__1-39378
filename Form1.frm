VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Images, VB & ADO"
   ClientHeight    =   5445
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7620
   Begin VB.TextBox Text1 
      DataField       =   "DESCRIPTION"
      Height          =   315
      Left            =   2130
      TabIndex        =   12
      Top             =   510
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load From File"
      Height          =   345
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   7515
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   7620
      TabIndex        =   0
      Top             =   4815
      Width           =   7620
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   600
         Left            =   1200
         TabIndex        =   13
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   600
         Left            =   6030
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   600
         Left            =   6405
         Picture         =   "Form1.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   600
         Left            =   6780
         Picture         =   "Form1.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   600
         Left            =   7140
         Picture         =   "Form1.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   600
         Left            =   1200
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   600
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   600
         Left            =   4675
         TabIndex        =   4
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   600
         Left            =   3521
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   600
         Left            =   2367
         TabIndex        =   2
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   600
         Left            =   59
         TabIndex        =   1
         Top             =   30
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Image Description"
      Height          =   345
      Left            =   90
      TabIndex        =   14
      Top             =   510
      Width           =   1905
   End
   Begin VB.Image Image1 
      DataField       =   "IMAGES"
      Height          =   3825
      Left            =   60
      Stretch         =   -1  'True
      Top             =   930
      Width           =   7485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author - Erick L. Asas
'Date 8/2/2002
'Notes:
'The code here is self-explanatory. The only point discussing here is the manner
'in which the file was saved to the database by the use of a byte array. This code
'still has many bugs but what I show here is the way I solve my problem of binding
'an image box or picture box to a SQL Server( or MS ACCESS database).
'---------------------------------------------------------------------------------------------
'API For Opening the open dialog box
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'API For Getting the File Size
Private Const OF_READ = &H0&
Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private lpFSHigh As Long

'Private declarations
Private WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Private mbChangedByCode As Boolean
Private mvBookMark As Variant
Private mbAddNewFlag As Boolean
Private mbEditFlag As Boolean
Private mbDataChanged As Boolean
Private strfilepath As String

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub Command1_Click()
On Error GoTo ET

    strfilepath = GetFile(Me)
    If strfilepath <> "" Then
      Image1.Picture = LoadPicture(strfilepath)
    End If
  Exit Sub
ET:
  Exit Sub
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  'Connection String
  'Definition for MS ACCESS database - comment this if using SQL Server
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & App.Path & "\db1.mdb;"
  
  'Definition for SQL Server - uncomment if using SQL Server
  'struid = Admin 'username
  'strpwd = Admin 'password
  'db.Open "PROVIDER=MSDataShape;Data PROVIDER=MSDASQL;" & _
    "dsn= Your_DSN_Name;uid=" & _
      struid & ";pwd=" & _
      strpwd & ";database=Your_Database_Name;"
            
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "tbl_images", db, adOpenStatic, adLockOptimistic

  'Bind the ole controls to the data provider
  Set Text1.DataSource = adoPrimaryRS
  Set Image1.DataSource = adoPrimaryRS

  mbDataChanged = False
  
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbAddNewFlag = False
  mbEditFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  With adoPrimaryRS
      SaveBitmap adoPrimaryRS, "IMAGES", strfilepath
      .UpdateBatch adAffectAll
  End With


  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If
  
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdEdit.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  Text1.Locked = bVal
  
End Sub

Private Function GetFile(ByRef frm As Form) As String
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = frm.hWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    OFName.lpstrFilter = "Bitmap (*.bmp)" + Chr$(0) + _
      "*.bmp" + Chr$(0) + _
      "Jpg (*.jpg)" + Chr$(0) + _
      "*.jpg" + Chr$(0) + _
      "Icons (*.ico)" + Chr$(0) + _
      "*.ico" + Chr$(0) + _
      "Windows Metafiles (*.wmf)" + Chr$(0) + _
      "*.wmf" + Chr$(0) + _
      "Jpeg (*.jpeg)" + Chr$(0) + _
      "*.jpeg" + Chr$(0) + _
      "Gif (*.gif)" + Chr$(0) + _
      "*.gif" + Chr$(0) + _
      "All Files (*.*)" + Chr$(0) + _
      "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    'OFName.lpstrInitialDir = "C:\" 'Commented so that the box opens on the last directory browsed
    'Set the title
    OFName.lpstrTitle = "Open Dialog Box"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
      GetFile = Trim$(OFName.lpstrFile)
    Else
      GetFile = ""
    End If

End Function

Public Sub SaveBitmap(ByRef adoRS As ADODB.Recordset, ByVal strField As String, ByVal SourceFile As String)
'This sub copies the actual file into a byte array.
'This byte array is then used as the value for
'the field having an image data type

    Dim Arr() As Byte
    Dim Pointer As Long
    Dim SizeOfThefile As Long
    
    Pointer = lOpen(SourceFile, OF_READ)
    'size of the file
    SizeOfThefile = GetFileSize(Pointer, lpFSHigh)
    lclose Pointer

    'Resize the array, then fill it with
    'the entire contents of the field
    ReDim Arr(SizeOfThefile)

    Open SourceFile For Binary Access Read As #1
    Get #1, , Arr
    Close #1
    adoRS(strField).Value = Arr
    Exit Sub
    
End Sub
