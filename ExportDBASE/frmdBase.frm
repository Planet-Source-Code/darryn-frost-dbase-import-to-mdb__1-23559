VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportDBASE 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "dBase to MDB Import"
   ClientHeight    =   5010
   ClientLeft      =   2430
   ClientTop       =   1575
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   330
      Left            =   1320
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "New or existing Database Name and Location:"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   6495
      Begin VB.TextBox txtDestPath 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Find or create an MDB"
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmdSetDest 
         Caption         =   "Find..."
         Height          =   330
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin VB.ListBox lstFields 
         Height          =   1425
         Left            =   3960
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         ToolTipText     =   "Select fields to import, 'ALL' is the default"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ListBox lstTables 
         Height          =   1425
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Choose the table you wish to import"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.CommandButton cmdSetSource 
         Caption         =   "Find..."
         Height          =   330
         Left            =   5520
         TabIndex        =   4
         Top             =   600
         Width           =   765
      End
      Begin VB.TextBox txtSourcePath 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Choose the file you wish to import from.... "
         Top             =   600
         Width           =   4980
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fields:"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Source Tables:"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Source file:"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog dlgDB 
      Left            =   120
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Export"
      Height          =   330
      Left            =   3120
      TabIndex        =   1
      Top             =   4440
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   4800
      TabIndex        =   0
      Top             =   4440
      Width           =   1065
   End
End
Attribute VB_Name = "frmImportDBASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private m_Fields As String
Private m_Filename As String
Private m_SourceTableName As String
Private m_DestTableName
Private m_SourceConn As ADODB.Connection
Private m_DestConn As ADODB.Connection
Private m_SourceCat As ADOX.Catalog
Private m_DestCat As ADOX.Catalog

Private Sub cmdSetSource_Click()

With dlgDB
    .DialogTitle = "Find Source File"
    .Flags = FileOpenConstants.cdlOFNPathMustExist Or FileOpenConstants.cdlOFNHideReadOnly
    .Filter = "dBase Files (*.dbf)|*.dbf"
    .filename = ""
    .FilterIndex = 1
    .ShowOpen
End With

If Len(dlgDB.filename) > 0 Then
    txtSourcePath = dlgDB.filename
Else
    Exit Sub
End If

If m_SourceConn.State = adStateOpen Then m_SourceConn.Close

On Error Resume Next

m_SourceConn.Open "Driver={Microsoft dBase Driver (*.dbf)};" _
                    & "DriverID=277;" _
                    & "Dbq=" & StripFileName(txtSourcePath) & ";"

m_SourceCat.ActiveConnection = m_SourceConn

If Err.Number <> 0 Then
     MsgBox "Error# " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Exit Sub
End If

ListSourceTables

End Sub

Private Sub ListSourceTables()
Dim cTable As ADOX.Table
For Each cTable In m_SourceCat.Tables
     lstTables.AddItem cTable.Name
Next
Set cTable = Nothing
End Sub

Private Sub cmdClear_Click()
txtSourcePath.Text = ""
txtDestPath.Text = ""
lstTables.Clear
lstFields.Clear

If m_SourceConn.State = adStateOpen Then m_SourceConn.Close
If m_DestConn.State = adStateOpen Then m_DestConn.Close
cmdDoIt.Enabled = False
End Sub

Private Sub lstTables_Click()
Dim i As Integer
For i = 0 To lstTables.ListCount - 1
    If lstTables.Selected(i) = True Then
        m_SourceTableName = lstTables.List(i)
        m_DestTableName = InputBox("If you want you may change the imported name of this table", "Enter table name", m_SourceTableName)
        If m_DestTableName = "" Then m_DestTableName = m_SourceTableName
    End If
Next i

For i = 0 To m_SourceCat.Tables(m_SourceTableName).Columns.Count - 1
       lstFields.AddItem m_SourceCat.Tables(m_SourceTableName).Columns(i).Name
Next i

If m_SourceTableName <> "" And txtDestPath <> "" Then cmdDoIt.Enabled = True

End Sub

Private Sub cmdSetDest_Click()

Dim ync As VbMsgBoxResult

With dlgDB
    .DialogTitle = "Set destination file"
    .Flags = FileOpenConstants.cdlOFNHideReadOnly
    .Filter = "Access Database Files (*.mdb)|*.mdb"
    .filename = ""
    .FilterIndex = 1
    .ShowSave
End With

On Error Resume Next

If Len(dlgDB.filename) > 0 Then
    txtDestPath = dlgDB.filename
    If Dir(txtDestPath) = "" Then
        ync = MsgBox("Would you like to Create an Access 2000 DB? " _
              & "(Default is 97)", vbYesNoCancel)
        If ync = vbYes Then
            m_DestCat.Create "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDestPath
            MsgBox "New Access 2000 mdb created as: " & txtDestPath
            m_DestCat.ActiveConnection = Nothing
        ElseIf ync = vbNo Then
            m_DestCat.Create "Provider = Microsoft.Jet.OLEDB.3.51;Data Source=" & txtDestPath
            MsgBox "New Access 97 mdb created as: " & txtDestPath
            m_DestCat.ActiveConnection = Nothing
        Else
            Exit Sub
        End If
    End If
Else
    Exit Sub
End If

If m_SourceTableName <> "" And txtDestPath <> "" Then cmdDoIt.Enabled = True

If Err.Number <> 0 Then
    MsgBox "Error# " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Exit Sub
End If

End Sub

Private Sub cmdClose_Click()
Set m_SourceConn = Nothing
Set m_DestConn = Nothing
Set m_SourceCat = Nothing
Set m_DestCat = Nothing
Unload Me
End Sub

Private Sub cmdDoIt_Click()
Dim rsSource As ADODB.Recordset
Dim rsDest As ADODB.Recordset
Dim stable As ADOX.Table
Dim dTable As ADOX.Table
Dim i As Integer
Dim j As Long
Dim intRecCount As Long

On Error Resume Next

m_DestConn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDestPath
m_DestConn.Open
m_DestCat.ActiveConnection = m_DestConn

Set stable = m_SourceCat.Tables(m_SourceTableName)

Set dTable = New ADOX.Table

dTable.Name = m_DestTableName

m_Fields = vbNullString


If lstFields.SelCount > 0 Then
    For i = 0 To lstFields.ListCount - 1
        If lstFields.Selected(i) = True Then
            dTable.Columns.Append stable.Columns(lstFields.List(i)).Name, , stable.Columns(lstFields.List(i)).DefinedSize
            If m_Fields = "" Then
                m_Fields = lstFields.List(i)
            Else
                m_Fields = m_Fields & ", " & lstFields.List(i)
            End If
        End If
    Next i
Else
    m_Fields = "*"
    For i = 0 To stable.Columns.Count - 1
        dTable.Columns.Append stable.Columns(i).Name, , stable.Columns(i).DefinedSize
    Next i
End If

tryagain:

m_DestCat.Tables.Append dTable

If Err.Number = -2147217857 Then
    If MsgBox("There is already a table named " & m_DestTableName & vbCrLf _
            & "Would you like to delete it?", vbYesNo) = vbYes Then
            m_DestCat.Tables.Delete dTable.Name  'just in case there is already a table with this name
            Err.Clear
            GoTo tryagain
    Else
        Exit Sub
    End If
End If

m_DestCat.ActiveConnection = Nothing
m_SourceCat.ActiveConnection = Nothing

If Err.Number <> 0 Then
    MsgBox "Error# " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Exit Sub
End If

On Error GoTo Transaction_error

Set rsSource = New ADODB.Recordset
Set rsDest = New ADODB.Recordset

With rsSource
    .ActiveConnection = m_SourceConn
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open "SELECT " & m_Fields & " FROM " & m_SourceTableName
End With

With rsDest
    .ActiveConnection = m_DestConn
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockBatchOptimistic
    .Open "SELECT " & m_Fields & " FROM " & m_DestTableName
End With
        
m_DestConn.BeginTrans
     
     For intRecCount = 1 To rsSource.RecordCount
            rsDest.AddNew
                For j = 0 To rsSource.Fields.Count - 1
                    rsDest.Fields(j) = "" & rsSource.Fields(j)
                Next j
            rsSource.MoveNext
     Next intRecCount
     
     rsDest.UpdateBatch
     rsDest.Close
     rsSource.Close
  
m_DestConn.CommitTrans
MsgBox "Importing successful", vbExclamation

m_SourceConn.Close
m_DestConn.Close

Set rsSource = Nothing
Set rsDest = Nothing

cmdClear_Click

Exit Sub

Transaction_error:
On Error Resume Next
    m_DestConn.RollbackTrans
    m_SourceConn.Close
    m_DestConn.Close
  MsgBox "Error importing. All changes to new table have been rolled back.", vbCritical
  
  Exit Sub
  
End Sub

Private Sub Form_Load()
Set m_SourceConn = New ADODB.Connection
Set m_DestConn = New ADODB.Connection
Set m_SourceCat = New ADOX.Catalog
Set m_DestCat = New ADOX.Catalog
cmdClear_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set m_SourceConn = Nothing
Set m_DestConn = Nothing
Set m_SourceCat = Nothing
Set m_DestCat = Nothing
End Sub

'------------------------------------------------------------
'this function strips the file name from a path\file string
'------------------------------------------------------------
Private Function StripFileName(rsFileName As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = "\" Then
      Exit For
    End If
  Next

  StripFileName = Mid(rsFileName, 1, i - 1)

End Function

'Strip the extension off filenames for databases that only need a table name
Private Function stripExtension(ByVal filename As String) As String
Dim stmp As String
Dim i As Integer
For i = Len(filename) To 1 Step -1
      If Mid(filename, i, 1) = "\" Then
        Exit For
      End If
    Next
stmp = Mid(filename, i + 1, Len(filename))
    'strip off the extension
    For i = 1 To Len(stmp)
      If Mid(stmp, i, 1) = "." Then
        Exit For
      End If
    Next
    stripExtension = Left(stmp, i - 1)

End Function

