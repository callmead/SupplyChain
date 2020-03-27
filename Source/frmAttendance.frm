VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmAttendance.frx":0E42
   ScaleHeight     =   7155
   ScaleWidth      =   9870
   Begin VB.CommandButton cmdSelectEmp 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtTimeOut 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   7320
      MaxLength       =   30
      TabIndex        =   6
      Text            =   "txtTimeOut"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtR 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1035
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmAttendance.frx":6802B
      Top             =   1920
      Width           =   6615
   End
   Begin VB.TextBox txtAID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      Text            =   "txtAID"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtTimeIn 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "txtTimeIn"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtEmpId 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Text            =   "txtEmpID"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7680
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Text            =   "txtSearch"
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox ST 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      ItemData        =   "frmAttendance.frx":68030
      Left            =   3600
      List            =   "frmAttendance.frx":6803A
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "EmpId"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdRDB 
      Caption         =   "Re&fresh DB"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdML 
      Caption         =   "Move &Last"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdN 
      Caption         =   "Ne&xt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdMF 
      Caption         =   "Move &First"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      DefColWidth     =   73
      Enabled         =   -1  'True
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   28
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   27
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   26
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp. Id"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      Caption         =   "AID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   480
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   9600
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim A_ID, SQL As String
'
'Private Sub cmdSelectEmp_Click()
'    ParentForm = "frmAttendance"
'    GridSQLString = "Select * from Employees ORDER BY EmpID"
'    SelectedField = 0
'    frmDataSelect.Show
'End Sub
'
'Private Sub Form_Load()
'    Connect
'    SQLString = "SELECT * FROM Attendance ORDER BY Date"
'    ShowAttnData (SQLString)
'    ShowAttnGrid
'
'    ClearFields
'
'    Normalize
'    txtSearch.Text = ""
'End Sub
'
'Private Sub cmdNew_Click()
'    EnterNewAttn
'End Sub
'
'Private Sub cmdAdd_Click()
''Checking Fields for Records
'    If (txtAID.Text = "" Or txtAID.Text = " ") Then
'        MsgBox "Enter Attendance ID !!!", vbOKOnly, "Information Required"
'        txtAID.SetFocus
'        Exit Sub
'    End If
'    If (txtTimeIn.Text = "" Or txtTimeIn.Text = " ") Then
'        MsgBox "Enter Time In !!!", vbOKOnly, "Information Required"
'        txtTimeIn.SetFocus
'        Exit Sub
'    End If
'    If (txtTimeOut.Text = "" Or txtTimeOut.Text = " ") Then
'        MsgBox "Enter Time Out !!!", vbOKOnly, "Information Required"
'        txtTimeOut.SetFocus
'        Exit Sub
'    End If
'    If (txtR.Text = "") Then txtR.Text = "-"
'
'    'Updating Database
'    If DupCheck(txtAID.Text) = True Then
'        MsgBox "Attendance ID Already Exists !!! ", , "ADMIN"
'    Else
'        SQL = "INSERT INTO Attendance(AID,Date,EmpID,TimeIn,TimeOut,Remarks) values('" & txtAID & "','" & txtDate & "','" & txtEmpId & "','" & txtTimeIn & "','" & txtTimeOut & "','" & txtR & "')"
'        'MsgBox sql
'        conn.Execute SQL
'    End If
'
'    Normalize
'    cmdRDB_Click
'    cmdNew.SetFocus
'    Exit Sub
'
'End Sub
'
'Private Sub cmdEdit_Click()
'    EnableFields
'    DisableButtons
'    txtSearch.Enabled = False
'    ST.Enabled = False
'    cmdEdit.Visible = False
'    cmdDelete.Visible = False
'    cmdCancel.Enabled = True
'    cmdSave.Enabled = True
'End Sub
'
'Private Sub cmdSave_Click()
'    SQL = "UPDATE Attendance SET Date='" & txtDate.Text & "',EmpID='" & txtEmpId.Text & "',TimeIn='" & txtTimeIn.Text & "',TimeOut='" & txtTimeOut.Text & "',Remarks='" & txtR.Text & "' Where AID='" & txtAID.Text & "'"
'    conn.Execute SQL
'    ShowAttnData (SQLString)
'    Set DataGrid1.DataSource = RsAttnGrid
'    ShowAttnGrid
'    DataGrid1.Row = Rx
'
'    Normalize
'    cmdRDB_Click
'
'End Sub
'
'Private Sub cmdCancel_Click()
'    Normalize
'End Sub
'
'Private Sub cmdDelete_Click()
'    Dim rsTemp As ADODB.Recordset
'    Set rsTemp = New ADODB.Recordset
'    rsTemp.CursorType = adOpenDynamic
'    rsTemp.LockType = adLockOptimistic
'    rsTemp.CursorLocation = adUseServer
'    SQL = "DELETE FROM Attendance Where AID='" & txtAID.Text & "'"
'
'    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
'        Set rsTemp = Nothing
'        Exit Sub
'    End If
'    rsTemp.Open SQL, conn, adOpenKeyset
'    Set rsTemp = Nothing
'
'    Rx = Rx - 1
'    Normalize
'    cmdRDB_Click
'    Set DataGrid1.DataSource = RsAttnGrid
'    ShowAttnGrid
'    If (Rx <> 0) Then DataGrid1.Row = Rx
'    ClearFields
'End Sub
'
'Private Sub cmdClose_Click()
'Unload Me
'End Sub
'
'Private Sub cmdRDB_Click()
'    SQLString = "SELECT * FROM Attendance ORDER BY AID"
'    Rx = 0
'    ShowAttnData (SQLString)
'    ShowAttnGrid
'End Sub
'
'Private Sub cmdMF_Click()
'    On Error Resume Next
'    Rx = 0
'    ShowAttnData (SQLString)
'    ShowAttnGrid
'    DataGrid1.Row = Rx
'End Sub
'
'Private Sub cmdML_Click()
'    On Error Resume Next
'    Rx = xCount - 1
'    ShowAttnData (SQLString)
'    ShowAttnGrid
'    DataGrid1.Row = Rx
'End Sub
'
'Private Sub cmdN_Click()
'    On Error Resume Next
'    Rx = Rx + 1
'    ShowAttnData (SQLString)
'    ShowAttnGrid
'    DataGrid1.Row = Rx
'End Sub
'
'Private Sub cmdP_Click()
'    On Error Resume Next
'    Rx = Rx - 1
'    ShowAttnData (SQLString)
'    ShowAttnGrid
'    DataGrid1.Row = Rx
'End Sub
'
'Private Sub cmdSearch_Click()
'If (txtSearch.Text = "" Or txtSearch.Text = " ") Then
'    MsgBox "Search what?", vbExclamation, "ADMIN"
'    txtSearch.SetFocus
'    SendKeys "{Home}+{End}"
'    Exit Sub
'End If
'
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseServer
'
'    SQLString = "SELECT * FROM Attendance WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
'    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
'
'    Set RsAttnGrid = New ADODB.Recordset
'    RsAttnGrid.CursorLocation = adUseClient
'    RsAttnGrid.CursorType = adOpenStatic
'    RsAttnGrid.LockType = adLockReadOnly
'    RsAttnGrid.Open SQLString, conn
'    Set DataGrid1.DataSource = RsAttnGrid
'
'    If rs.EOF = True Then
'        rs.Close
'        Set rs = Nothing
'
'        MsgBox "Record Not Found !!!", vbInformation, ""
'        txtSearch.SetFocus
'        SendKeys "{Home}+{End}"
'        cmdRDB_Click
'        Exit Sub
'    End If
'    If IsNull(rs!AID) Then
'        ClearFields
'    Else
'
'    txtAID.Text = rs!AID
'    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
'    txtEmpId.Text = rs!EmpID
'    txtTimeIn.Text = rs!TimeIn
'    txtTimeOut.Text = rs!TimeOut
'    txtR.Text = rs!Remarks
'
'    End If
'    rs.Close
'    Set rs = Nothing
'
'End Sub
'
'Private Sub ClearFields()
'    txtAID.Text = ""
'    txtDate.Text = ""
'    txtEmpId.Text = ""
'    txtTimeIn.Text = ""
'    txtTimeOut.Text = ""
'    txtR.Text = ""
'End Sub
'
'Private Sub EnableFields()
'    txtDate.Enabled = True
'    txtTimeIn.Enabled = True
'    txtTimeOut.Enabled = True
'    txtR.Enabled = True
'End Sub
'Private Sub DisableFields()
'    txtDate.Enabled = False
'    txtTimeIn.Enabled = False
'    txtTimeOut.Enabled = False
'    txtR.Enabled = False
'End Sub
'
'Private Sub EnableButtons()
'    cmdNew.Enabled = True
'    cmdAdd.Enabled = True
'    cmdEdit.Enabled = True
'    cmdSave.Enabled = True
'    cmdCancel.Enabled = True
'    cmdDelete.Enabled = True
'    cmdRDB.Enabled = True
'    cmdMF.Enabled = True
'    cmdN.Enabled = True
'    cmdP.Enabled = True
'    cmdML.Enabled = True
'    cmdSearch.Enabled = True
'    cmdClose.Enabled = True
'
'    cmdSelectEmp.Enabled = False
'End Sub
'Private Sub DisableButtons()
'    cmdNew.Enabled = False
'    cmdAdd.Enabled = False
'    cmdEdit.Enabled = False
'    cmdSave.Enabled = False
'    cmdCancel.Enabled = False
'    cmdDelete.Enabled = False
'    cmdRDB.Enabled = False
'    cmdMF.Enabled = False
'    cmdN.Enabled = False
'    cmdP.Enabled = False
'    cmdML.Enabled = False
'    cmdSearch.Enabled = False
'    cmdClose.Enabled = False
'
'    cmdSelectEmp.Enabled = True
'End Sub
'
'Private Sub Normalize()
'    DisableFields
'    EnableButtons
'    cmdNew.Visible = True
'    cmdEdit.Visible = True
'    cmdDelete.Visible = True
'    'cmdNew.SetFocus
'    txtSearch.Enabled = True
'    ST.Enabled = True
'    cmdRDB_Click
'End Sub
'
'Public Sub EnterNewAttn()
'
'    DisableButtons
'    EnableFields
'    txtSearch.Enabled = False
'    ST.Enabled = False
'    cmdNew.Visible = False
'    cmdDelete.Visible = False
'    cmdCancel.Enabled = True
'    cmdAdd.Enabled = True
'    ClearFields
'    GenerateID
'    GetDate
'    txtDate.Text = DateToday
'    cmdSelectEmp.SetFocus
'
'End Sub
'Private Sub GenerateID()
'    txtAID.Text = "A" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
'End Sub
'
'Private Sub ST_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
'End Sub
'
'Private Function DupCheck(chkID As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseServer
'    rs.CursorType = adOpenStatic
'    rs.LockType = adLockOptimistic
'
'    SQL = "SELECT * from Attendance WHERE AID='" & chkID & "'"
'    rs.Open SQL, conn
'    If rs.EOF = True Then
'        rs.Close
'        Set rs = Nothing
'        Exit Function
'    End If
'    If chkID = rs!AID Then
'        DupCheck = True
'    Else
'        DupCheck = False
'    End If
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Private Sub ST_LostFocus()
'    If ST.Text = "Date" Then
'        txtSearch.ToolTipText = "Date Format YYYY-MM-DD"
'        txtSearch.Text = "2006-03-03"
'    Else
'        Exit Sub
'    End If
'End Sub
'
'Private Sub txtSearch_GotFocus()
'    SendKeys "{Home}+{End}"
'End Sub
Private Sub cmdSelectEmp_Click()

End Sub
