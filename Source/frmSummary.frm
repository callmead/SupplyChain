VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: DATABASE SUMMARY :."
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "frmSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSummary.frx":0BC2
   ScaleHeight     =   7590
   ScaleWidth      =   10245
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   240
      Top             =   6480
   End
   Begin VB.Timer TmPB 
      Interval        =   140
      Left            =   240
      Top             =   6960
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   6960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblpb 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label lblSDMS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   2760
      TabIndex        =   27
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblSDRP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Minus Stock"
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
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "ReORder Products"
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
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Details"
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
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   4320
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   9960
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label lblTGSA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblTSA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblTEA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   8040
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblTIA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblTSD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblTSBA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   8040
      TabIndex        =   17
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblTCD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblTCBA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "General Sale"
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
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   9600
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sale"
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
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   4320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expense"
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
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   9600
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Income"
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
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   4320
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Due"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bills Amount"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Accounts"
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
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   9600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Due"
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
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bills Amount"
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
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblPo 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Accounts"
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
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   240
      Y1              =   120
      Y2              =   7440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   4320
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblTCBA.Caption = "Waiting..."
    lblTCD.Caption = "Waiting..."
    lblTSBA.Caption = "Waiting..."
    lblTSD.Caption = "Waiting..."
    lblTIA.Caption = "Waiting..."
    lblTEA.Caption = "Waiting..."
    lblTSA.Caption = "Waiting..."
    lblTGSA.Caption = "Waiting..."
    lblSDRP.Caption = "Waiting..."
End Sub

Private Sub GetCustomerAccountInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Total_Bills_Amount) as 'TBA',SUM(Total_Due) as 'TD' FROM Customer;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If rs!TBA <> "" Then
        lblTCBA.Caption = rs!TBA
    Else
        lblTCBA.Caption = "0"
    End If
    If rs!TD <> "" Then
        lblTCD.Caption = rs!TD
    Else
        lblTCD.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetSupplierAccountInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Total_Bills_Amount) as 'TBA',SUM(Total_Due) as 'TD' FROM Supplier;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If rs!TBA <> "" Then
        lblTSBA.Caption = rs!TBA
        lblTSD.Caption = rs!TD
    Else
        lblTSBA.Caption = "0"
        lblTSD.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTotalIncometInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Amount) as 'Sum' FROM Income;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!Sum <> "" Then
        lblTIA.Caption = rs!Sum
    Else
        lblTIA.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTotalExpenseInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Amount) as 'Sum' FROM Expenditure;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If rs!Sum <> "" Then
        lblTEA.Caption = rs!Sum
    Else
        lblTEA.Caption = "0"
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTotalSaleInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Amount_Paid) as 'Sum' FROM Sales;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!Sum <> "" Then
        lblTSA.Caption = rs!Sum
    Else
        lblTSA.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTotalGSaleInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Price) as 'Sum' FROM G_Sale;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If rs!Sum <> "" Then
        lblTGSA.Caption = rs!Sum
    Else
        lblTGSA.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetStockROLInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT COUNT(*) as 'No' FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!No <> "" Then
        lblSDRP.Caption = rs!No
    Else
        lblSDRP.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetStockMinusInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT COUNT(*) as No FROM Stock WHERE Stock_In_Hand<=0;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!No <> "" Then
        lblSDMS.Caption = rs!No
    Else
        lblSDMS.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Timer2_Timer()
    lblpb.Caption = "Data Loaded"
    PB1.Visible = False
    
    GetCustomerAccountInfo
    GetSupplierAccountInfo
    GetTotalIncometInfo
    GetTotalExpenseInfo
    GetTotalSaleInfo
    GetTotalGSaleInfo
    GetStockROLInfo
    GetStockMinusInfo
    
End Sub

Private Sub TmPB_Timer()
    PB1.Value = PB1.Value + 5
    If (PB1.Value = PB1.Max) Then
        TmPB.Enabled = False
    End If
End Sub
