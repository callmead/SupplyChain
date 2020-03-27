VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSale2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: GENERAL SALE :."
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "frmSale2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSale2.frx":0BC2
   ScaleHeight     =   7710
   ScaleWidth      =   10740
   Begin VB.TextBox txtSalesman 
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
      Left            =   6600
      TabIndex        =   34
      Text            =   "txtSalesman"
      Top             =   720
      Width           =   2295
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
      Left            =   8520
      TabIndex        =   20
      Top             =   4680
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
      TabIndex        =   17
      Text            =   "txtSearch"
      Top             =   4680
      Width           =   4095
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
      Left            =   6480
      TabIndex        =   19
      Top             =   4680
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
      ItemData        =   "frmSale2.frx":C1006
      Left            =   4440
      List            =   "frmSale2.frx":C1016
      Sorted          =   -1  'True
      TabIndex        =   18
      Text            =   "Customer"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   6600
      TabIndex        =   6
      Text            =   "txtPrice"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtCustomer 
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
      Left            =   2280
      TabIndex        =   2
      Text            =   "txtCustomer"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   6600
      TabIndex        =   4
      Text            =   "txtDescription"
      Top             =   1560
      Width           =   2295
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
      Left            =   9000
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   22
      Top             =   600
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   16
      Top             =   4080
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   0
      Top             =   240
      Width           =   1455
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
      Height          =   360
      Left            =   6600
      TabIndex        =   1
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtInvNo 
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
      Left            =   2280
      TabIndex        =   21
      Text            =   "txtInvNo"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtQuantity 
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
      Left            =   2280
      TabIndex        =   5
      Text            =   "txtQuantity"
      Top             =   2040
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
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmSale2.frx":C1040
      Top             =   3360
      Width           =   6615
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
      Left            =   9000
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtProduct 
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
      Left            =   2280
      TabIndex        =   3
      Text            =   "txtProduct"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   2280
      TabIndex        =   7
      Text            =   "txtTotal"
      Top             =   2520
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   23
      Top             =   5160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      DefColWidth     =   93
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
      Left            =   9000
      TabIndex        =   9
      Top             =   240
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   24
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblSalesman 
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman"
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
      Left            =   4920
      TabIndex        =   35
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   4920
      TabIndex        =   33
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice #"
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
      Left            =   360
      TabIndex        =   32
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Left            =   360
      TabIndex        =   31
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblCustomer 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      Left            =   360
      TabIndex        =   30
      Top             =   720
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
      Left            =   360
      TabIndex        =   29
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblDate 
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
      Left            =   4920
      TabIndex        =   28
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   4920
      TabIndex        =   27
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   10440
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
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
      Left            =   360
      TabIndex        =   26
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   360
      TabIndex        =   25
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frmSale2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sup_ID, sql As String
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock As Boolean

Private Sub Form_Load()

    Connect
    GetDate
    
    SQLString = "SELECT * FROM G_Sale ORDER BY Invoice_No"
    ShowGSaleData (SQLString)
    ShowGSaleGrid (SQLString)
    
    ClearFields
    
    Normalize
    txtSearch.Text = ""
    
    'For Int TextBoxes
    Dim tmp, tmp1, tmp2 As Long
    tmp = SetWindowLong(txtPrice.hwnd, GWL_STYLE, GetWindowLong(txtPrice.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp1 = SetWindowLong(txtTotal.hwnd, GWL_STYLE, GetWindowLong(txtTotal.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp2 = SetWindowLong(txtQuantity.hwnd, GWL_STYLE, GetWindowLong(txtQuantity.hwnd, GWL_STYLE) Or ES_NUMBER)
    
End Sub

Private Sub cmdNew_Click()
    EnterNewG_Sale
    txtPrice.Text = "0"
    txtTotal.Text = "0"
    txtQuantity.Text = "0"
    txtSalesman.Text = UserName
End Sub

Private Sub cmdAdd_Click()
    
    'Checking Fields for Records
    If (txtCustomer.Text = "" Or txtCustomer.Text = " ") Then
        MsgBox "Please provide Customer Name !!!", vbOKOnly, "Information Required"
        txtCustomer.SetFocus
        Exit Sub
    End If
    If (txtProduct.Text = "" Or txtProduct.Text = " ") Then
        MsgBox "Please provide a Product !!!", vbOKOnly, "Information Required"
        txtProduct.SetFocus
        Exit Sub
    End If
    If (txtQuantity.Text = "" Or txtQuantity.Text = " ") Then
        MsgBox "Please provide Quantity for the given Product !!!", vbOKOnly, "Information Required"
        txtQuantity.SetFocus
        Exit Sub
    End If
    If (txtPrice.Text = "" Or txtPrice.Text = " ") Then
        MsgBox "Please provide Price for the given Product !!!", vbOKOnly, "Information Required"
        txtPrice.SetFocus
        Exit Sub
    End If
    If (txtDescription.Text = "" Or txtDescription.Text = " ") Then txtDescription.Text = "-"
    If (txtR.Text = "") Then txtR.Text = "-"
    
    'Updating Database
    If DupCheck("SELECT * FROM G_Sale WHERE Invoice_No='" & txtInvNo.Text & "'") = True Then
        MsgBox "Invoice Already Exists !!! ", , "General Error"
    Else
        sql = "INSERT INTO G_Sale VALUES('" & txtInvNo & "','" & txtDate & "','" & txtCustomer & "','" & txtSalesman & "','" & txtProduct & "','" & txtDescription & "'," & txtQuantity & "," & txtPrice & "," & txtTotal & ",'" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
    End If
        
    Normalize
    cmdRDB_Click
    cmdNew.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdEdit_Click()
    SetFields (True)
    txtCustomer.SetFocus
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE G_Sale SET Date='" & txtDate.Text & "',Customer='" & txtCustomer.Text & "',Customer='" & txtSalesman.Text & "',Product='" & txtProduct.Text & "',Description='" & txtDescription.Text & "',Quantity=" & txtQuantity.Text & ",Price=" & txtPrice.Text & ",Total=" & txtTotal.Text & ",Remarks='" & txtR.Text & "' Where Invoice_No='" & txtInvNo.Text & "'"
    conn.Execute sql
    ShowGSaleData (SQLString)
    Set DataGrid1.DataSource = RsGSaleGrid
    ShowGSaleGrid ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    DataGrid1.Row = Rx

    Normalize
    cmdRDB_Click
    
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorLocation = adUseServer
    sql = "DELETE FROM G_Sale Where Invoice_No='" & txtInvNo.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    rsTemp.Open sql, conn, adOpenKeyset
    Set rsTemp = Nothing
    
    Rx = Rx - 1
    Normalize
    cmdRDB_Click
    Set DataGrid1.DataSource = RsGSaleGrid
    ShowGSaleGrid ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    If (Rx <> 0) Then DataGrid1.Row = Rx
    ClearFields
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    SQLString = "SELECT * FROM G_Sale ORDER BY Invoice_No"
    Rx = 0
    ShowGSaleData (SQLString)
    ShowGSaleGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowGSaleData ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    ShowGSaleGrid ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowGSaleData ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    ShowGSaleGrid ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowGSaleData ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    ShowGSaleGrid ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowGSaleData ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    ShowGSaleGrid ("SELECT * FROM G_Sale ORDER BY Invoice_No")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdSearch_Click()
If (txtSearch.Text = "" Or txtSearch.Text = " ") Then
    MsgBox "Search what?", vbExclamation, "General Error"
    txtSearch.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    
    SQLString = "SELECT * FROM G_Sale WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RsGSaleGrid = New ADODB.Recordset
    RsGSaleGrid.CursorLocation = adUseClient
    RsGSaleGrid.CursorType = adOpenStatic
    RsGSaleGrid.LockType = adLockReadOnly
    RsGSaleGrid.Open SQLString, conn
    Set DataGrid1.DataSource = RsGSaleGrid
      
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        
        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!Invoice_No) Then
        ClearFields
    Else
       
    txtInvNo.Text = rs!Invoice_No
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    txtCustomer.Text = rs!Customer
    txtSalesman.Text = rs!Salesman
    txtProduct.Text = rs!Product
    txtDescription.Text = rs!Description
    txtQuantity.Text = rs!Quantity
    txtPrice.Text = rs!Price
    txtTotal.Text = rs!Total
    txtR.Text = rs!Remarks
    
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub ClearFields()
    txtInvNo.Text = ""
    txtDate.Text = ""
    txtCustomer.Text = ""
    txtSalesman.Text = ""
    txtProduct.Text = ""
    txtDescription.Text = ""
    txtQuantity.Text = ""
    txtPrice.Text = ""
    txtTotal.Text = ""
    txtR.Text = ""
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtDate.Enabled = TextFieldLock
    txtCustomer.Enabled = TextFieldLock
    txtProduct.Enabled = TextFieldLock
    txtDescription.Enabled = TextFieldLock
    txtQuantity.Enabled = TextFieldLock
    txtPrice.Enabled = TextFieldLock
    'txtTotal.Enabled = TextFieldLock
    txtR.Enabled = TextFieldLock
End Sub

Private Sub SetButtons(ButtonLock As Boolean)
    cmdNew.Enabled = ButtonLock
    cmdAdd.Enabled = ButtonLock
    cmdEdit.Enabled = ButtonLock
    cmdSave.Enabled = ButtonLock
    cmdCancel.Enabled = ButtonLock
    cmdDelete.Enabled = ButtonLock
    cmdRDB.Enabled = ButtonLock
    cmdMF.Enabled = ButtonLock
    cmdN.Enabled = ButtonLock
    cmdP.Enabled = ButtonLock
    cmdML.Enabled = ButtonLock
    cmdSearch.Enabled = ButtonLock
    cmdClose.Enabled = ButtonLock
End Sub

Private Sub Normalize()
    SetFields (False)
    SetButtons (True)
    cmdNew.Visible = True
    cmdEdit.Visible = True
    cmdDelete.Enabled = True

    txtSearch.Enabled = True
    ST.Enabled = True
    cmdRDB_Click
End Sub

Public Sub EnterNewG_Sale()
    
    SetButtons (False)
    SetFields (True)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True
    ClearFields
    GenerateID
    GetDate
    txtDate.Text = DateToday
    txtCustomer.SetFocus
    
End Sub
Private Sub GenerateID()
    txtInvNo.Text = "I" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Sub ST_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    'SQL = "SELECT * FROM G_Sale WHERE Invoice_No='" & chkID & "'"
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtInvNo.Text = rs!Invoice_No Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub ST_LostFocus()
    If ST.Text = "Date" Then
        txtSearch.ToolTipText = "Date Format YYYY-MM-DD"
        txtSearch.Text = "2006-03-03"
    Else
        Exit Sub
    End If
End Sub

Private Sub txtCustomer_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtDescription_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPrice_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPrice_KeyUp(KeyCode As Integer, Shift As Integer)
    txtTotal.Text = Val(txtQuantity.Text) * Val(txtPrice.Text)
End Sub

Private Sub txtProduct_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtQuantity_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtQuantity_KeyUp(KeyCode As Integer, Shift As Integer)
txtTotal.Text = Val(txtQuantity.Text) * Val(txtPrice.Text)
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtTotal_GotFocus()
SendKeys "{Home}+{End}"
End Sub
