VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: STOCK MANAGEMENT :."
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   Icon            =   "frmStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmStock.frx":0BC2
   ScaleHeight     =   8295
   ScaleWidth      =   10695
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5640
      ScaleHeight     =   825
      ScaleWidth      =   3225
      TabIndex        =   39
      ToolTipText     =   "BarCode has been Copied to Clipboard!"
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtCompany 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Text            =   "txtCompany"
      Top             =   1680
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   8760
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   1335
      Left            =   8400
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.ComboBox PType 
      Enabled         =   0   'False
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
      ItemData        =   "frmStock.frx":C1006
      Left            =   6600
      List            =   "frmStock.frx":C1008
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "ProductType"
      Top             =   1200
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
      TabIndex        =   23
      Top             =   5280
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
      TabIndex        =   20
      Text            =   "txtSearch"
      Top             =   5280
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
      TabIndex        =   22
      Top             =   5280
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
      ItemData        =   "frmStock.frx":C100A
      Left            =   4440
      List            =   "frmStock.frx":C102C
      Sorted          =   -1  'True
      TabIndex        =   21
      Text            =   "Product"
      Top             =   5280
      Width           =   1935
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
      Top             =   1200
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
      Left            =   2280
      TabIndex        =   7
      Text            =   "txtDescription"
      Top             =   2160
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
      TabIndex        =   15
      Top             =   3120
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
      TabIndex        =   24
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
      TabIndex        =   19
      Top             =   4560
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
      TabIndex        =   14
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
      Left            =   2280
      TabIndex        =   2
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtPID 
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
      TabIndex        =   1
      Text            =   "txtPID"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtStock 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   9
      Text            =   "txtStock"
      Top             =   3000
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
      TabIndex        =   11
      Text            =   "frmStock.frx":C109D
      Top             =   3840
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
      TabIndex        =   18
      Top             =   4200
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
      TabIndex        =   17
      Top             =   3840
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
      TabIndex        =   16
      Top             =   3480
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
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPS 
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
      Text            =   "txtPS"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtPricePU 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   8
      Text            =   "txtPricePU"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtROL 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   10
      Text            =   "txtROL"
      Top             =   3000
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   25
      Top             =   5760
      Width           =   10215
      _ExtentX        =   18018
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
      Left            =   9000
      TabIndex        =   12
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
      TabIndex        =   26
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      TabIndex        =   38
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblPID 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      TabIndex        =   36
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblST 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock in Hand"
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
      TabIndex        =   35
      Top             =   3000
      Width           =   1815
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
      TabIndex        =   34
      Top             =   1200
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
      TabIndex        =   33
      Top             =   3840
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
      Left            =   360
      TabIndex        =   32
      Top             =   720
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
      Left            =   360
      TabIndex        =   31
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   8040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   10440
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblPT 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type"
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
      TabIndex        =   30
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblPS 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Size"
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
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblPPU 
      BackStyle       =   0  'Transparent
      Caption         =   "Price per unit"
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
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblROL 
      BackStyle       =   0  'Transparent
      Caption         =   "ReOrder Level"
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
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim rn As Integer
Dim nLastKeyAscii As Integer
Private iC, iR As Integer
Private TextFieldLock, ButtonLock, isChecked As Boolean

Private Sub Form_Load()

    Connect
    ClearFields
    Normalize
    GetDate
    
    CheckROL
    CheckMinusStock

    CheckB4Connect
    ShowStockData (SQLString)
    ShowStockGrid (SQLString)

End Sub

Private Sub CheckB4Connect()
    isChecked = False
    
    SQLString = "SELECT * FROM Stock ORDER BY Product_ID"
    If isReOrder = True Then SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
    If isStockMinus = True Then SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<0;"
End Sub

Private Sub cmdNew_Click()
    EnterNewProduct
    GetComboData
    RemoveComboDuplicates
    txtStock.Text = "0"
    txtROL.Text = "10"
End Sub

Private Sub cmdAdd_Click()

    'Checking Fields for Records
    If (txtPID.Text = "" Or txtPID.Text = " ") Then
        MsgBox "Please provide a Product ID !!!", vbOKOnly, "Information Required"
        txtPID.SetFocus
        Exit Sub
    End If
    If (txtProduct.Text = "" Or txtProduct.Text = " ") Then
        MsgBox "Please provide a Product Name !!!", vbOKOnly, "Information Required"
        txtProduct.SetFocus
        Exit Sub
    End If
    If (PType.Text = "" Or PType.Text = " ") Then
        MsgBox "Please provide Product Type for Product " + txtProduct.Text + " !!!", vbOKOnly, "Information Required"
        PType.SetFocus
        Exit Sub
    End If
    If (txtCompany.Text = "" Or txtCompany.Text = " ") Then txtCompany.Text = "-"
    If (txtPS.Text = "" Or txtPS.Text = " ") Then txtPS.Text = "-"
    If (txtPricePU.Text = "" Or txtPricePU.Text = " ") Then txtPricePU.Text = "0"
    If (txtDescription.Text = "" Or txtDescription.Text = " ") Then txtDescription.Text = "-"
    If (txtStock.Text = "" Or txtStock.Text = " ") Then txtStock.Text = "0"
    If (txtROL.Text = "" Or txtROL.Text = " ") Then txtROL.Text = "10"
    If (txtR.Text = "") Then txtR.Text = "-"
    
    'Updating Database
    If DupCheck("SELECT * FROM Stock WHERE Product_ID='" & txtPID.Text & "' AND Product='" & txtProduct.Text & "' AND Product_Type='" & PType.Text & "' AND Product_Size='" & txtPS.Text & "' AND Company='" & txtCompany.Text & "'") = True Then
        MsgBox "Product Already Exists in Stock!!! ", , "General Error"
    Else
        sql = "INSERT INTO Stock VALUES('" & txtPID & "','" & txtDate & "','" & UCase(txtProduct) & "','" & UCase(PType) & "','" & txtPS & "','" & txtCompany & "'," & txtStock & ",'" & txtDescription & "'," & txtPricePU & "," & txtROL & ",'" & txtR & "')"
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
    txtProduct.SetFocus
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Stock SET Date='" & txtDate.Text & "',Product='" & txtProduct.Text & "',Product_Type='" & PType.Text & "',Product_Size='" & txtPS.Text & "',Company='" & txtCompany.Text & "',Stock_In_Hand=" & txtStock.Text & ",Description='" & txtDescription.Text & "',Price_Per_Unit=" & txtPricePU.Text & ",ReOrder_Level=" & txtROL.Text & ",Remarks='" & txtR.Text & "' Where Product_ID='" & txtPID.Text & "'"
    conn.Execute sql
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_ID")
    Set DataGrid1.DataSource = RsProductGrid
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
    DataGrid1.Row = Rx

    Normalize
    cmdRDB_Click
    
End Sub

Private Sub cmdCancel_Click()
    ClearFields
    Normalize
End Sub

Private Sub cmdDelete_Click()
    Dim sqlR, sqlIn, sqlS As String
        
    If MsgBox("This will DELETE Complete Data of the current Product from Database[Receivings, Invoices & Stock]. ARE YOU SURE?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        'Set rsTemp = Nothing
        Exit Sub
    End If
    
    'Deleting Receivings info
    sqlR = "DELETE FROM Receivings WHERE Product_ID='" & txtPID.Text & "'"
'    MsgBox "SQL IS " & sqlR
    conn.Execute sqlR
    
    'Deleting Invoice info
    sqlIn = "DELETE FROM Invoice WHERE Product_ID='" & txtPID.Text & "'"
'    MsgBox "SQL IS " & sqlSA
    conn.Execute sqlIn
    
    'Deleting Purchase_Order info
    sqlS = "DELETE FROM Stock WHERE Product_ID='" & txtPID.Text & "'"
'    MsgBox "SQL IS " & sqlPO
    conn.Execute sqlS
       
    Rx = Rx - 1
    Normalize
    cmdRDB_Click
    Set DataGrid1.DataSource = RsProductGrid
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
    If (Rx <> 0) Then DataGrid1.Row = Rx
    ClearFields
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    SQLString = "SELECT * FROM Stock ORDER BY Product_ID"
    Rx = 0
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_ID")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_ID")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_ID")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_ID")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_ID")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
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
    
    SQLString = "SELECT * FROM Stock WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RsProductGrid = New ADODB.Recordset
    RsProductGrid.CursorLocation = adUseClient
    RsProductGrid.CursorType = adOpenStatic
    RsProductGrid.LockType = adLockReadOnly
    RsProductGrid.Open SQLString, conn
    Set DataGrid1.DataSource = RsProductGrid
      
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        
        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!Product_ID) Then
        ClearFields
    Else
       
    txtPID.Text = rs!Product_ID
    txtDate.Text = Format(rs!Dated, "YYYY-MM-DD")
    txtProduct.Text = rs!Product
    PType.Text = rs!Product_Type
    txtPS.Text = rs!Product_Size
    txtCompany.Text = rs!Company
    txtPricePU.Text = rs!Price_Per_Unit
    txtDescription.Text = rs!Description
    txtStock.Text = rs!Stock_In_Hand
    txtROL.Text = rs!ReOrder_Level
    txtR.Text = rs!Remarks
    
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub ClearFields()
    txtPID.Text = ""
    txtDate.Text = ""
    txtProduct.Text = ""
    PType.Text = ""
    txtPS.Text = ""
    txtCompany.Text = ""
    txtPricePU.Text = ""
    txtDescription.Text = ""
    txtStock.Text = ""
    txtROL.Text = ""
    txtR.Text = ""
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtDate.Enabled = TextFieldLock
    txtProduct.Enabled = TextFieldLock
    PType.Enabled = TextFieldLock
    txtPS.Enabled = TextFieldLock
    txtCompany.Enabled = TextFieldLock
    txtPricePU.Enabled = TextFieldLock
    txtDescription.Enabled = TextFieldLock
    txtStock.Enabled = TextFieldLock
    txtROL.Enabled = TextFieldLock
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

    txtSearch.Enabled = True
    txtSearch.Text = ""
    ST.Enabled = True
    cmdRDB_Click
End Sub

Public Sub EnterNewProduct()
    
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
    txtProduct.SetFocus
    
End Sub
Private Sub GenerateID()
    txtPID.Text = "P" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtPID.Text = rs!Product_ID Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub ST_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ST_LostFocus()
    If ST.Text = "Date" Then
        txtSearch.ToolTipText = "Date Format YYYY-MM-DD"
        txtSearch.Text = "2006-03-03"
    Else
        Exit Sub
    End If
End Sub

Private Sub GetComboData()
    Adodc.ConnectionString = conn
    Adodc.CursorLocation = adUseClient
    Adodc.CursorType = adOpenDynamic
    Adodc.RecordSource = "SELECT Product_Type FROM Stock ORDER BY Product_Type"
    Set DataGrid.DataSource = Adodc
    
    If Adodc.Recordset.BOF Then
        Exit Sub
    Else

    'For Item1 and Item Combo
        Dim X As Integer
        For X = 0 To (Adodc.Recordset.RecordCount - 1)
            PType.AddItem Adodc.Recordset.Fields(0)
            Adodc.Recordset.MoveNext
        Next X
    End If
    
End Sub

Public Function RemoveComboDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = PType.ListCount + 1
    For X = 1 To PType.ListCount
        Y = Y - 1
        If PType.List(Y) = PType.List(Y - 1) Then
            PType.RemoveItem (Y)
        End If
    Next
End Function

Private Sub PType_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(PType)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(PType)
   End Select
End Sub
Private Sub PType_LostFocus()
    'PType.Text = UCase(PType.Text)
End Sub
Private Sub PType_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(PType.SelText) <> 0 And PType.SelStart > 0 Then
         PType.SelStart = PType.SelStart - 1
         PType.SelLength = CB_MAXLENGTH
   End If
End Sub


Private Sub txtDescription_GotFocus()
    SendKeys "{Home}+{End}"
End Sub


Private Sub txtPID_Change()
Call DrawBarcode(txtPID, Picture1)
End Sub

Private Sub txtPricePU_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtProduct_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPS_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtROL_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtStock_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub CheckROL()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT COUNT(*) as No FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    If Val(rs!No) > 0 Then
        If MsgBox("Some products needs to be ReOrdered!, would you like to have a look?", vbYesNo + vbDefaultButton2, "Stock") = vbYes Then
        'SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
        'MsgBox rs!No & " Product(s) needs to be ReOrdered!", vbInformation, "Stock"
        isReOrder = True
        End If
    Else
        isReOrder = False
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub CheckMinusStock()
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
    If Val(rs!No) > 0 Then
        If MsgBox("Some product's quantities are in minus in stock, which needs to be urgently ReOrdered!, would you like to have a look?", vbYesNo + vbDefaultButton2, "Stock") = vbYes Then
        'SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<0;"
        'MsgBox rs!No & " Product(s) in Stock needs to be ReOrdered URGENTLY!", vbCritical, "Stock"
        isStockMinus = True
        End If
    Else
        isStockMinus = False
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
End Sub
