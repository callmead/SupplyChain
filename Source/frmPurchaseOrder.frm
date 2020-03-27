VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: PURCHASE ORDER :."
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "frmPurchaseOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmPurchaseOrder.frx":0BC2
   ScaleHeight     =   9645
   ScaleWidth      =   10980
   Begin VB.CommandButton cmdAd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      Picture         =   "frmPurchaseOrder.frx":C1006
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add to Cart"
      Top             =   3600
      Width           =   375
   End
   Begin VB.ComboBox Prod 
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
      ItemData        =   "frmPurchaseOrder.frx":C1CD0
      Left            =   360
      List            =   "frmPurchaseOrder.frx":C1CD2
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   "Prod"
      Top             =   3600
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
      Left            =   7320
      TabIndex        =   12
      Text            =   "txtDescript"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtQty 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
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
      Left            =   6240
      TabIndex        =   11
      Text            =   "txtQty"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtPT 
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
      Left            =   2760
      TabIndex        =   9
      Text            =   "txtPT"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtSize 
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
      Left            =   4800
      TabIndex        =   10
      Text            =   "txtSize"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtDD 
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
      TabIndex        =   5
      Text            =   "txtDD"
      Top             =   720
      Width           =   2295
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
      TabIndex        =   26
      Top             =   960
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
      TabIndex        =   25
      Top             =   2040
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
      TabIndex        =   24
      Top             =   2400
      Width           =   1455
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
      TabIndex        =   23
      Top             =   2760
      Width           =   1455
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
      Height          =   795
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmPurchaseOrder.frx":C1CD4
      Top             =   1560
      Width           =   6615
   End
   Begin VB.TextBox txtTID 
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
      Text            =   "txtTID"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtPO 
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
      Text            =   "txtPO"
      Top             =   240
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
      Height          =   360
      Left            =   6600
      TabIndex        =   2
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   240
      Width           =   2295
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
      TabIndex        =   22
      Top             =   1320
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
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
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
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtSID 
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
      Text            =   "txtSID"
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox ST 
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
      ItemData        =   "frmPurchaseOrder.frx":C1CD9
      Left            =   4560
      List            =   "frmPurchaseOrder.frx":C1CE9
      Sorted          =   -1  'True
      TabIndex        =   18
      Text            =   "PO_No"
      Top             =   6600
      Width           =   1935
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
      Left            =   6600
      TabIndex        =   17
      Top             =   6600
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
      Left            =   360
      TabIndex        =   16
      Text            =   "txtSearch"
      Top             =   6600
      Width           =   4095
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
      Left            =   8640
      TabIndex        =   15
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectSup 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid PrdGrid 
      Height          =   2085
      Left            =   360
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3678
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16744576
      ForeColor       =   16777215
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorSel    =   8421631
      BackColorBkg    =   9081241
      GridColor       =   4210752
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      MousePointer    =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPurchaseOrder.frx":C1D16
   End
   Begin MSComCtl2.UpDown ScrollBar 
      Height          =   855
      Left            =   10320
      TabIndex        =   41
      Top             =   4080
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   1508
      _Version        =   393216
      Enabled         =   -1  'True
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
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Top             =   600
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      TabIndex        =   28
      Top             =   240
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   8880
      Top             =   8640
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
      Left            =   8520
      TabIndex        =   42
      Top             =   8040
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   360
      TabIndex        =   27
      Top             =   7080
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
   Begin VB.Label lblDescription 
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
      Left            =   7680
      TabIndex        =   39
      Top             =   3240
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
      Left            =   6240
      TabIndex        =   38
      Top             =   3240
      Width           =   855
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
      Left            =   2760
      TabIndex        =   37
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
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
      Left            =   5160
      TabIndex        =   36
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Product "
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
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblDD 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
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
      TabIndex        =   34
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   10560
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   9360
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
      TabIndex        =   33
      Top             =   240
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
      TabIndex        =   32
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblSupplierID 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
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
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblTID 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction #"
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
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblPo 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order #"
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
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock, AddingData As Boolean
Private sql2, sql3 As String
Dim nLastKeyAscii As Integer

Private Sub cmdAd_Click()
    If (Prod.Text = "" Or Prod.Text = " ") Then
        MsgBox "Please select a Product !!!", vbOKOnly, "Information Required"
        Prod.SetFocus
        Exit Sub
    End If
    If (txtPT.Text = "" Or txtPT.Text = " ") Then
        MsgBox "Please provide a Product Type !!!", vbOKOnly, "Information Required"
        txtPT.SetFocus
        Exit Sub
    End If
    If (txtQty.Text = "" Or txtQty.Text = " ") Then
        MsgBox "Please provide Quantity for selected Product !!!", vbOKOnly, "Information Required"
        txtDD.SetFocus
        Exit Sub
    End If
    If (txtSize.Text = "" Or txtSize.Text = " ") Then
        MsgBox "Please provide Size for selected Product !!!", vbOKOnly, "Information Required"
        txtSize.SetFocus
        Exit Sub
    End If
    If (txtDescription.Text = "" Or txtSize.Text = " ") Then txtDescription.Text = "-"
    
    AddingData = True
    GenerateTID
    Prod.SetFocus
    
End Sub

Private Sub Form_Load()
  
    Connect
    
    ShowPOData ("SELECT * FROM Purchase_Order ORDER BY PO_No")
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
    
    ClearFields
    
    GetDate
    GridSet
    
    Normalize
    txtSearch.Text = ""
    
    txtSearch.Enabled = True
    ST.Enabled = True
    
    'For Int TextBoxes
    Dim tmp1 As Long
    tmp1 = SetWindowLong(txtQty.hwnd, GWL_STYLE, GetWindowLong(txtQty.hwnd, GWL_STYLE) Or ES_NUMBER)

    GetComboData
    RemoveComboDuplicates
    AddingData = False
End Sub

Private Sub Prod_Change()
    Select Case nLastKeyAscii
        Case vbKeyBack
            Call Combo_Lookup(Prod)
        Case vbKeyDelete
            Case Else
        Call Combo_Lookup(Prod)
    End Select
   
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
    PrdGrid.TextMatrix(iR, 1) = Prod.Text
End Sub

Private Sub Prod_KeyPress(KeyAscii As Integer)
    If AddingData = True Then
        AddingData = False
        iR = iR + 1

        PrdGrid.Rows = PrdGrid.Rows + 1
        iR = PrdGrid.Rows - 1
    End If
End Sub

Private Sub Prod_LostFocus()
    If AddingData = True Then
        AddingData = False
        iR = iR + 1

        PrdGrid.Rows = PrdGrid.Rows + 1
        iR = PrdGrid.Rows - 1
    End If
    
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 1) = Prod.Text
End Sub

Private Sub Prod_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(Prod.SelText) <> 0 And Prod.SelStart > 0 Then
         Prod.SelStart = Prod.SelStart - 1
         Prod.SelLength = CB_MAXLENGTH
   End If
End Sub

Private Sub txtDD_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtTID_Change()
    On Error Resume Next
    'iR = PrdGrid.Rows - 1
    'PrdGrid.TextMatrix(iR, 0) = txtTID.Text
End Sub
Private Sub txtTID_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

'Private Sub txtProduct_Change()
'    On Error Resume Next
'    PrdGrid.TextMatrix(iR, 1) = txtProduct.Text
'End Sub
'Private Sub txtProduct_GotFocus()
'    SendKeys "{Home}+{End}"
'End Sub

Private Sub txtPT_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 2) = txtPT.Text
End Sub
Private Sub txtPT_LostFocus()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
    PrdGrid.TextMatrix(iR, 1) = Prod.Text
    PrdGrid.TextMatrix(iR, 2) = txtPT.Text
End Sub
Private Sub txtPT_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtQty_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 4) = txtQty.Text
End Sub
Private Sub txtQty_LostFocus()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 4) = txtQty.Text
End Sub
Private Sub txtQty_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtSize_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 3) = txtSize.Text
End Sub
Private Sub txtSize_LostFocus()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 3) = txtSize.Text
End Sub
Private Sub txtSize_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDescription_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 5) = txtDescription.Text
End Sub
Private Sub txtDescription_LostFocus()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 5) = txtDescription.Text
End Sub
Private Sub txtDescription_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub PrdGrid_Click()
    On Error Resume Next
    iR = PrdGrid.Row
    iC = PrdGrid.Col
    
    If iC = 0 Then
        txtTID.Text = PrdGrid.TextMatrix(iR, iC)
        'txtTID.SetFocus
        'SendKeys "{Home}+{End}"
    End If
    If iC = 1 Then
        'txtProduct.Text = PrdGrid.TextMatrix(iR, iC)
        Prod.Text = PrdGrid.TextMatrix(iR, iC)
        Prod.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 2 Then
        txtPT.Text = PrdGrid.TextMatrix(iR, iC)
        txtPT.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 3 Then
        txtSize.Text = PrdGrid.TextMatrix(iR, iC)
        txtSize.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 4 Then
        txtQty.Text = PrdGrid.TextMatrix(iR, iC)
        txtQty.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 5 Then
        txtDescription.Text = PrdGrid.TextMatrix(iR, iC)
        txtDescription.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub ScrollBar_DownClick()
    PrdGrid.Rows = PrdGrid.Rows + 1
End Sub

Private Sub ScrollBar_UpClick()
    If PrdGrid.Rows > 2 Then PrdGrid.Rows = PrdGrid.Rows - 1
    If PrdGrid.Rows = 2 Then PrdGrid.Rows = 1: PrdGrid.Rows = 2
End Sub

Private Sub cmdSelectSup_Click()
    ParentForm = "frmPurchaseOrder"
    GridSQLString = "Select Supplier_ID,Name,Company from Supplier ORDER BY Company"
    SelectedField = 0
    frmDataSelect.Show vbModal
    If txtSID.Text = "" Then
        frmSupplier.Show
        frmSupplier.EnterNewSupplier
    End If
End Sub

Private Sub cmdNew_Click()
    PrdGrid.Rows = 1
    PrdGrid.Rows = 2
    
    iR = PrdGrid.Rows - 1
    
    ClearFields
    GeneratePO
    txtDate.Text = DateToday
    txtDD.Text = DateToday
    GenerateTID
    SetFields (True)
    SetButtons (False)
    cmdSelectSup.Enabled = True
    cmdAdd.Enabled = True
    cmdSelectSup.SetFocus
    
    cmdEdit.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    cmdAd.Enabled = True
    txtSearch.Enabled = False
    ST.Enabled = False
End Sub
Private Sub cmdAdd_Click()
    
    'Checking Fields for Records
    If (txtSID.Text = "" Or txtSID.Text = " ") Then
        MsgBox "Please select a Supplier !!!", vbOKOnly, "Information Required"
        txtSID.SetFocus
        Exit Sub
    End If
    If (txtDD.Text = "" Or txtDD.Text = " ") Then
        MsgBox "Please provide a Delivery Date !!!", vbOKOnly, "Information Required"
        txtDD.SetFocus
        Exit Sub
    End If
    If (Prod.Text = "" Or Prod.Text = " ") Then
        MsgBox "Please select a Product !!!", vbOKOnly, "Information Required"
        Prod.SetFocus
        Exit Sub
    End If
    If (txtPT.Text = "" Or txtPT.Text = " ") Then
        MsgBox "Please provide a Product Type !!!", vbOKOnly, "Information Required"
        txtPT.SetFocus
        Exit Sub
    End If
    If (txtQty.Text = "" Or txtQty.Text = " ") Then
        MsgBox "Please provide Quantity for selected Product !!!", vbOKOnly, "Information Required"
        txtDD.SetFocus
        Exit Sub
    End If
    If (txtSize.Text = "" Or txtSize.Text = " ") Then
        MsgBox "Please provide Size for selected Product !!!", vbOKOnly, "Information Required"
        txtSize.SetFocus
        Exit Sub
    End If
    
    If (txtR.Text = "") Then txtR.Text = "-"
    
    iR = PrdGrid.Rows - 1
    
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
    PrdGrid.TextMatrix(iR, 1) = Prod.Text
    PrdGrid.TextMatrix(iR, 2) = txtPT.Text
    PrdGrid.TextMatrix(iR, 3) = txtSize.Text
    PrdGrid.TextMatrix(iR, 4) = txtQty.Text
    PrdGrid.TextMatrix(iR, 5) = txtDescription.Text
    
    'Updating Database
    If DupCheck("SELECT * from Purchase_Order WHERE PO_No='" & txtPO.Text & "'") = True Then
        MsgBox "Purchase Order Already Exists !!! ", , "General Error"
    Else
        sql = "INSERT INTO Purchase_Order values('" & txtPO & "','" & txtDate & "','" & txtSID & "','" & txtDD & "','" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
            
        'Save PO_Details
        If Len(txtTID.Text) > 0 And Len(PrdGrid.TextMatrix(1, 1)) > 0 Then
            
            rn = 1
        
            For rn = 1 To PrdGrid.Rows - 1
            If PrdGrid.TextMatrix(rn, 0) <> "" Then
        
                sql = "INSERT INTO PO_Details Values("
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 0)) & "',"
                sql = sql & "'" & txtPO.Text & "',"
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 1)) & "',"
                sql = sql & "'" & UCase((PrdGrid.TextMatrix(rn, 2))) & "',"
                sql = sql & "'" & UCase((PrdGrid.TextMatrix(rn, 3))) & "',"
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 4))) & ","
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 5)) & "');"
                
                conn.Execute sql
        
            End If
            Next
            
            MsgBox "Data Saved Successfully", vbInformation, "Admin"
            SetFields (False)
            ClearFields
            txtSearch.Enabled = True
            ST.Enabled = True
            cmdNew.Visible = True
            cmdSave.Enabled = False
                        
            Normalize
            cmdRDB_Click
            cmdNew.SetFocus
        Else
            MsgBox "Data Not Available", vbCritical, "Admin"
        End If
        Exit Sub
    End If

End Sub

Private Sub cmdEdit_Click()
    'SetFields (True)
    'cmdSelectSup.Enabled = True
    'cmdSelectSup.SetFocus
    
    txtDD.Enabled = True
    txtR.Enabled = True
    
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdNew.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()

    sql = "UPDATE Purchase_Order SET Delivery_Date='" & txtDD.Text & ",Remarks='" & txtR.Text & "' Where Supplier_ID='" & txtSID.Text & "'"
    conn.Execute sql
    ShowPOData (SQLString)
    Set DataGrid1.DataSource = RsSuppGrid
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
    DataGrid1.Row = Rx

    Normalize
    cmdRDB_Click

End Sub
Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()

    Dim rsTemp, rsTemp2, rsTemp3 As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp2 = New ADODB.Recordset
    Set rsTemp3 = New ADODB.Recordset
    rsTemp.CursorType = adOpenDynamic
    rsTemp2.CursorType = adOpenDynamic
    rsTemp3.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    rsTemp2.LockType = adLockOptimistic
    rsTemp3.LockType = adLockOptimistic
    rsTemp.CursorLocation = adUseServer
    rsTemp2.CursorLocation = adUseServer
    rsTemp3.CursorLocation = adUseServer
    
    sql2 = "DELETE FROM PO_Details WHERE PO_No='" & txtPO.Text & "'"
    sql3 = "DELETE FROM Supplier_Account WHERE PO_No='" & txtPO.Text & "'"
    sql = "DELETE FROM Purchase_Order Where PO_No='" & txtPO.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record? This will all Details of the selected Purchase Order from Database!", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Set rsTemp = Nothing
        Set rsTemp2 = Nothing
        Exit Sub
    End If
    rsTemp3.Open sql3, conn, adOpenKeyset
    rsTemp2.Open sql2, conn, adOpenKeyset
    rsTemp.Open sql, conn, adOpenKeyset
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing
    Set rsTemp = Nothing
    
    Rx = Rx - 1
    Normalize
    cmdRDB_Click
    Set DataGrid1.DataSource = RsPOGrid
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
    If (Rx <> 0) Then DataGrid1.Row = Rx
    ClearFields
    

End Sub

Private Sub GridSet()
    With PrdGrid
    .Cols = 6
    .Rows = 2
    .ColWidth(0) = 1900
    .ColWidth(1) = 2500
    .ColWidth(2) = 1700
    .ColWidth(3) = 800
    .ColWidth(4) = 950
    .ColWidth(5) = 1850
    
    .TextMatrix(0, 0) = " Transaction"
    .TextMatrix(0, 1) = " Product"
    .TextMatrix(0, 2) = " Product Type"
    .TextMatrix(0, 3) = " Size"
    .TextMatrix(0, 4) = " Quantity"
    .TextMatrix(0, 5) = " Description"
    End With
End Sub

Private Sub GeneratePO()
    txtPO.Text = "PO" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub
Private Sub GenerateTID()
    txtTID.Text = "T" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtDD.Enabled = TextFieldLock
    Prod.Enabled = TextFieldLock
    txtR.Enabled = TextFieldLock
    'txtProduct.Enabled = TextFieldLock
    txtPT.Enabled = TextFieldLock
    txtSize.Enabled = TextFieldLock
    txtQty.Enabled = TextFieldLock
    txtDescription.Enabled = TextFieldLock
End Sub

Private Sub SetButtons(ButtonLock As Boolean)
    cmdDelete.Enabled = ButtonLock
    cmdRDB.Enabled = ButtonLock
    cmdMF.Enabled = ButtonLock
    cmdML.Enabled = ButtonLock
    cmdN.Enabled = ButtonLock
    cmdP.Enabled = ButtonLock
    cmdSearch.Enabled = ButtonLock
    cmdClose.Enabled = ButtonLock
End Sub

Private Sub ClearFields()
    txtPO.Text = ""
    txtDate.Text = ""
    Prod.Text = ""
    txtSID.Text = ""
    txtDD.Text = ""
    txtR.Text = ""
    txtTID.Text = ""
    'txtProduct.Text = ""
    txtPT.Text = ""
    txtSize.Text = ""
    txtQty.Text = ""
    txtDescription.Text = ""
    txtSearch.Text = ""
End Sub

Private Sub ST_LostFocus()
    If ST.Text = "Date" Then
        txtSearch.ToolTipText = "Date Format YYYY-MM-DD"
        txtSearch.Text = "2006-03-03"
    Else
        Exit Sub
    End If
End Sub

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub Normalize()
    SetFields (False)
    SetButtons (True)
    cmdNew.Visible = True
    cmdEdit.Visible = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdSelectSup.Enabled = False
    cmdAd.Enabled = False
    cmdAdd.Enabled = False
    
    txtSearch.Enabled = True
    ST.Enabled = True
    
    PrdGrid.Rows = 1
    PrdGrid.Rows = 2
    
    iR = PrdGrid.Rows - 1
    cmdRDB_Click
    AddingData = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    Rx = 0
    ShowPOData ("SELECT * FROM Purchase_Order ORDER BY PO_No")
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowPOData ("SELECT * FROM Purchase_Order ORDER BY PO_No")
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowPOData ("SELECT * FROM Purchase_Order ORDER BY PO_No")
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowPOData ("SELECT * FROM Purchase_Order ORDER BY PO_No")
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowPOData ("SELECT * FROM Purchase_Order ORDER BY PO_No")
    ShowPOGrid ("SELECT * FROM PO_Details WHERE PO_No = '" & txtPO.Text & "' ORDER BY Product")
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
    
    SQLString = "SELECT * FROM Purchase_Order WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    sql2 = "SELECT * FROM PO_Details WHERE PO_No='" & txtPO & "'"
    
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
          
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        
        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!PO_No) Then
        ClearFields
    Else
       
    txtPO.Text = rs!PO_No
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    txtSID.Text = rs!Supplier_ID
    txtDD = Format(rs!Delivery_Date, "YYYY-MM-DD")
    txtR.Text = rs!Remarks
    
    Set RsPOGrid = New ADODB.Recordset
    RsPOGrid.CursorLocation = adUseClient
    RsPOGrid.CursorType = adOpenStatic
    RsPOGrid.LockType = adLockReadOnly
    RsPOGrid.Open sql2, conn
    Set DataGrid1.DataSource = RsPOGrid
    
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    'SQL = "SELECT * from Supplier WHERE Supplier_ID='" & chkID & "'"
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtPO.Text = rs!PO_No Then
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

Private Sub GetComboData()
    Adodc.ConnectionString = conn
    Adodc.CursorLocation = adUseClient
    Adodc.CursorType = adOpenDynamic
    Adodc.RecordSource = "SELECT Product FROM Stock ORDER BY Product"
    Set DataGrid.DataSource = Adodc
    
    If Adodc.Recordset.BOF Then
        Exit Sub
    Else

    'For Item1 and Item Combo
        Dim X As Integer
        For X = 0 To (Adodc.Recordset.RecordCount - 1)
            Prod.AddItem Adodc.Recordset.Fields(0)
            Adodc.Recordset.MoveNext
        Next X
    End If
    
End Sub

Public Function RemoveComboDuplicates()
    Dim Y As Integer
    Dim X As Integer
    Y = Prod.ListCount + 1
    For X = 1 To Prod.ListCount
        Y = Y - 1
        If Prod.List(Y) = Prod.List(Y - 1) Then
            Prod.RemoveItem (Y)
        End If
    Next
End Function
