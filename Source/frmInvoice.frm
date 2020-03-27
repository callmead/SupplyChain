VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: INVOICE :."
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   Icon            =   "frmInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmInvoice.frx":0BC2
   ScaleHeight     =   9900
   ScaleWidth      =   10935
   Begin VB.CommandButton cmdAd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      Picture         =   "frmInvoice.frx":240C04
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add to Cart"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtAD 
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
      Left            =   9000
      TabIndex        =   19
      Text            =   "txtAD"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox txtAP 
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
      Left            =   7320
      TabIndex        =   18
      Text            =   "txtAP"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.ComboBox PM 
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
      ItemData        =   "frmInvoice.frx":2418CE
      Left            =   5640
      List            =   "frmInvoice.frx":2418DB
      Sorted          =   -1  'True
      TabIndex        =   17
      Text            =   "PM"
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox txtProduct 
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
      Text            =   "txtProduct"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtPID 
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
      Left            =   360
      TabIndex        =   7
      Text            =   "txtPID"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelProd 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtGT 
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
      Left            =   3960
      TabIndex        =   16
      Text            =   "txtGT"
      Top             =   6120
      Width           =   1575
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
      Left            =   9120
      TabIndex        =   0
      Top             =   240
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
      Left            =   9240
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectCus 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   720
      Width           =   375
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
      TabIndex        =   31
      Top             =   6840
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
      TabIndex        =   28
      Text            =   "txtSearch"
      Top             =   6840
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
      Left            =   6600
      TabIndex        =   30
      Top             =   6840
      Width           =   1935
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
      ItemData        =   "frmInvoice.frx":2418FA
      Left            =   4560
      List            =   "frmInvoice.frx":241907
      Sorted          =   -1  'True
      TabIndex        =   29
      Text            =   "Invoice_No"
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox txtCID 
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
      Text            =   "txtCID"
      Top             =   720
      Width           =   1815
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
      Left            =   9120
      TabIndex        =   23
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
      Left            =   9120
      TabIndex        =   27
      Top             =   2760
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
      Left            =   9120
      TabIndex        =   22
      Top             =   960
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
      TabIndex        =   2
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtInv 
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
      Text            =   "txtInv"
      Top             =   240
      Width           =   2295
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
      TabIndex        =   3
      Text            =   "txtTID"
      Top             =   720
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
      Height          =   795
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmInvoice.frx":24192A
      Top             =   1440
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
      Left            =   9120
      TabIndex        =   26
      Top             =   2400
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
      Left            =   9120
      TabIndex        =   25
      Top             =   2040
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
      Left            =   9120
      TabIndex        =   24
      Top             =   1680
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
      Left            =   9120
      TabIndex        =   21
      Top             =   600
      Width           =   1455
   End
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
      Left            =   360
      TabIndex        =   14
      Text            =   "txtSalesman"
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox txtQty 
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
      Left            =   4320
      TabIndex        =   10
      Text            =   "txtQty"
      Top             =   3120
      Width           =   1095
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
      Left            =   5520
      TabIndex        =   11
      Text            =   "txtPrice"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtNT 
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
      Left            =   6960
      TabIndex        =   12
      Text            =   "txtNT"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtDiscount 
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
      TabIndex        =   15
      Text            =   "txtDiscount"
      Top             =   6120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid PrdGrid 
      Height          =   2085
      Left            =   360
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3600
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
      MouseIcon       =   "frmInvoice.frx":24192F
   End
   Begin MSComCtl2.UpDown ScrollBar 
      Height          =   855
      Left            =   10320
      TabIndex        =   33
      Top             =   3600
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   1508
      _Version        =   393216
      Enabled         =   -1  'True
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
      Left            =   9120
      TabIndex        =   20
      Top             =   240
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
      Left            =   9120
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   9000
      Top             =   9120
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
      TabIndex        =   53
      Top             =   8280
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
      TabIndex        =   36
      Top             =   7320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      DefColWidth     =   107
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
   Begin VB.Label lblAD 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due"
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
      Left            =   9000
      TabIndex        =   52
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblAP 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid"
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
      Left            =   7320
      TabIndex        =   51
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblPM 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode"
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
      Left            =   5640
      TabIndex        =   50
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   2280
      TabIndex        =   49
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblProd 
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
      TabIndex        =   48
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblGT 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
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
      Left            =   3960
      TabIndex        =   47
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label lblInvNo 
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
      TabIndex        =   46
      Top             =   240
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
      TabIndex        =   45
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblCID 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      TabIndex        =   44
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
      TabIndex        =   43
      Top             =   1440
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
      TabIndex        =   42
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   9600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   10560
      Y1              =   6600
      Y2              =   6600
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
      Left            =   360
      TabIndex        =   41
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblNT 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total"
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
      Left            =   6960
      TabIndex        =   40
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   8880
      Y1              =   1200
      Y2              =   1200
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
      Left            =   5520
      TabIndex        =   39
      Top             =   2760
      Width           =   1095
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
      Left            =   4320
      TabIndex        =   38
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblDisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   2280
      TabIndex        =   37
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock, MinusQuantity, ChangeQuantity, UpdateAmounts, MinusAmounts, ChangeAmounts, AddingData As Boolean
Private sql2, sql3, Query, SalesmanName, Prod_Name, Prod_ID, Cuss_ID, S_Product_ID As String
Dim GrandTotal, ExistingQuantity, ExistingAmount, ExistingDue, CurrentAmount, CurrentDue, NewAmount, NewDue, NewQuantity, CurrentQuantity, ppu As Integer

Private Sub Form_Load()
    
    Connect
    
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
    
    ClearFields
    
    GetDate
    GridSet
    
    Normalize
    txtSearch.Text = ""
    
    MinusQuantity = False
    MinusAmounts = False
    UpdateAmounts = False
    ChangeAmounts = False
    ChangeQuantity = False
    
    txtSearch.Enabled = True
    ST.Enabled = True
    
     'For Int TextBoxes
    Dim tmp1, tmp2, tmp3, tmp4, tmp5 As Long
    tmp1 = SetWindowLong(txtQty.hwnd, GWL_STYLE, GetWindowLong(txtQty.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp3 = SetWindowLong(txtPrice.hwnd, GWL_STYLE, GetWindowLong(txtPrice.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp4 = SetWindowLong(txtAP.hwnd, GWL_STYLE, GetWindowLong(txtAP.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp5 = SetWindowLong(txtDiscount.hwnd, GWL_STYLE, GetWindowLong(txtDiscount.hwnd, GWL_STYLE) Or ES_NUMBER)
    
    AddingData = False
End Sub

Private Sub cmdNew_Click()
    PrdGrid.Rows = 1
    PrdGrid.Rows = 2
    
    iR = PrdGrid.Rows - 1
    
    ClearFields
    txtDate.Text = DateToday
    SetFields (True)
    SetButtons (False)
    cmdSelectCus.Enabled = True
    cmdSelectCus.SetFocus
    cmdSelProd.Enabled = True
    
    cmdEdit.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    txtSearch.Enabled = False
    ST.Enabled = False
    
    txtQty.Text = 0
    txtPrice.Text = 0
    txtDiscount.Text = "0"
    txtGT.Text = "0"
    txtAP.Text = "0"
    txtAD.Text = "0"
    GrandTotal = 0
    GenerateTID
    txtSalesman.Text = UserName
    cmdAd.Enabled = True
    GenerateInvoiceNo
    
End Sub
Private Sub cmdAdd_Click()
    
    'Checking Fields for Records
    If (txtCID.Text = "" Or txtCID.Text = " ") Then
        MsgBox "Please select a Customer !!!", vbOKOnly, "Information Required"
        txtCID.SetFocus
        Exit Sub
    End If
    If (txtPID.Text = "" Or txtPID.Text = " ") Then
        MsgBox "Please select a Product !!!", vbOKOnly, "Information Required"
        txtPID.SetFocus
        Exit Sub
    End If
    If (PM.Text = "" Or PM.Text = " ") Then
        MsgBox "Please select Payment Mode !!!", vbOKOnly, "Information Required"
        PM.SetFocus
        Exit Sub
    End If
    If (txtAP.Text = "" Or txtAP.Text = " ") Then
        MsgBox "Please provide Amount Paid by the Customer !!!", vbOKOnly, "Information Required"
        txtAP.SetFocus
        Exit Sub
    End If
    
    
    If (txtR.Text = "") Then txtR.Text = "-"
    If (txtDiscount.Text = "") Then txtDiscount.Text = "0"
    
    iR = PrdGrid.Rows - 1
    
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
    PrdGrid.TextMatrix(iR, 1) = txtPID.Text
    PrdGrid.TextMatrix(iR, 2) = txtProduct.Text
    PrdGrid.TextMatrix(iR, 3) = txtQty.Text
    PrdGrid.TextMatrix(iR, 4) = txtPrice.Text
    PrdGrid.TextMatrix(iR, 5) = txtNT.Text
    
    'Updating Database
    If DupCheck("SELECT * from Sales WHERE Invoice_No='" & txtInv.Text & "'") = True Then
        MsgBox "Invoice No Already Exists !!! ", , "General Error"
    Else
    
        'Update Sales Table...
        sql = "INSERT INTO Sales VALUES('" & txtInv & "','" & txtDate & "','" & UserName & "','" & txtCID & "'," & txtGT & "," & txtDiscount & ",'" & PM.Text & "'," & txtAP & "," & txtAD & ",'" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
                
        'Update Invoice Table...
        If Len(txtTID.Text) > 0 And Len(PrdGrid.TextMatrix(1, 1)) > 0 Then
            
            rn = 1
        
            For rn = 1 To PrdGrid.Rows - 1
            If PrdGrid.TextMatrix(rn, 0) <> "" Then
        
                sql = "INSERT INTO Invoice Values("
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 0)) & "',"
                sql = sql & "'" & txtInv.Text & "',"
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 1)) & "',"
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 3))) & ","
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 4))) & ","
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 5))) & ");"
                
                Prod_ID = (PrdGrid.TextMatrix(rn, 1))
                Cus_ID = txtCID.Text
                
                NewQuantity = (Val(PrdGrid.TextMatrix(rn, 3)))
                
                'MsgBox sql
                conn.Execute sql
                
                MinusQuantity = True
                UpdateAmounts = True
                UpdateQuantities
                
            End If
            Next
            
            'Update Customer Table...
            UpdateCustomer
    
            'Update Customer Account...
            sql = "INSERT INTO Customer_Account VALUES('" & txtTID & "','" & txtCID & "','" & txtDate & "','" & txtInv & "'," & txtGT & ",'" & PM.Text & "'," & txtAP & "," & txtAD & ",'Sale Data')"
            'MsgBox sql
            conn.Execute sql
            
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
            MsgBox "Data Not Available", vbInformation, "Admin"
        End If
        Exit Sub
    End If
    
End Sub

Private Sub cmdAd_Click()
    If (txtQty.Text = "" Or txtQty.Text = " ") Then
        MsgBox "Please provide Quantity for selected Product !!!", vbOKOnly, "Information Required"
        txtQty.SetFocus
        Exit Sub
    End If
    If (txtPrice.Text = "" Or txtPrice.Text = " ") Then
        MsgBox "Please provide price for selected Product !!!", vbOKOnly, "Information Required"
        txtPrice.SetFocus
        Exit Sub
    End If
                
    'Quantity Check is not required by the Client, if this feature is disabled then the stock goes to Minus
    'Discuss with client!
    'Confirm Stock quantity
    If CheckQuanty(txtPID.Text) = True Then
        AddingData = True
        GenerateTID
        cmdSelProd.SetFocus

        Exit Sub
    Else
        'MsgBox "Not Enough Quantity in Stock !!! ", , "Admin"
        If MsgBox("The Quantity that you have entered is not available in stock, Do you wish to countinue?", vbYesNo, "Admin") = vbNo Then
            txtQty.SetFocus
            Exit Sub
        Else
            'iR = iR + 1
            
            GrandTotal = GrandTotal + Val(txtNT.Text)
            txtGT.Text = GrandTotal
            CalculateDue
            
'            If PrdGrid.Rows = iR Then
'                txtTID.Text = ""
'                txtPID.Text = ""
'                txtProduct.Text = ""
'                txtQty.Text = ""
'                txtPrice.Text = ""
'                txtNT.Text = ""
'            End If

            AddingData = True
            GenerateTID
            cmdSelProd.SetFocus
            
        End If
        
    End If
End Sub

Private Sub UpdateQuantities()
    
    NewQuantity = 0
    GetExistingStock (Prod_ID)
    
    If MinusQuantity = True Then
        NewQuantity = ExistingQuantity - (Val(PrdGrid.TextMatrix(rn, 3)))
    
    ElseIf (MinusQuantity = False And ChangeQuantity = False) Then
        'NewQuantity = ExistingQuantity + Val(txtQty.Text)
        'Add a function to Delete Sale Data from Invoice Table and Update Quantities in Stock Table
        UpdateStock
    
    ElseIf ChangeQuantity = True Then
        NewQuantity = (ExistingQuantity - CurrentQuantity) - Val(txtQty.Text)
    End If
    
    Query = "UPDATE Stock SET Stock_In_Hand=" & NewQuantity & " WHERE Product_ID='" & Prod_ID & "'"
    'MsgBox Query
    conn.Execute Query
    
    MinusQuantity = False
    ChangeQuantity = False
    
'    rsTmp.Close
'    Set rsTmp = Nothing
    
End Sub

Private Sub UpdateCustomer()
    
    NewAmount = 0
    NewDue = 0
    
    Set rsTmp = New ADODB.Recordset
    Query = "SELECT Total_Bills_Amount,Total_Due FROM Customer WHERE Customer_ID='" & txtCID.Text & "'"

    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenStatic
    rsTmp.LockType = adLockReadOnly
    rsTmp.Open Query, conn
        If rsTmp.EOF = True Then
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Sub
        End If
    xCount = rsTmp.RecordCount
        If Rx > rsTmp.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsTmp.RecordCount - 1
        End If
    rsTmp.Move Rx
    
    ExistingAmount = Val(rsTmp!Total_Bills_Amount)
    ExistingDue = Val(rsTmp!Total_Due)
    
    If UpdateAmounts = True Then
        NewAmount = ExistingAmount + Val(txtGT.Text)
        NewDue = ExistingDue + Val(txtAD.Text)
    
    ElseIf MinusAmounts = True Then
        NewAmount = ExistingAmount - Val(txtGT.Text)
        NewDue = ExistingDue - Val(txtAD.Text)
    
'    ElseIf (MinusAmounts = False And ChangeAmounts = False) Then
'        NewAmount = ExistingAmount + Val(txtGT.Text)
'        NewDue = ExistingDue + Val(txtAD.Text)
    
    ElseIf ChangeAmounts = True Then
        NewAmount = (ExistingAmount - CurrentAmount) + Val(txtGT.Text)
        NewDue = (ExistingDue - CurrentDue) + Val(txtAD.Text)
    End If
    
    Query = "UPDATE Customer SET Total_Bills_Amount=" & NewAmount & ",Total_due=" & NewDue & " WHERE Customer_ID='" & txtCID.Text & "'"
    'MsgBox Query
    conn.Execute Query
    
    MinusAmounts = False
    ChangeAmounts = False
    UpdateAmounts = False
    
    rsTmp.Close
    Set rsTmp = Nothing
    
End Sub

Private Sub UpdateStock()
    Dim newstock As Integer
    Adodc.ConnectionString = conn
    Adodc.CursorLocation = adUseClient
    Adodc.CursorType = adOpenDynamic
    Adodc.RecordSource = "SELECT * FROM Invoice WHERE Invoice_No='" & txtInv.Text & "'"
    Set DataGrid.DataSource = Adodc
    
    If Adodc.Recordset.BOF Then
        Exit Sub
    Else
        Dim X As Integer
        For X = 0 To (Adodc.Recordset.RecordCount - 1)
        
            S_Product_ID = Adodc.Recordset.Fields(2)
            GetExistingStock (S_Product_ID)
            newstock = ExistingQuantity + Val(Adodc.Recordset.Fields(3))
            'MsgBox "Product ID=" & S_Product_ID & " existing quantity=" & ExistingQuantity & " new qty=" & newstock
            sql = "UPDATE Stock SET Stock_In_Hand=" & newstock & " WHERE Product_ID='" & S_Product_ID & "'"
            conn.Execute sql
            
            Adodc.Recordset.MoveNext
        Next X
    End If
    
End Sub

Private Function GetExistingStock(PID As String)
    NewQuantity = 0
    Set rsTmp = New ADODB.Recordset
    Query = "SELECT Stock_In_Hand FROM Stock WHERE Product_ID='" & PID & "'"

    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenStatic
    rsTmp.LockType = adLockReadOnly
    rsTmp.Open Query, conn
        If rsTmp.EOF = True Then
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Function
        End If
    xCount = rsTmp.RecordCount
        If Rx > rsTmp.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsTmp.RecordCount - 1
        End If
    rsTmp.Move Rx
    
    ExistingQuantity = Val(rsTmp!Stock_In_Hand)
End Function
Private Sub cmdEdit_Click()
    'SetFields (True)

    CurrentQuantity = Val(txtQty.Text)
    ChangeQuantity = True
    CurrentAmount = Val(txtGT.Text)
    CurrentDue = Val(txtAD.Text)
    ChangeAmounts = True

    'cmdSelectSup.Enabled = True
    'cmdSelectSup.SetFocus
    SetButtons (False)

    txtDiscount.Enabled = True
    txtAP.Enabled = True

    txtSearch.Enabled = False
    cmdNew.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Sales SET Discount=" & txtDiscount.Text & ",Amount_Paid=" & txtAP.Text & " Where Customer_ID='" & txtCID.Text & "'"
    conn.Execute sql
    
    Prod_ID = txtPID.Text
    UpdateQuantities
    UpdateCustomer
    
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    Set DataGrid1.DataSource = RsInvoiceGrid
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
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
    
    MinusQuantity = False
    MinusAmounts = True
    
    sql2 = "DELETE FROM Invoice WHERE Invoice_No='" & txtInv.Text & "'"
    sql3 = "DELETE FROM Customer_Account WHERE Invoice_No='" & txtInv.Text & "'"
    sql = "DELETE FROM Sales Where Invoice_No='" & txtInv.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record? This will all Details of the selected Invoice from Database!", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Set rsTemp = Nothing
        Set rsTemp2 = Nothing
        Set rsTemp3 = Nothing
        Exit Sub
    End If
    
    UpdateStock
    UpdateCustomer
    
    rsTemp3.Open sql3, conn, adOpenKeyset
    rsTemp2.Open sql2, conn, adOpenKeyset
    rsTemp.Open sql, conn, adOpenKeyset
    Set rsTemp3 = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp = Nothing
        
    Rx = Rx - 1
    Normalize
    cmdRDB_Click
    Set DataGrid1.DataSource = RsPOGrid
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
    If (Rx <> 0) Then DataGrid1.Row = Rx
    ClearFields
    
End Sub

Private Sub txtQty_LostFocus()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 3) = txtQty.Text
End Sub

Private Sub txtTID_Change()
    On Error Resume Next
    'iR = PrdGrid.Rows - 1
    'PrdGrid.TextMatrix(iR, 0) = txtTID.Text
End Sub
Private Sub txtPID_Change()
    On Error Resume Next
    If AddingData = True Then
        AddingData = False
        iR = iR + 1

        PrdGrid.Rows = PrdGrid.Rows + 1
        iR = PrdGrid.Rows - 1
    End If
    
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
    PrdGrid.TextMatrix(iR, 1) = txtPID.Text
    GetProductName
End Sub
Private Sub txtProduct_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 2) = txtProduct.Text
End Sub
Private Sub txtQty_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 3) = txtQty.Text
End Sub
Private Sub txtPrice_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 4) = txtPrice.Text
End Sub
Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    txtNT.Text = Val(txtQty.Text) * Val(txtPrice.Text)
End Sub
Private Sub txtPrice_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    txtNT.Text = Val(txtQty.Text) * Val(txtPrice.Text)
End Sub
Private Sub txtNT_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 5) = txtNT.Text
End Sub
Private Sub txtAP_KeyUp(KeyCode As Integer, Shift As Integer)
    CalculateDue
End Sub
Private Sub txtDiscount_KeyUp(KeyCode As Integer, Shift As Integer)
    CalculateDue
End Sub
Private Sub txtPrice_LostFocus()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 4) = txtPrice.Text
    PrdGrid.TextMatrix(iR, 5) = txtNT.Text
    
    GrandTotal = GrandTotal + Val(txtNT.Text)
    txtGT.Text = GrandTotal
    CalculateDue
End Sub

Private Sub PrdGrid_Click()
    On Error Resume Next
    iR = PrdGrid.Row
    iC = PrdGrid.Col
    
    If iC = 0 Then
        txtTID.Text = PrdGrid.TextMatrix(iR, iC)
    End If
    If iC = 1 Then
        txtPID.Text = PrdGrid.TextMatrix(iR, iC)
    End If
    If iC = 2 Then
        txtProduct.Text = PrdGrid.TextMatrix(iR, iC)
    End If
    If iC = 3 Then
        txtQty.Text = PrdGrid.TextMatrix(iR, iC)
        txtQty.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 4 Then
        txtPrice.Text = PrdGrid.TextMatrix(iR, iC)
        txtPrice.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 5 Then
        txtNT.Text = PrdGrid.TextMatrix(iR, iC)
        txtNT.SetFocus
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
Private Sub GridSet()
    With PrdGrid
    .Cols = 6
    .Rows = 2
    .ColWidth(0) = 1900
    .ColWidth(1) = 1900
    .ColWidth(2) = 2500
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
    .TextMatrix(0, 0) = " Transaction"
    .TextMatrix(0, 1) = " Product ID"
    .TextMatrix(0, 2) = " Product"
    .TextMatrix(0, 3) = " Quantity"
    .TextMatrix(0, 4) = " Price"
    .TextMatrix(0, 5) = " Net Total"
    End With
End Sub
Private Sub GenerateTID()
    txtTID.Text = "T" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub
Private Sub GenerateInvoiceNo()
    txtInv.Text = "IN" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Sub cmdSelectCus_Click()
    ParentForm = "frmInvoice"
    GridSQLString = "Select Customer_ID,Name,Mobile_No,Address from Customer ORDER BY Name"
    SelectedField = 0
    frmDataSelect.Show vbModal
    If txtCID.Text = "" Then
        frmCustomer.Show
        frmCustomer.EnterNewCustomer
    End If
End Sub
Private Sub cmdSelProd_Click()
    ParentForm = "frmInvoicePr"
    GridSQLString = "Select Product_ID,Product,Product_Type,Product_Size,Stock_In_Hand,Price_per_unit FROM Stock ORDER BY Product"
    SelectedField = 0
    frmDataSelect.Show vbModal
End Sub


Private Sub SetFields(TextFieldLock As Boolean)
    txtR.Enabled = TextFieldLock
    txtQty.Enabled = TextFieldLock
    txtPrice.Enabled = TextFieldLock
    'txtNT.Enabled = TextFieldLock
    txtDiscount.Enabled = TextFieldLock
    txtAP.Enabled = TextFieldLock
    PM.Enabled = TextFieldLock
    'txtGT.Enabled = TextFieldLock
End Sub

Private Sub SetButtons(ButtonLock As Boolean)
    cmdDelete.Enabled = ButtonLock
    cmdRDB.Enabled = ButtonLock
    cmdMF.Enabled = ButtonLock
    cmdML.Enabled = ButtonLock
    cmdN.Enabled = ButtonLock
    cmdP.Enabled = ButtonLock
    cmdSearch.Enabled = ButtonLock
End Sub

Private Sub ClearFields()
    txtInv.Text = ""
    txtDate.Text = ""
    txtTID.Text = ""
    txtCID.Text = ""
    txtR.Text = ""
    txtPID.Text = ""
    txtProduct.Text = ""
    txtQty.Text = ""
    txtPrice.Text = ""
    txtNT.Text = ""
    txtDiscount.Text = ""
    txtGT.Text = ""
    txtSalesman.Text = ""
    PM.Text = ""
    txtAP.Text = ""
    txtAD.Text = ""
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
    ClearFields
    SetFields (False)
    SetButtons (True)
    cmdNew.Visible = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdNew.Enabled = True

    cmdSelectCus.Enabled = False
    cmdSelProd.Enabled = False
    cmdAd.Enabled = False
    
    txtSearch.Enabled = True
    ST.Enabled = True
    
    PrdGrid.Rows = 1
    PrdGrid.Rows = 2
    
    iR = PrdGrid.Rows - 1
    cmdRDB_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    Rx = 0
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowInvoiceData ("SELECT * FROM Sales ORDER BY Invoice_No")
    ShowInvoiceGrid ("SELECT * FROM Invoice WHERE Invoice_No = '" & txtInv.Text & "' ORDER BY TID")
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
    
    SQLString = "SELECT * FROM Sales WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    sql2 = "SELECT * FROM Invoice WHERE Invoice_No='" & txtInv & "'"
    
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
       
    txtInv.Text = rs!Invoice_No
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    txtSalesman.Text = rs!Salesman
    txtCID.Text = rs!Customer_ID
    txtGT.Text = rs!Grand_Total
    txtDiscount.Text = rs!Discount
    txtGT.Text = rs!Grand_Total
    PM.Text = rs!Payment_Mode
    txtAP.Text = rs!Amount_Paid
    txtAD.Text = rs!Amount_Due
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
    If txtInv.Text = rs!Invoice_No Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub GetProductName()
    
    Set rsTmp = New ADODB.Recordset
    Query = "SELECT Product FROM Stock WHERE Product_ID='" & txtPID.Text & "'"

    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenStatic
    rsTmp.LockType = adLockReadOnly
    rsTmp.Open Query, conn
        If rsTmp.EOF = True Then
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Sub
        End If
    xCount = rsTmp.RecordCount
        If Rx > rsTmp.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsTmp.RecordCount - 1
        End If
    rsTmp.Move Rx
    
    txtProduct.Text = rsTmp!Product
    
    rsTmp.Close
    Set rsTmp = Nothing
    
End Sub

Private Function CheckQuanty(PID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT Stock_In_Hand FROM Stock WHERE Product_ID='" & PID & "'"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If Val(rs!Stock_In_Hand) > txtQty Then
        CheckQuanty = True
    Else
        CheckQuanty = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub ST_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtDiscount_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtGT_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtNT_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPrice_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtQty_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub PM_Change()
KeyAscii = 0
End Sub

Private Sub CalculateDue()
    txtAD.Text = (Val(txtGT.Text) - Val(txtDiscount.Text)) - Val(txtAP.Text)
End Sub
