Attribute VB_Name = "Module"
Option Explicit
Public UserName, Pass, LoginTime, UserTypeUsing, SQLString, SQLErr, GridSQLString, NewSupplier, RptName, Query, RptSql, RptStr, RptPathIs, RptDate1, RptDate2 As String
Public Starting, isStockMinus, isReOrder As Boolean

Public DateToday As String

Public SelectedField, n, c As Integer
Public ReturnValue As String
Public ParentForm As String

'SendMessage API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'CB Constants
Public Const CB_MAXLENGTH = 50
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_LIMITTEXT = &H141

'Mouse Cursor
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GCL_HCURSOR = (-12)
Public hOldCursor As Long

'Database
Public conn As ADODB.Connection
Public RsGrid As New ADODB.Recordset
Public RsLogin As New ADODB.Recordset

Public rsSupplier As New ADODB.Recordset
Public RsSuppGrid As New ADODB.Recordset

Public rsSupplierAccount As New ADODB.Recordset
Public RsSuppAccountGrid As New ADODB.Recordset

Public rsPO As New ADODB.Recordset
Public RsPOGrid As New ADODB.Recordset
Public rsPODetails As New ADODB.Recordset
Public RsPODetailsGrid As New ADODB.Recordset

Public rsProduct As New ADODB.Recordset
Public RsProductGrid As New ADODB.Recordset

Public rsReceivings As New ADODB.Recordset
Public RsReceivingsGrid As New ADODB.Recordset

Public rsCustomer As New ADODB.Recordset
Public RsCustomerGrid As New ADODB.Recordset

Public rsInvoice As New ADODB.Recordset
Public RsInvoiceGrid As New ADODB.Recordset

Public rsIncome As New ADODB.Recordset
Public RsIncomeGrid As New ADODB.Recordset
Public rsExpense As New ADODB.Recordset
Public RsExpenseGrid As New ADODB.Recordset

Public RsUser As New ADODB.Recordset
Public RsUserGrid As New ADODB.Recordset

Public rsGSale As New ADODB.Recordset
Public RsGSaleGrid As New ADODB.Recordset

Public rsCombo As New ADODB.Recordset
Public rsTmp As New ADODB.Recordset

Private ST As String
Public Rx, RxOS, RxIC, RxNS As Long
Public AddNewStatus As Boolean
Public xCount, xCountIC, xCountOS, xCountNS As Integer
Public db_name, db_server, db_port, db_user, db_pass, constr As String

'TextBox Limit
Public Const ES_NUMBER = &H2000&
Public Const GWL_STYLE = (-16)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Sub Main()
On Error GoTo DBerror

    db_name = "AK_INV"
    db_server = "localhost"
    db_port = ""
    db_user = "root"
    db_pass = "samsung"
    
    Connect
    
    Rx = 0
    RsLogin.Open "SELECT * FROM Login", conn
    RsLogin.Close
    
    frmSplash.Show
    'frmLogin.Show
    'MDIForm1.Show
    
    Exit Sub

DBerror:
    CreateDatabase
End Sub

Public Function Connect()
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
End Function

Private Sub CreateDatabase()
    On Error GoTo ServerErr
    
    'Drop...
    conn.Execute "DROP TABLE IF EXISTS Login"
    conn.Execute "DROP TABLE IF EXISTS Supplier_Account"
    conn.Execute "DROP TABLE IF EXISTS Customer_Account"
    conn.Execute "DROP TABLE IF EXISTS Customer"
    conn.Execute "DROP TABLE IF EXISTS Receivings"
    conn.Execute "DROP TABLE IF EXISTS Purchase_Order"
    conn.Execute "DROP TABLE IF EXISTS PO_Details"
    conn.Execute "DROP TABLE IF EXISTS Supplier"
    conn.Execute "DROP TABLE IF EXISTS Sales"
    conn.Execute "DROP TABLE IF EXISTS Invoice"
    conn.Execute "DROP TABLE IF EXISTS Stock"
    conn.Execute "DROP TABLE IF EXISTS Expenditure"
    conn.Execute "DROP TABLE IF EXISTS Income"
        
    'Create...
    conn.Execute "CREATE TABLE Login(User varchar(15), Password varchar(10), Type varchar(15), Name varchar(20), Designation varchar(20), Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Supplier(Supplier_ID varchar(20) Primary Key, Date Date, Name varchar(30), Company varchar(20), Contact_Person varchar(30), Address varchar(50), Office_No varchar(15), Mobile_No varchar(15), Other_No varchar(15), Fax_No varchar(15), Total_Bills_Amount int, Total_Due int, Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Customer(Customer_ID varchar(20) Primary Key, Date Date, Name varchar(30), CNIC_No varchar(15), Address varchar(50), Occupation varchar(30), Phone_No varchar(15), Mobile_No varchar(15), Other_No varchar(15), Total_Bills_Amount int, Total_Due int, Remarks varchar(50))", , adExecuteNoRecords
    'conn.Execute "CREATE TABLE Salesman(Salesman_ID varchar(20) Primary Key, Name varchar(30), CNIC_No varchar(15), Address varchar(50), Phone_No varchar(15), Mobile_No varchar(15), Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Sales(Invoice_No varchar(20) Primary Key, Date Date, Salesman varchar(20), Customer_ID varchar(20), Grand_Total int, Discount int, Payment_Mode varchar(15), Amount_Paid int, Amount_Due int, Remarks varchar(50), CONSTRAINT fk_cust_id FOREIGN KEY (Customer_ID) REFERENCES Customer(Customer_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE G_Sale(Invoice_No varchar(20) Primary Key, Date Date, Customer varchar(30), Salesman varchar(20),Product varchar(30), Description varchar(30), Quantity int, Price int, Total int, Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Purchase_Order(PO_No varchar(20) Primary Key, Date Date, Supplier_ID varchar(20), Delivery_Date Date, Remarks varchar(50), CONSTRAINT fk_supp_id1 FOREIGN KEY (Supplier_ID) REFERENCES Supplier(Supplier_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Stock(Product_ID varchar(20) Primary Key, Date Date, Product varchar(30), Product_Type varchar(20), Product_Size varchar(10), Company varchar(30), Stock_In_Hand int, Description varchar(25),Price_per_unit int, ReOrder_Level int, Remarks varchar(50))", , adExecuteNoRecords
    
    conn.Execute "CREATE TABLE Expenditure(TID varchar(20) Primary Key, Date Date, Expense_Type varchar(20),Supplier varchar(30), Payment_Mode varchar(20),Particulars varchar(30),Amount int,Remarks varchar(50))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Income(TID varchar(20) Primary Key, Date Date, Income_Type varchar(20),Customer varchar(30),Payment_Mode varchar(20),Particulars varchar(30), Amount int,Remarks varchar(50))", , adExecuteNoRecords
    
    conn.Execute "CREATE TABLE PO_Details(TID varchar(20) Primary Key, PO_No varchar(20), Product varchar(20), Product_Type varchar(20), Product_Size varchar(10), Quantity int, Description varchar(30), CONSTRAINT fk_PO1_No FOREIGN KEY (PO_No) REFERENCES Purchase_Order(PO_No))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Supplier_Account(TID varchar(20) Primary Key, Supplier_ID varchar(20), Date Date, PO_No varchar(20), Total_Amount int, Payment_Mode varchar(15), Paid_Amount int, Due_Amount int, Remarks varchar(50), CONSTRAINT fk_supp_id2 FOREIGN KEY (Supplier_ID) REFERENCES Supplier(Supplier_ID), CONSTRAINT fk_PO_No FOREIGN KEY (PO_No) REFERENCES Purchase_Order(PO_No))", , adExecuteNoRecords
    'conn.Execute "CREATE TABLE Invoice(TID varchar(20) Primary Key, Invoice_No varchar(20), Date Date, Salesman_ID varchar(20), Customer_ID varchar(20), Product_ID varchar(20), Quantity int, Price int, Net_Total int, Remarks varchar(50), CONSTRAINT fk_Inv_No FOREIGN KEY (Invoice_No) REFERENCES Sales(Invoice_No), CONSTRAINT fk_Sm_id FOREIGN KEY (Salesman_ID) REFERENCES Salesman(Salesman_ID), CONSTRAINT fk_cust_id FOREIGN KEY (Customer_ID) REFERENCES Customer(Customer_ID), CONSTRAINT fk_prod_id FOREIGN KEY (Product_ID) REFERENCES Stock(Product_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Invoice(TID varchar(20) Primary Key, Invoice_No varchar(20), Product_ID varchar(20), Quantity int, Price int, Net_Total int, CONSTRAINT fk_Inv_No FOREIGN KEY (Invoice_No) REFERENCES Sales(Invoice_No), CONSTRAINT fk_prod_id FOREIGN KEY (Product_ID) REFERENCES Stock(Product_ID))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Customer_Account(TID varchar(20) Primary Key, Customer_ID varchar(20), Date Date, Invoice_No varchar(20), Total_Amount int, Payment_Mode varchar(15), Amount_Paid int, Amount_Due int, Remarks varchar(50), CONSTRAINT fk_Inv1_No FOREIGN KEY (Invoice_No) REFERENCES Sales(Invoice_No))", , adExecuteNoRecords
    conn.Execute "CREATE TABLE Receivings(TID varchar(20) Primary Key, Date Date, PO_No varchar(20), Product_ID varchar(20), Quantity int, Price int, Price_per_unit int, Remarks varchar(50), CONSTRAINT fk_PO_No1 FOREIGN KEY (PO_No) REFERENCES Purchase_Order(PO_No), CONSTRAINT fk_prod_id3 FOREIGN KEY (Product_ID) REFERENCES Stock(Product_ID))", , adExecuteNoRecords
    
    'Insert...
    conn.Execute "INSERT INTO Login values('admin','admin','Admin','Admin User','Administration','-')", , adExecuteNoRecords
    conn.Execute "INSERT INTO Login values('manager','manager','Manager','Manager User','Management','-')", , adExecuteNoRecords
    conn.Execute "INSERT INTO Login values('salesman','salesman','Salesman','Sales User','Sales Dept','-')", , adExecuteNoRecords
    
    MsgBox "DATABASE CREATED & POPULATED WITH INITIAL DATA, Please RUN PROGRAM AGAIN!!!", vbInformation, "ADMIN"
    Exit Sub
ServerErr:
    MsgBox "Unable to Locate Database On Server, Please RUN PROGRAM AGAIN!!!", vbInformation, "DATABASE NOT FOUND"
    constr = "Provider=MSDASQL.1;Password=samsung;Persist Security Info=True;User ID=root;Extended Properties=" & Chr$(34) & "DRIVER={MySQL ODBC 3.51 Driver};DESC=;SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
    conn.Execute "Create Database " & Trim$("AK_INV"), , adExecuteNoRecords
    End
End Sub

Public Sub GetDate()
    Dim DateYear As String
    Dim DateMonth As String
    Dim DateDay As String

    DateYear = Year(Date)
    DateMonth = Month(Date)
    DateDay = Day(Date)
    
    DateToday = "" + DateYear + "-" + DateMonth + "-" + DateDay
End Sub

'Combo
Public Sub Combo_Lookup(ctlCombo As ComboBox)
   Dim lngItemPos As Long
   Dim strCombo As String

   strCombo = ctlCombo.Text

   ' Use SendMessage() API to Find Combobox Values
   lngItemPos = SendMessage(ctlCombo.hwnd, CB_FINDSTRING, -1, ByVal strCombo)

   If lngItemPos >= 0 Then
      ctlCombo.ListIndex = lngItemPos
   End If

   ctlCombo.SelStart = Len(strCombo)
   ctlCombo.SelLength = CB_MAXLENGTH
End Sub

Public Sub SendErrorReport(FormName As String)
    Dim Error As String
    Error = Err.Description
    'Save the Error to Database

End Sub

Public Sub UnloadForms()
    'Unload frmSecurity
    Unload frmSupplier
End Sub

Public Sub CheckUser()
    If (UserTypeUsing = "Manager") Then
        MDIForm1.mnPurchase.Enabled = False
        MDIForm1.mnSM.Enabled = False
        MDIForm1.mnSales.Enabled = False
        MDIForm1.mnUM.Enabled = False
    End If
    If (UserTypeUsing = "Salesman") Then
        MDIForm1.mnPurchase.Enabled = False
        MDIForm1.mnSM.Enabled = False
        MDIForm1.mnUM.Enabled = False
    End If
    
End Sub
