VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMoneyTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Declare Type for all Variables
Private Type MoneyTransaction
    D_C                     As String
    TRANSACTION_key         As Long
    TRANSACTION_No          As Long
    Type_Of_Account         As String
    Account_No              As Long
    Type_Of_AccountTo       As String
    Account_NoTo            As Long
    TYPE_OF_TRANSACTION     As String
    Payment_Withdrawl       As String
    DATE_OF_TRANSACTION     As Date
    CHECK_NO                As Long
    CHECK_DATE              As Date
    TOKEN_NO                As Long
    SCROLL_NO               As Long
    Remarks                 As String
    Gl_No                   As String
    Amount                  As Double
    TerminalName            As String
    UserName                As String
    InsertModifyDate        As Date
    FD_NO                   As String
    name                    As String
    Address                 As String
    Fd_Amount               As Double
    Maturity_Amt            As Double
    Fd_Date                 As Date
    ROI                     As Double
    Approved_By1            As String
    Approved_By2            As String
    Approved_By3            As String
    New_ROI_FD              As Byte
    OpeningBalance          As Double
    TOTDb                   As Double
    TOTCr                   As Double
    FD_Matured              As Byte
    Payment                 As Byte
    Closed_Operative        As String
    Severally_Joint         As String
    Item                    As String
    Opening_Date            As Date
    Period                  As Integer
    Installment             As Double
    End Type

'Public Enum dbState
'    DBInsert = 1
'    DBModify = 2
'    DBDelete = 3
'End Enum
Private SaveMoneyTransaction     As MoneyTransaction
Public Property Let Closed_Operative(ByVal v_Value As String)
    SaveMoneyTransaction.Closed_Operative = v_Value
End Property
Public Property Let Severally_Joint(ByVal v_Value As String)
    SaveMoneyTransaction.Severally_Joint = v_Value
End Property
Public Property Let Item(ByVal v_Value As String)
    SaveMoneyTransaction.Item = v_Value
End Property
Public Property Let Opening_Date(ByVal v_Value As Date)
    SaveMoneyTransaction.Opening_Date = v_Value
End Property
Public Property Let Period(ByVal v_Value As Integer)
    SaveMoneyTransaction.Period = v_Value
End Property
Public Property Let Installment(ByVal v_Value As Double)
    SaveMoneyTransaction.Installment = v_Value
End Property

Public Property Get Closed_Operative() As String
    Closed_Operative = SaveMoneyTransaction.Closed_Operative
End Property
Public Property Get Severally_Joint() As String
    Severally_Joint = SaveMoneyTransaction.Severally_Joint
End Property
Public Property Get Item() As String
    Item = SaveMoneyTransaction.Item
End Property
Public Property Get Opening_Date() As Date
    Opening_Date = SaveMoneyTransaction.Opening_Date
End Property
Public Property Get Period() As Integer
    Period = SaveMoneyTransaction.Period
End Property
Public Property Get Installment() As Double
    Installment = SaveMoneyTransaction.Installment
End Property

Public Property Let FD_Matured(ByVal v_Value As Byte)
    SaveMoneyTransaction.FD_Matured = v_Value
End Property
Public Property Let Payment(ByVal v_Value As Byte)
    SaveMoneyTransaction.Payment = v_Value
End Property
Public Property Get FD_Matured() As Byte
    FD_Matured = SaveMoneyTransaction.FD_Matured
End Property
Public Property Get Payment() As Byte
    Payment = SaveMoneyTransaction.Payment
End Property

Public Property Let Approved_By1(ByVal v_Value As String)
    SaveMoneyTransaction.Approved_By1 = v_Value
End Property
Public Property Let Approved_By2(ByVal v_Value As String)
    SaveMoneyTransaction.Approved_By2 = v_Value
End Property
Public Property Let Approved_By3(ByVal v_Value As String)
    SaveMoneyTransaction.Approved_By3 = v_Value
End Property
Public Property Get Approved_By1() As String
    Approved_By1 = SaveMoneyTransaction.Approved_By1
End Property
Public Property Get Approved_By2() As String
    Approved_By2 = SaveMoneyTransaction.Approved_By2
End Property
Public Property Get Approved_By3() As String
    Approved_By3 = SaveMoneyTransaction.Approved_By3
End Property
Public Property Let D_C(ByVal v_Value As String)
    SaveMoneyTransaction.D_C = v_Value
End Property
Public Property Let Fd_Date(ByVal v_Value As Date)
    SaveMoneyTransaction.Fd_Date = v_Value
End Property
Public Property Let ROI(ByVal v_Value As Double)
    SaveMoneyTransaction.ROI = v_Value
End Property
Public Property Let New_ROI_FD(ByVal v_Value As Double)
    SaveMoneyTransaction.New_ROI_FD = v_Value
End Property
Public Property Let Fd_Amount(ByVal v_Value As Double)
    SaveMoneyTransaction.Fd_Amount = v_Value
End Property
Public Property Let Maturity_Amt(ByVal v_Value As Double)
    SaveMoneyTransaction.Maturity_Amt = v_Value
End Property
Public Property Let name(ByVal v_Value As String)
    SaveMoneyTransaction.name = v_Value
End Property
Public Property Let Address(ByVal v_Value As String)
    SaveMoneyTransaction.Address = v_Value
End Property
Public Property Let FD_NO(ByVal v_Value As String)
    SaveMoneyTransaction.FD_NO = v_Value
End Property
Public Property Let TRANSACTION_key(ByVal v_Value As Long)
    SaveMoneyTransaction.TRANSACTION_key = v_Value
End Property
Public Property Let TRANSACTION_No(ByVal v_Value As Long)
    SaveMoneyTransaction.TRANSACTION_No = v_Value
End Property
Public Property Let Type_Of_Account(ByVal v_Value As String)
    SaveMoneyTransaction.Type_Of_Account = v_Value
End Property
Public Property Let Type_Of_AccountTo(ByVal v_Value As String)
    SaveMoneyTransaction.Type_Of_AccountTo = v_Value
End Property
Public Property Let Account_No(ByVal v_Value As Long)
    SaveMoneyTransaction.Account_No = v_Value
End Property
Public Property Let Account_NoTo(ByVal v_Value As Long)
    SaveMoneyTransaction.Account_NoTo = v_Value
End Property
Public Property Let TYPE_OF_TRANSACTION(ByVal v_Value As String)
    SaveMoneyTransaction.TYPE_OF_TRANSACTION = v_Value
End Property
Public Property Let Payment_Withdrawl(ByVal v_Value As String)
    SaveMoneyTransaction.Payment_Withdrawl = v_Value
End Property
Public Property Let DATE_OF_TRANSACTION(ByVal v_Value As Date)
    SaveMoneyTransaction.DATE_OF_TRANSACTION = v_Value
End Property
Public Property Let CHECK_NO(ByVal v_Value As Long)
    SaveMoneyTransaction.CHECK_NO = v_Value
End Property
Public Property Let CHECK_DATE(ByVal v_Value As Date)
    SaveMoneyTransaction.CHECK_DATE = v_Value
End Property
Public Property Let TOKEN_NO(ByVal v_Value As Long)
    SaveMoneyTransaction.TOKEN_NO = v_Value
End Property
Public Property Let SCROLL_NO(ByVal v_Value As Long)
    SaveMoneyTransaction.SCROLL_NO = v_Value
End Property
Public Property Let Remarks(ByVal v_Value As String)
    SaveMoneyTransaction.Remarks = v_Value
End Property
Public Property Let Gl_No(ByVal v_Value As String)
    SaveMoneyTransaction.Gl_No = v_Value
End Property
Public Property Let Amount(ByVal v_Value As Double)
    SaveMoneyTransaction.Amount = v_Value
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveMoneyTransaction.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveMoneyTransaction.UserName = v_Value
End Property
Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveMoneyTransaction.InsertModifyDate = v_Value
End Property
Public Property Let OpeningBalance(ByVal v_Value As Double)
    SaveMoneyTransaction.OpeningBalance = v_Value
End Property
Public Property Let TOTDb(ByVal v_Value As Double)
    SaveMoneyTransaction.TOTDb = v_Value
End Property
Public Property Let TOTCr(ByVal v_Value As Double)
    SaveMoneyTransaction.TOTCr = v_Value
End Property
'Public Property Let OpeningBalance(ByVal v_Value As Double)
'    SaveMoneyTransaction.OpeningBalance = v_Value
'End Property

Public Property Get ROI() As Double
    ROI = SaveMoneyTransaction.ROI
End Property
Public Property Get New_ROI_FD() As Double
    New_ROI_FD = SaveMoneyTransaction.New_ROI_FD
End Property

Public Property Get Fd_Amount() As Double
    Fd_Amount = SaveMoneyTransaction.Fd_Amount
End Property
Public Property Get Fd_Date() As Date
    Fd_Date = SaveMoneyTransaction.Fd_Date
End Property
Public Property Get Maturity_Amt() As Double
    Maturity_Amt = SaveMoneyTransaction.Maturity_Amt
End Property
Public Property Get name() As String
    name = SaveMoneyTransaction.name
End Property
Public Property Get Address() As String
    Address = SaveMoneyTransaction.Address
End Property
Public Property Get FD_NO() As String
    FD_NO = SaveMoneyTransaction.FD_NO
End Property
Public Property Get D_C() As String
    D_C = SaveMoneyTransaction.D_C
End Property
Public Property Get TRANSACTION_key() As Long
    TRANSACTION_No = SaveMoneyTransaction.TRANSACTION_key
End Property
Public Property Get TRANSACTION_No() As Long
    TRANSACTION_No = SaveMoneyTransaction.TRANSACTION_No
End Property
Public Property Get Type_Of_Account() As String
    Type_Of_Account = SaveMoneyTransaction.Type_Of_Account
End Property
Public Property Get Type_Of_AccountTo() As String
    Type_Of_AccountTo = SaveMoneyTransaction.Type_Of_AccountTo
End Property
Public Property Get Account_No() As Long
    Account_No = SaveMoneyTransaction.Account_No
End Property
Public Property Get Account_NoTo() As Long
    Account_NoTo = SaveMoneyTransaction.Account_NoTo
End Property
Public Property Get TYPE_OF_TRANSACTION() As String
    TYPE_OF_TRANSACTION = SaveMoneyTransaction.TYPE_OF_TRANSACTION
End Property
Public Property Get Payment_Withdrawl() As String
    Payment_Withdrawl = SaveMoneyTransaction.Payment_Withdrawl
End Property
Public Property Get DATE_OF_TRANSACTION() As Date
    DATE_OF_TRANSACTION = SaveMoneyTransaction.DATE_OF_TRANSACTION
End Property
Public Property Get CHECK_NO() As Long
    CHECK_NO = SaveMoneyTransaction.CHECK_NO
End Property
Public Property Get CHECK_DATE() As Date
    CHECK_DATE = SaveMoneyTransaction.CHECK_DATE
End Property
Public Property Get TOKEN_NO() As Long
    TOKEN_NO = SaveMoneyTransaction.TOKEN_NO
End Property
Public Property Get SCROLL_NO() As Long
    SCROLL_NO = SaveMoneyTransaction.SCROLL_NO
End Property
Public Property Get Remarks() As String
    Remarks = SaveMoneyTransaction.Remarks
End Property
Public Property Get Gl_No() As String
    Gl_No = SaveMoneyTransaction.Gl_No
End Property
Public Property Get Amount() As Double
    Amount = SaveMoneyTransaction.Amount
End Property
Public Property Get TerminalName() As String
    TerminalName = SaveMoneyTransaction.TerminalName
End Property
Public Property Get UserName() As String
    UserName = SaveMoneyTransaction.UserName
End Property
Public Property Get InsertModifyDate() As Date
 InsertModifyDate = SaveMoneyTransaction.InsertModifyDate
End Property
Public Property Get OpeningBalance() As Double
    OpeningBalance = SaveMoneyTransaction.OpeningBalance
End Property
Public Property Get TOTDb() As Double
    TOTDb = SaveMoneyTransaction.TOTDb
End Property
Public Property Get TOTCr() As Double
    TOTCr = SaveMoneyTransaction.TOTCr
End Property
Public Sub UpdateIsPaidFD(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    If SaveMoneyTransaction.Type_Of_Account <> "FD" And SaveMoneyTransaction.Type_Of_Account <> "MM" Then
        Exit Sub
    End If
    With objCommand
             .CommandText = "kccBank.PACK_Money_Transaction.UpdateIsPaidFD"
             .CommandType = adCmdStoredProc
             'Send connection
             .ActiveConnection = objDatasource
             'Set command time out
             .CommandTimeout = 0
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
             .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
             .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
             .Execute
     End With
             Set objCommand = Nothing
End Sub
Public Sub UpdateIsPaidFDRev(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    If SaveMoneyTransaction.Type_Of_Account <> "FD" And SaveMoneyTransaction.Type_Of_Account <> "MM" Then
        Exit Sub
    End If
    With objCommand
             .CommandText = "kccBank.PACK_Money_Transaction.UpdateIsPaidFDRev"
             .CommandType = adCmdStoredProc
             'Send connection
             .ActiveConnection = objDatasource
             'Set command time out
             .CommandTimeout = 0
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
             .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
             .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
             .Execute
     End With
             Set objCommand = Nothing
End Sub

Public Sub UpdateIsPaidRD(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    If SaveMoneyTransaction.Type_Of_Account = "RD" Then
        Exit Sub
    End If
    With objCommand
             .CommandText = "kccBank.PACK_Money_Transaction.UpdateIsPaidRD"
             .CommandType = adCmdStoredProc
             'Send connection
             .ActiveConnection = objDatasource
             'Set command time out
             .CommandTimeout = 0
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
             .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
             .Execute
    End With
             Set objCommand = Nothing
End Sub

Public Sub UpdateIsPaidRDRev(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    If SaveMoneyTransaction.Type_Of_Account = "RD" Then
        Exit Sub
    End If
    With objCommand
             .CommandText = "kccBank.PACK_Money_Transaction.UpdateIsPaidRDRev"
             .CommandType = adCmdStoredProc
             'Send connection
             .ActiveConnection = objDatasource
             'Set command time out
             .CommandTimeout = 0
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
             .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
             .Execute
    End With
             Set objCommand = Nothing
End Sub


Public Sub INSERTData(ByRef objDatasource As ADODB.Connection, ByVal DebitCredit As Byte, ByVal dowork As Byte, CashTransfer As Byte)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    'For Debit and Credit
     If DebitCredit = 1 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTDebitDATA"
     ElseIf DebitCredit = 2 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTCreditDATA"
     End If
         .CommandType = adCmdStoredProc
         'Send connection
         .ActiveConnection = objDatasource
         'Set command time out
         .CommandTimeout = 0
     If dowork = 1 And CashTransfer = 2 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        If SaveMoneyTransaction.TYPE_OF_TRANSACTION = "Q" Then
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RC")
        ElseIf SaveMoneyTransaction.TYPE_OF_TRANSACTION = "R" Then
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IT")
        Else
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
        End If
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 1 And CashTransfer = 0 Then
          'Bank account
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        If SaveMoneyTransaction.Type_Of_Account = "CL" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "LI")
        Else
            If SaveMoneyTransaction.TYPE_OF_TRANSACTION = "Q" Then
                 .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RC")
            Else
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
            End If
        End If
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
      ElseIf dowork = 1 And CashTransfer = 1 Then
         'Transfer To Account
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_NoTo)
     End If
     If dowork = 3 And CashTransfer = 1 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
     End If
     If dowork = 2 Then
          'Main Account
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
     End If
     'Set Parameters
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.TYPE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Payment_Withdrawl)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveMoneyTransaction.CHECK_NO)
     If SaveMoneyTransaction.CHECK_DATE = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
     Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.CHECK_DATE)
     End If
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.SCROLL_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveMoneyTransaction.Amount)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Gl_No)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.ROI)
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
     If DebitCredit = 1 Then
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
     End If
     .Execute
     End With
    If DebitCredit = 1 Then
     If Not IsNull(objCommand("Param1").Value) Then
          'Return Transaction Key value and useet it for Credit
          SaveMoneyTransaction.TRANSACTION_key = CLng(objCommand("Param1").Value)
     End If
    End If
     Set objCommand = Nothing
End Sub
Public Sub INSERTDataFD(ByRef objDatasource As ADODB.Connection, ByVal DebitCredit As Byte, ByVal dowork As Byte, CashTransfer As Byte)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    'For Debit and Credit
     If DebitCredit = 1 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTDebitDATA"
     ElseIf DebitCredit = 2 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTCreditDATA"
     End If
         .CommandType = adCmdStoredProc
         'Send connection
         .ActiveConnection = objDatasource
         'Set command time out
         .CommandTimeout = 0
     If dowork = 0 And CashTransfer = 0 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         '.Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FD")
         If SaveMoneyTransaction.FD_Matured = 1 Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FM")
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
         Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
         End If
     End If
     If dowork = 2 And CashTransfer = 2 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IF")
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 4 And CashTransfer = 4 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If

     If dowork = 3 And CashTransfer = 3 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
    If dowork = 5 And CashTransfer = 5 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IF")
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
    If dowork = 6 And CashTransfer = 6 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
    End If
     If dowork = 1 And CashTransfer = 1 Then
         'Transfer To Account
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         If SaveMoneyTransaction.TYPE_OF_TRANSACTION = "T" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_NoTo)
         Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
         End If
     End If
     'Set Parameters
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.TYPE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Payment_Withdrawl)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveMoneyTransaction.CHECK_NO)
     If SaveMoneyTransaction.CHECK_DATE = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
     Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.CHECK_DATE)
     End If
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.SCROLL_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveMoneyTransaction.Amount)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Gl_No)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.ROI)
     If SaveMoneyTransaction.Payment = 1 Then
         If SaveMoneyTransaction.FD_Matured = 1 Then
           .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, "Y")
        Else
           .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, "N")
        End If
      Else
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
      End If
     If DebitCredit = 1 Then
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
     End If
     .Execute
     End With
    If DebitCredit = 1 Then
     If Not IsNull(objCommand("Param1").Value) Then
          'Return Transaction Key value and useet it for Credit
          SaveMoneyTransaction.TRANSACTION_key = CLng(objCommand("Param1").Value)
     End If
    End If
     Set objCommand = Nothing
End Sub
Public Sub INSERTFDInterest(ByRef objDatasource As ADODB.Connection, ByVal DebitCredit As Byte, ByVal dowork As Byte)

    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    'For Debit and Credit
     If DebitCredit = 1 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTDebitDATA"
     ElseIf DebitCredit = 2 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTCreditDATA"
     End If
         .CommandType = adCmdStoredProc
         'Send connection
         .ActiveConnection = objDatasource
         'Set command time out
         .CommandTimeout = 0
     If dowork = 1 Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 2 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
      End If
      If dowork = 3 Then
         'Transfer To Account
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_NoTo)
     End If
     'Set Parameters
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.TYPE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Payment_Withdrawl)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveMoneyTransaction.CHECK_NO)
     If SaveMoneyTransaction.CHECK_DATE = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
     Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.CHECK_DATE)
     End If
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.SCROLL_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveMoneyTransaction.Amount)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Gl_No)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.ROI)
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
     If DebitCredit = 1 Then
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
     End If
     .Execute
     End With
    If DebitCredit = 1 Then
     If Not IsNull(objCommand("Param1").Value) Then
          'Return Transaction Key value and useet it for Credit
          SaveMoneyTransaction.TRANSACTION_key = CLng(objCommand("Param1").Value)
     End If
    End If
     Set objCommand = Nothing
End Sub
Public Sub INSERTDataRD(ByRef objDatasource As ADODB.Connection, ByVal DebitCredit As Byte, ByVal dowork As Byte, CashTransfer As Byte)
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    'For Debit and Credit
     If DebitCredit = 1 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTDebitDATA"
     ElseIf DebitCredit = 2 Then
         .CommandText = "kccBank.PACK_Money_Transaction.INSERTCreditDATA"
     End If
         .CommandType = adCmdStoredProc
         'Send connection
         .ActiveConnection = objDatasource
         'Set command time out
         .CommandTimeout = 0
     If dowork = 0 And CashTransfer = 0 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         If SaveMoneyTransaction.FD_Matured = 1 Then
            If SaveMoneyTransaction.Type_Of_Account = "MM" Then
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FM")
            Else
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RM")
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
         Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
         End If
     End If
     If dowork = 2 And CashTransfer = 2 Then
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         If SaveMoneyTransaction.Type_Of_Account = "MM" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
         Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RI")
         End If
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 1 And CashTransfer = 1 Then
         'Transfer To Account
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
         .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
         If SaveMoneyTransaction.TYPE_OF_TRANSACTION = "T" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_NoTo)
         Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
         End If
     End If
     If dowork = 4 Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        If SaveMoneyTransaction.Type_Of_Account = "MM" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
        Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RI")
        End If
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 5 Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        If SaveMoneyTransaction.Type_Of_Account = "MM" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IF")
        Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IR")
        End If
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 6 Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        If SaveMoneyTransaction.Type_Of_Account = "MM" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IF")
        Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "IR")
        End If
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If
     If dowork = 7 Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_No)
        If SaveMoneyTransaction.Type_Of_Account = "MM" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
        Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RI")
        End If
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
     End If

     'Set Parameters
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.TYPE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Payment_Withdrawl)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveMoneyTransaction.CHECK_NO)
     If SaveMoneyTransaction.CHECK_DATE = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
     Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.CHECK_DATE)
     End If
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.SCROLL_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveMoneyTransaction.Amount)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Gl_No)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
     .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
     .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
     .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.ROI)
     If SaveMoneyTransaction.Payment = 1 Then
        If SaveMoneyTransaction.FD_Matured = 1 Then
           .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, "Y")
        Else
           .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, "N")
        End If
     Else
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
     End If
     If DebitCredit = 1 Then
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
     End If
     .Execute
     End With
    If DebitCredit = 1 Then
     If Not IsNull(objCommand("Param1").Value) Then
          'Return Transaction Key value and useet it for Credit
          SaveMoneyTransaction.TRANSACTION_key = CLng(objCommand("Param1").Value)
     End If
    End If
     Set objCommand = Nothing
End Sub

Public Function ModifyData(ByRef objDatasource As ADODB.Connection, ByVal DebitCredit As Byte, ByVal dowork As Byte, CashTransfer As Byte) As Boolean
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    If DebitCredit = 1 Then
        .CommandText = "kccBank.PACK_Money_Transaction.ModifyDebitDATA"
    ElseIf DebitCredit = 2 Then
        .CommandText = "kccBank.PACK_Money_Transaction.ModifyCreditDATA"
    End If
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
    If dowork = 1 And CashTransfer = 0 Then
        'Bank Account
        If SaveMoneyTransaction.Type_Of_Account = "CL" Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "LI")
        ElseIf SaveMoneyTransaction.TYPE_OF_TRANSACTION = "Q" Then
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "RC")
        Else
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
        End If
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
    ElseIf dowork = 2 And CashTransfer = 0 Then
        'Main Account
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_No)
    End If
    If (dowork = 2 Or dowork = 1) And CashTransfer = 1 Then
        'Transfer To account
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_NoTo)
    End If
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.TYPE_OF_TRANSACTION)
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Payment_Withdrawl)
    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveMoneyTransaction.CHECK_NO)
    If SaveMoneyTransaction.CHECK_DATE = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
     Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.CHECK_DATE)
     End If
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.SCROLL_NO)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveMoneyTransaction.Amount)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Gl_No)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.ROI)
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, Null)
    .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
    .Execute
     End With
    If Not IsNull(objCommand("Param1").Value) Then
         If (objCommand("Param1").Value) = 0 Then
            'if modify Successfully
            ModifyData = True
         Else
            ModifyData = False
         End If
    End If
    Set objCommand = Nothing
End Function
Public Function ModifyFDInterest(ByRef objDatasource As ADODB.Connection, ByVal DebitCredit As Byte, ByVal dowork As Byte) As Boolean
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    If DebitCredit = 1 Then
        .CommandText = "kccBank.PACK_Money_Transaction.ModifyDebitDATA"
    ElseIf DebitCredit = 2 Then
        .CommandText = "kccBank.PACK_Money_Transaction.ModifyCreditDATA"
    End If
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.TRANSACTION_key)
    If dowork = 1 Then
        'Bank Account
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FI")
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
    ElseIf dowork = 2 Then
        'Main Account
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "CH")
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, Null)
    End If
    If dowork = 3 Then
        'Transfer To account
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.Account_NoTo)
    End If
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.TYPE_OF_TRANSACTION)
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Payment_Withdrawl)
    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveMoneyTransaction.CHECK_NO)
    If SaveMoneyTransaction.CHECK_DATE = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
     Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.CHECK_DATE)
     End If
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 8, SaveMoneyTransaction.SCROLL_NO)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveMoneyTransaction.Amount)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Gl_No)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.ROI)
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
    .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
    .Execute
     End With
    If Not IsNull(objCommand("Param1").Value) Then
         If (objCommand("Param1").Value) = 0 Then
            'if modify Successfully
            ModifyFDInterest = True
         Else
            ModifyFDInterest = False
         End If
    End If
    Set objCommand = Nothing
End Function

Public Function CheckChequeNo(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.CheckChequeNo"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.CHECK_NO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput)
        .Parameters.Append .CreateParameter("Param4", adChar, adParamOutput, 2)
        .Parameters.Append .CreateParameter("Param5", adNumeric, adParamOutput)
        
        'Execute to fetch records
        .Execute
        If Not IsNull(objCommand("Param3").Value) Then
            If CLng(objCommand("Param3").Value) > 0 Then
                CheckChequeNo = True
                SaveMoneyTransaction.Type_Of_Account = objCommand("Param4")
                SaveMoneyTransaction.Account_No = objCommand("Param5")
            Else
                CheckChequeNo = False
            End If
        End If
'        If Not IsNull(objCommand("Param1").Value) Then
'            SaveMoneyTransaction.Type_Of_Account = objCommand("Param1").Value
'        End If
'        If Not IsNull(objCommand("Param2").Value) Then
'            SaveMoneyTransaction.Account_No = objCommand("Param2").Value
'        End If
    End With
    Set objCommand = Nothing
End Function
Public Function ChequeUsed(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.ChequeUsed"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.CHECK_NO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
        .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
                ChequeUsed = True
            Else
                ChequeUsed = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function ChequeIssued(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.ChequeIssued"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.Account_No)
        'Execute to fetch records
        Set ChequeIssued = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Function CheckAccountNo(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.CheckAccountNo"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.Account_No)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
        .Execute
          If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
                CheckAccountNo = True
            Else
                CheckAccountNo = False
            End If
        End If
End With
    Set objCommand = Nothing
End Function
Public Function CheckIs_Paid(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        If SaveMoneyTransaction.Type_Of_Account = "FL" Then
            .CommandText = "kccBank.PACK_Money_Transaction.CheckIs_PaidFD"
        ElseIf SaveMoneyTransaction.Type_Of_Account = "RL" Then
            .CommandText = "kccBank.PACK_Money_Transaction.CheckIs_PaidRD"
        End If
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.Account_No)
        .Parameters.Append .CreateParameter("Param1", adChar, adParamOutput, 1)
        'Execute to fetch records
        .Execute
          If Not IsNull(objCommand("Param1").Value) Then
            If (objCommand("Param1").Value) = "Y" Or (objCommand("Param1").Value) = "T" Then
                CheckIs_Paid = True
            Else
                CheckIs_Paid = False
            End If
        End If
End With
    Set objCommand = Nothing
End Function

Public Function CheckTokenNo(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.CheckTokenNo"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 6, SaveMoneyTransaction.TOKEN_NO)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
        .Execute
          If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
                CheckTokenNo = True
            Else
                CheckTokenNo = False
            End If
        End If
End With
    Set objCommand = Nothing
End Function
Public Function CheckStopPayment(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    'If Cheque is Exists the return its Type of acccount,Account No
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.CheckStopPayment"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.CHECK_NO)
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput)
        'Execute to fetch records
        .Execute
        If Not IsNull(objCommand("Param3").Value) Then
            If CLng(objCommand("Param3").Value) > 0 Then
                CheckStopPayment = True
            Else
                CheckStopPayment = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function CheckAccount(ByRef objDatasource As ADODB.Connection) As String
    'Check Account if it Exists
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.CheckAccount"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.Account_No)
        .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
        'Execute to fetch records
        .Execute
        If Not (objCommand("Param1").Value) = "-1" Then
            CheckAccount = objCommand("Param1").Value
         Else
                CheckAccount = Empty
        End If
    End With
    Set objCommand = Nothing
End Function
'Public Function CheckAccountNoTo(ByRef objDatasource As adodb.Connection) As String
'    'Check Account if it Exists
'    Dim objCommand As New adodb.Command
'    With objCommand
'        .CommandText = "kccBank.PACK_Money_Transaction.CheckAccount"
'        .CommandType = adCmdStoredProc
'        'Send connection
'        .ActiveConnection = objDatasource
'        'Set command time out
'        .CommandTimeout = 0
'        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
'        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.Account_NoTo)
'        .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
'        'Execute to fetch records
'        .Execute
'        If Not (objCommand("Param1").Value) = -1 Then
'            CheckAccountNoTo = objCommand("Param1").Value
'         Else
'                CheckAccountNoTo = Empty
'        End If
'    End With
'    Set objCommand = Nothing
'End Function
Public Function CheckAccountNoTo(ByRef objDatasource As ADODB.Connection, ByVal dowork As Byte) As String
    'Check Account if it Exists
    Dim objCommand As New ADODB.Command
    With objCommand
        If dowork = 1 Then
            .CommandText = "kccBank.PACK_Money_Transaction.CheckAccount"
        ElseIf dowork = 2 Then
            .CommandText = "kccBank.PACK_Money_Transaction.CheckAccountRD"
        End If
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_AccountTo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.Account_NoTo)
        .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
        'Execute to fetch records
        .Execute
        If Not (objCommand("Param1").Value) = -1 Then
            CheckAccountNoTo = objCommand("Param1").Value
         Else
                CheckAccountNoTo = Empty
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function CheckAccountNOToFD(ByRef objDatasource As ADODB.Connection) As String
    'Check Account if it Exists
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.CheckAccountFD"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
        .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
        'Execute to fetch records
        .Execute
        If Not (objCommand("Param1").Value) = -1 Then
            CheckAccountNOToFD = objCommand("Param1").Value
         Else
                CheckAccountNOToFD = Empty
        End If
    End With
    Set objCommand = Nothing
End Function


Public Function GetAccountType(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "{call KccBank.PACK_ACCOUNT_OPENING.GetAccountType()}"
            .CommandType = adCmdText
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            'Execute to fetch records
            Set GetAccountType = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GetScrollNum(ByRef objDatasource As ADODB.Connection) As Long
    'Get Max Scroll No
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.GetScrollNum"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DATE_OF_TRANSACTION)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
        .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            GetScrollNum = objCommand("Param1").Value
         Else
                GetScrollNum = Empty
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function GetTransactionNum(ByRef objDatasource As ADODB.Connection) As Long
    'Get Max Transaction No.
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccBank.PACK_Money_Transaction.GetTransactionNum"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
        .Execute
        If Not IsNull(objCommand("Param1").Value) Then
                GetTransactionNum = objCommand("Param1").Value
         Else
                GetTransactionNum = Empty
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function GetApproval(ByRef objDatasource As ADODB.Connection) _
                                As Byte
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GetApproval"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.TRANSACTION_key)
            .Parameters.Append .CreateParameter("Param1", adChar, adParamOutput, 1)
            .Parameters.Append .CreateParameter("Param2", adChar, adParamOutput, 1)
            .Parameters.Append .CreateParameter("Param3", adChar, adParamOutput, 1)
            .Parameters.Append .CreateParameter("Param4", adNumeric, adParamOutput)
            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param4").Value) Then
                GetApproval = objCommand("Param4").Value
                If objCommand("param1").Value = "N" Then
                    SaveMoneyTransaction.Approved_By1 = 0
                Else
                        SaveMoneyTransaction.Approved_By1 = 1
                End If
                If objCommand("param2").Value = "N" Then
                    SaveMoneyTransaction.Approved_By2 = 0
                Else
                        SaveMoneyTransaction.Approved_By2 = 1
                End If
                If objCommand("param3").Value = "N" Then
                    SaveMoneyTransaction.Approved_By3 = 0
                Else
                    SaveMoneyTransaction.Approved_By3 = 1
                End If
                GetApproval = 0
            End If
            Set objCommand = Nothing
End Function

Public Function DeleteData(ByRef objDatasource As ADODB.Connection, ByVal dowork As Byte) _
                                As Boolean
    'Delete Record from database
    Dim objCommand As New ADODB.Command
    With objCommand
        If dowork = 1 Then
            .CommandText = "PACK_Money_Transaction.DeleteDataDebit"
        ElseIf dowork = 2 Then
            .CommandText = "PACK_Money_Transaction.DeleteDataCredit"
        ElseIf dowork = 3 Then
            .CommandText = "PACK_Money_Transaction.DeleteData"
        End If
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        If dowork = 3 Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.TRANSACTION_No)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        Else
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.TRANSACTION_No)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        End If
        'Execute to fetch records
         .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            If objCommand("Param1").Value = 0 Then
                DeleteData = True
            Else
                DeleteData = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function PopupLov(ByRef objDatasource As ADODB.Connection) As Boolean
    'Get the fields from t_Transaction_credit and t_transaction_debit table Using LOV
    'Make the object of LOV
    Dim look As New Lov.LookUp
    Dim strProcedure As String
    Dim sqlString As String
    If SaveMoneyTransaction.Type_Of_Account = "RD" Then
        sqlString = "select Type_Of_Account , Account_No ,Amount, TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='RD' And Delete_Flag='N' And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount, TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where    TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='RD' And Delete_Flag='N'"
    ElseIf SaveMoneyTransaction.Type_Of_Account = "MM" Then
                sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='MM' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where    TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='MM' And Delete_Flag='N'"
    
    ElseIf SaveMoneyTransaction.Type_Of_Account = "FD" Then
                sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='FD' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where    TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='FD' And Delete_Flag='N'"

    ElseIf SaveMoneyTransaction.Type_Of_Account = "RL" Then
                sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='RL' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where    TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='RL' And Delete_Flag='N'"
                
    ElseIf SaveMoneyTransaction.Type_Of_Account = "FL" Then
                sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='FL' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where    TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='FL' And Delete_Flag='N'"
    ElseIf SaveMoneyTransaction.Type_Of_Account = "SB" Then
        sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='SB' AND TYPE_OF_TRANSACTION<>'D' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='SB' And Delete_Flag='N'AND TYPE_OF_TRANSACTION<>'D'"
    ElseIf SaveMoneyTransaction.Type_Of_Account = "SS" Then
        sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='SS' AND TYPE_OF_TRANSACTION<>'D' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='SS' And Delete_Flag='N'AND TYPE_OF_TRANSACTION<>'D'"
                
    ElseIf SaveMoneyTransaction.Type_Of_Account = "CC" Then
        sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='CC' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='CC' And Delete_Flag='N'"
    ElseIf SaveMoneyTransaction.Type_Of_Account = "CA" Then
        sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='CA' And Delete_Flag='N'And Type_of_Transaction = 'C' AND TYPE_OF_TRANSACTION<>'D'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='CA' And Delete_Flag='N' AND TYPE_OF_TRANSACTION<>'D'"
    ElseIf SaveMoneyTransaction.Type_Of_Account = "CL" Then
        sqlString = "select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_Credit where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='CL' And Delete_Flag='N'And Type_of_Transaction = 'C'" _
                & "Union select Type_Of_Account , Account_No ,Amount,TYPE_OF_TRANSACTION,DC,TRANSACTION_key,FD_No,New_ROI_FD,DATE_OF_TRANSACTION,IS_Approved1,IS_Approved2,IS_Approved3,TRANSACTION_NO From t_Transaction_DEBIT where TO_CHAR(DATE_OF_TRANSACTION,'DD-MON-YYYY') = TO_CHAR( TO_DATE('" & SaveMoneyTransaction.DATE_OF_TRANSACTION & "','DD/MM/YYYY'),'DD-MON-YYYY')" _
                & " AND (Type_Of_Account is not null and Account_No is not null) AND Type_Of_Account='CL' And Delete_Flag='N'"
    
    End If
    With look
        'Provide column header for the LOV
        .AddColumnHeaders "Type of Account", "Account No.", "Amount", "TYPE OF TRANSACTION", "DC", "Transaction Key", "FD No.", "New ROI FD", "DATE OF TRANSACTION", "IS_Approved1", "IS_Approved2", "IS_Approved3", "TRANSACTION NO"
        'Provide actual column for the ListView According to the Stored Procedure"
        .AddDisplayFields "TYPE_OF_ACCOUNT", "ACCOUNT_NO", "Amount", "TYPE_OF_TRANSACTION", "DC", _
         "TRANSACTION_key", "FD_No", "New_ROI_FD", "DATE_OF_TRANSACTION", "IS_Approved1", "IS_Approved2", _
         "IS_Approved3", "TRANSACTION_NO"
        .Connection = objDatasource
        .SetColumnsWidth 900, 1500, 1800, 900, 0, 0, 2000, 0, 1500, 0, 0, 0, 1200
        .ProcedureText = sqlString
        .TotalColumns = 12
        .PopulateList
        If .LOVState = False Then
            SaveMoneyTransaction.Type_Of_Account = .DisplayValueByName("TYPE_OF_ACCOUNT")
            If Not IsNull(.DisplayValueByName("ACCOUNT_NO")) Then
                SaveMoneyTransaction.Account_No = .DisplayValueByName("ACCOUNT_NO")
            End If
            SaveMoneyTransaction.TYPE_OF_TRANSACTION = .DisplayValueByName("TYPE_OF_TRANSACTION")
            SaveMoneyTransaction.D_C = .DisplayValueByName("DC")
            SaveMoneyTransaction.TRANSACTION_key = (.DisplayValueByName("TRANSACTION_key"))
            SaveMoneyTransaction.FD_NO = .DisplayValueByName("FD_No")
            SaveMoneyTransaction.New_ROI_FD = .DisplayValueByName("New_ROI_FD")
            SaveMoneyTransaction.DATE_OF_TRANSACTION = .DisplayValueByName("DATE_OF_TRANSACTION")
            SaveMoneyTransaction.Approved_By1 = .DisplayValueByName("IS_Approved1")
            SaveMoneyTransaction.Approved_By2 = .DisplayValueByName("IS_Approved2")
            SaveMoneyTransaction.Approved_By3 = .DisplayValueByName("IS_Approved3")
            SaveMoneyTransaction.TRANSACTION_No = .DisplayValueByName("TRANSACTION_NO")
            PopupLov = True
            Exit Function
        End If
        PopupLov = False
    End With
End Function
Public Function GETNetBalance(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
            If SaveMoneyTransaction.Type_Of_Account = "RD" Then
                .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceRD"
            ElseIf SaveMoneyTransaction.Type_Of_Account = "RL" Then
                .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceRL"
            Else
                .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalance"
            End If
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            'Execute to fetch records
            Set GETNetBalance = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GETInstallmentRD(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
                .CommandText = "KccBank.PACK_Money_Transaction.GETInstallmentRD"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            'Execute to fetch records
            Set GETInstallmentRD = .Execute
         End With
    Set objCommand = Nothing
End Function

Public Function GETNetBalanceFD(ByRef objDatasource As ADODB.Connection) _
                                As Double
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
                .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceFD"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param2", adVarChar, adParamOutput, 200)
            .Parameters.Append .CreateParameter("Param3", adDate, adParamOutput, 8)
            .Parameters.Append .CreateParameter("Param4", adNumeric, adParamOutput, 16)
            .Parameters.Append .CreateParameter("Param5", adNumeric, adParamOutput, 16)
            .Parameters.Append .CreateParameter("Param6", adDate, adParamOutput, 8)
            .Parameters.Append .CreateParameter("Param7", adChar, adParamOutput, 1)
            .Parameters.Append .CreateParameter("Param8", adNumeric, adParamOutput, 5)
            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param5").Value) And objCommand("Param5").Value <> -1 Then
                GETNetBalanceFD = objCommand("Param5").Value
                If Not IsNull(objCommand("param3").Value) Then
                    SaveMoneyTransaction.InsertModifyDate = objCommand("param3").Value
                End If
                If Not IsNull(objCommand("param2").Value) Then
                    SaveMoneyTransaction.Address = objCommand("param2").Value
                End If
                If Not IsNull(objCommand("param1").Value) Then
                        SaveMoneyTransaction.name = objCommand("param1").Value
                End If
                If Not IsNull(objCommand("param4").Value) Then
                        SaveMoneyTransaction.Maturity_Amt = objCommand("param4").Value
                End If
                If Not IsNull(objCommand("param6").Value) Then
                        SaveMoneyTransaction.Fd_Date = objCommand("param6").Value
                End If
                If Not IsNull(objCommand("param8").Value) Then
                        SaveMoneyTransaction.ROI = objCommand("param8").Value
                End If
            Else
                GETNetBalanceFD = 0
            End If
                If Not IsNull(objCommand("param7").Value) Then
                    If objCommand("param7").Value = "Y" Then
                        GETNetBalanceFD = -1
                    End If
                End If
            Set objCommand = Nothing
End Function
Public Function GETNetBalanceCL(ByRef objDatasource As ADODB.Connection) _
                                As Double
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
                .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceCL"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param2", adVarChar, adParamOutput, 200)
            .Parameters.Append .CreateParameter("Param3", adChar, adParamOutput, 1)
            .Parameters.Append .CreateParameter("Param4", adChar, adParamOutput, 1)
            .Parameters.Append .CreateParameter("Param5", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param6", adNumeric, adParamOutput, 15)
            .Parameters.Append .CreateParameter("Param7", adDate, adParamOutput, 8)
            .Parameters.Append .CreateParameter("Param8", adNumeric, adParamOutput, 3)
            .Parameters.Append .CreateParameter("Param9", adNumeric, adParamOutput, 2)
            .Parameters.Append .CreateParameter("Param10", adNumeric, adParamOutput, 15)

            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param6").Value) And objCommand("Param6").Value <> -1 Then
                GETNetBalanceCL = objCommand("Param6").Value
                If Not IsNull(objCommand("param1").Value) Then
                    SaveMoneyTransaction.name = objCommand("param1").Value
                End If
                If Not IsNull(objCommand("param2").Value) Then
                    SaveMoneyTransaction.Address = objCommand("param2").Value
                End If
                If Not IsNull(objCommand("param3").Value) Then
                        SaveMoneyTransaction.Closed_Operative = objCommand("param3").Value
                End If
                If Not IsNull(objCommand("param4").Value) Then
                        SaveMoneyTransaction.Severally_Joint = objCommand("param4").Value
                End If
                If Not IsNull(objCommand("param5").Value) Then
                        SaveMoneyTransaction.Item = objCommand("param5").Value
                End If
                If Not IsNull(objCommand("param6").Value) Then
                        SaveMoneyTransaction.Amount = objCommand("param6").Value
                End If
                If Not IsNull(objCommand("param7").Value) Then
                        SaveMoneyTransaction.Opening_Date = objCommand("param7").Value
                End If
                
                If Not IsNull(objCommand("param8").Value) Then
                        SaveMoneyTransaction.Period = objCommand("param8").Value
                End If
                If Not IsNull(objCommand("param9").Value) Then
                        SaveMoneyTransaction.ROI = objCommand("param9").Value
                End If
                If Not IsNull(objCommand("param10").Value) Then
                        SaveMoneyTransaction.Installment = objCommand("param10").Value
                End If
            Else
                GETNetBalanceCL = -1
            End If
'                If Not IsNull(objCommand("param7").Value) Then
'                    If objCommand("param7").Value = "Y" Then
'                        GETNetBalanceCL = -1
'                    End If
'                End If
            Set objCommand = Nothing
End Function

Public Function GETCustomerBalance(ByRef objDatasource As ADODB.Connection)
                            
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GETCustomerBalance"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            If SaveMoneyTransaction.Type_Of_Account = "FD" Then
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, "FL")
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_NoTo))
            Else
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            End If
            If SaveMoneyTransaction.FD_NO = Empty Then
                SaveMoneyTransaction.FD_NO = 0
             End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            .Parameters.Append .CreateParameter("Param2", adNumeric, adParamOutput)
            .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput)
            .Execute
            End With
            If objCommand("Param1").Value = -1 Then
                SaveMoneyTransaction.OpeningBalance = 0
                SaveMoneyTransaction.TOTDb = 0
                SaveMoneyTransaction.TOTCr = 0
                Set objCommand = Nothing
                Exit Function
            End If
            If Not IsNull(objCommand("Param1").Value) Then
                SaveMoneyTransaction.OpeningBalance = objCommand("Param1").Value
             Else
                SaveMoneyTransaction.OpeningBalance = 0
            End If
            If Not IsNull(objCommand("Param2").Value) Then
                SaveMoneyTransaction.TOTDb = objCommand("Param2").Value
            Else
                SaveMoneyTransaction.TOTDb = 0
            End If
            If Not IsNull(objCommand("Param3").Value) Then
                SaveMoneyTransaction.TOTCr = objCommand("Param3").Value
            Else
                SaveMoneyTransaction.TOTCr = 0
            End If
    Set objCommand = Nothing
    End Function

Public Function GETNetBalanceFL(ByRef objDatasource As ADODB.Connection) _
                                As Double
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
                .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceFL"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            '.Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param2", adVarChar, adParamOutput, 200)
            .Parameters.Append .CreateParameter("Param4", adDate, adParamOutput)
            .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput, 16)
            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param3").Value) And objCommand("Param3").Value <> -1 Then
                GETNetBalanceFL = objCommand("Param3").Value
                If Not IsNull(objCommand("Param1").Value) Then
                    SaveMoneyTransaction.name = objCommand("Param1").Value
                End If
                If Not IsNull(objCommand("Param2").Value) Then
                    SaveMoneyTransaction.Address = objCommand("Param2").Value
                End If
                If Not IsNull(objCommand("Param4").Value) Then
                    SaveMoneyTransaction.Fd_Date = objCommand("Param4").Value
                End If

            Else
                GETNetBalanceFL = 0
            End If
    Set objCommand = Nothing
End Function

Public Function GETNetBalanceRL(ByRef objDatasource As ADODB.Connection) _
                                As Double
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceRL"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param2", adVarChar, adParamOutput, 200)
            .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput, 16)
            .Parameters.Append .CreateParameter("Param4", adDate, adParamOutput)
            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param3").Value) And objCommand("Param2").Value <> -1 Then
                GETNetBalanceRL = objCommand("Param3").Value
                If Not IsNull(objCommand("Param1").Value) Then
                    SaveMoneyTransaction.name = objCommand("Param1").Value
                End If
                If Not IsNull(objCommand("Param2").Value) Then
                    SaveMoneyTransaction.Address = objCommand("Param2").Value
                End If
                If Not IsNull(objCommand("Param4").Value) Then
                    SaveMoneyTransaction.Fd_Date = objCommand("Param4").Value
                End If
                
            Else
                GETNetBalanceRL = 0
            End If
    Set objCommand = Nothing
End Function
Public Function GETNetBalanceNS(ByRef objDatasource As ADODB.Connection) _
                                As Double
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GETNetBalanceNS"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param2", adVarChar, adParamOutput, 200)
            .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput, 16)
            .Parameters.Append .CreateParameter("Param4", adDate, adParamOutput)
            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param3").Value) And objCommand("Param3").Value <> -1 Then
                GETNetBalanceNS = objCommand("Param3").Value
                If Not IsNull(objCommand("Param1").Value) Then
                    SaveMoneyTransaction.name = objCommand("Param1").Value
                End If
                If Not IsNull(objCommand("Param2").Value) Then
                    SaveMoneyTransaction.Address = objCommand("Param2").Value
                End If
                If Not IsNull(objCommand("Param4").Value) Then
                    SaveMoneyTransaction.Fd_Date = objCommand("Param4").Value
                End If
                
            Else
                GETNetBalanceNS = 0
            End If
    Set objCommand = Nothing
End Function

Public Function GetData(ByRef objDatasource As ADODB.Connection, ByVal dowork As Byte) _
                                As ADODB.Recordset
    'Get records from debit and credit tables
    Dim objCommand As New ADODB.Command
        With objCommand
            If dowork = 1 Then
                .CommandText = "KccBank.PACK_Money_Transaction.GetDataDebit"
            ElseIf dowork = 2 Then
                .CommandText = "KccBank.PACK_Money_Transaction.GetDataCredit"
            End If
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.TRANSACTION_key)
            'Execute to fetch records
            Set GetData = .Execute
        End With
End Function
Public Function GETAccNo(ByRef objDatasource As ADODB.Connection) _
                                As Long
    'Get Max Account No.
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GETAccNo"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param1").Value) Then
                GETAccNo = objCommand("Param1").Value
            End If
         End With
    Set objCommand = Nothing
End Function
Public Function GetInterestPaid(ByRef objDatasource As ADODB.Connection, ByVal dowork As Byte) _
                                As Double
    'Get Max Account No.
    Dim objCommand As New ADODB.Command
        With objCommand
            If dowork = 1 Then
                .CommandText = "KccBank.PACK_Money_Transaction.GetInterestPaidFD"
            Else
                .CommandText = "KccBank.PACK_Money_Transaction.GetInterestPaidRD"
            End If
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param1").Value) And objCommand("Param1").Value <> -1 Then
                GetInterestPaid = objCommand("Param1").Value
            Else
                GetInterestPaid = 0
            End If
         End With
    Set objCommand = Nothing
End Function
Public Function GetLoanFD(ByRef objDatasource As ADODB.Connection) _
                                As Double
    'Get Max Account No.
    Dim objCommand As New ADODB.Command
    Dim accountNo As Long
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GetLoanFD"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adChar, adParamOutput, 2)
            .Parameters.Append .CreateParameter("Param2", adNumeric, adParamOutput)
            .Parameters.Append .CreateParameter("Param3", adVarChar, adParamOutput, 100)
            .Parameters.Append .CreateParameter("Param4", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param4").Value) And objCommand("Param4").Value <> -1 Then
                GetLoanFD = objCommand("Param4").Value
                If Not IsNull(objCommand("Param1").Value) Then
                    SaveMoneyTransaction.Type_Of_AccountTo = objCommand("Param1").Value
                End If
                If Not IsNull(objCommand("Param2").Value) Then
                    SaveMoneyTransaction.Account_NoTo = objCommand("Param2").Value
                End If
                If Not IsNull(objCommand("Param3").Value) Then
                    SaveMoneyTransaction.name = objCommand("Param3").Value
                End If
            Else
                GetLoanFD = 0
            End If
         End With
    Set objCommand = Nothing
End Function
Public Function GetLoanRD(ByRef objDatasource As ADODB.Connection) _
                                As Double
    'Get Max Account No.
    Dim objCommand As New ADODB.Command
    Dim accountNo As Long
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GetLoanRD"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.Account_No)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            .Parameters.Append .CreateParameter("Param2", adNumeric, adParamOutput)
            .Parameters.Append .CreateParameter("Param3", adNumeric, adParamOutput)
            .Parameters.Append .CreateParameter("Param4", adDate, adParamOutput)
            .Parameters.Append .CreateParameter("Param5", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param5").Value) And objCommand("Param5").Value <> -1 Then
                GetLoanRD = objCommand("Param2").Value
            Else
                GetLoanRD = 0
            End If
            If Not IsNull(objCommand("Param1").Value) Then
                SaveMoneyTransaction.Account_NoTo = objCommand("Param1").Value
                SaveMoneyTransaction.Type_Of_AccountTo = "RL"
            End If
         End With
    Set objCommand = Nothing
End Function
Public Function GetFDInterest(ByRef objDatasource As ADODB.Connection) _
                                As Double
    'Get Max Account No.
    Dim objCommand As New ADODB.Command
    Dim accountNo As Long
        With objCommand
            .CommandText = "KccBank.PACK_Money_Transaction.GetFDInterest"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param1").Value) Then
                GetFDInterest = objCommand("Param1").Value
            Else
                GetFDInterest = 0
            End If
         End With
    Set objCommand = Nothing
End Function

Public Function GETStandingInstructions(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    'Get Max Account No.
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = " KccBank.PACK_Money_Transaction.GETStandingInstructions"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.FD_NO)
            .Parameters.Append .CreateParameter("Param1", adChar, adParamOutput, 2)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param1").Value) And objCommand("Param1").Value <> "N" Then
                GETStandingInstructions = False
            Else
                GETStandingInstructions = True
            End If
         End With
    Set objCommand = Nothing
End Function

Public Function GetACHolderName(ByRef objDatasource As ADODB.Connection) _
                                As String
    'Get Account Holder's Name
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = " KccBank.PACK_Money_Transaction.GetACHolderName"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.Account_No))
            .Parameters.Append .CreateParameter("Param1", adVarChar, adParamOutput, 100)
            'Execute to fetch records
             .Execute
            If Not IsNull(objCommand("Param1").Value) And objCommand("Param1").Value <> -1 Then
                GetACHolderName = objCommand("Param1").Value
            Else
                GetACHolderName = Empty
            End If
         End With
    Set objCommand = Nothing
End Function
