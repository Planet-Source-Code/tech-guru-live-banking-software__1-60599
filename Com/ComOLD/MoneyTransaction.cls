VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type MoneyTransaction
    TRANSACTION_No          As Long
    Type_Of_Account         As String
    Account_No              As Long
    TYOE_OF_TRANSACTION     As String
    PAYMENT_WITHDRAWL       As String
    DATE_OF_TRANSACTION     As Date
    CHECK_NO                As Long
    CHECK_DATE              As Date
    TOKEN_NO                As Long
    SCROLL_NO               As Long
    Remarks                 As String
    CLEARING_FLAG           As String
    DELETE_FLAG             As String
    TerminalName            As String
    UserName                As String
    InsertModifyDate        As Date
End Type
Public Enum dbState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum
Private SaveMoneyTransaction     As MoneyTransaction
Public Property Let TRANSACTION_No(ByVal v_Value As Long)
    SaveMoneyTransaction.TRANSACTION_No = v_Value
End Property

Public Property Let Type_Of_Account(ByVal v_Value As String)
    SaveMoneyTransaction.Type_Of_Account = v_Value
End Property
Public Property Let Account_No(ByVal v_Value As Long)
    SaveMoneyTransaction.AccountNo = v_Value
End Property
Public Property Let Type_Of_Transaction(ByVal v_Value As String)
    SaveMoneyTransaction.TYOE_OF_TRANSACTION = v_Value
End Property
Public Property Let PAYMENT_WITHDRAWL(ByVal v_Value As String)
    SaveMoneyTransaction.PAYMENT_WITHDRAWL = v_Value
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

Public Property Let CLEARING_FLAG(ByVal v_Value As String)
    SaveMoneyTransaction.CLEARING_FLAG = v_Value
End Property
Public Property Let DELETE_FLAG(ByVal v_Value As String)
    SaveMoneyTransaction.DELETE_FLAG = v_Value
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
Public Property Get TRANSACTION_No() As Long
    TRANSACTION_No = SaveMoneyTransaction.TRANSACTION_No
End Property

Public Property Get Type_Of_Account() As String
    Type_Of_Account = SaveMoneyTransaction.Type_Of_Account
End Property
Public Property Get Account_No() As Long
    AccountNo = SaveMoneyTransaction.AccountNo
End Property
Public Property Get Type_Of_Transaction() As String
    TYOE_OF_TRANSACTION = SaveMoneyTransaction.TYOE_OF_TRANSACTION
End Property
Public Property Get PAYMENT_WITHDRAWL() As String
    PAYMENT_WITHDRAWL = SaveMoneyTransaction.PAYMENT_WITHDRAWL
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
Public Property Get CLEARING_FLAG() As String
    CLEARING_FLAG = SaveMoneyTransaction.CLEARING_FLAG
End Property
Public Property Get DELETE_FLAG() As String
    DELETE_FLAG = SaveMoneyTransaction.DELETE_FLAG
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
Public Sub INSERTModifyDATA(ByRef objDatasource As ADODB.Connection, ByVal doWork As Byte)
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                If doWork = 1 Then
                    .CommandText = "kccBank.PACK_Account_Opening.INSERTDATA"
                Else
                    .CommandText = "kccBank.PACK_Account_Opening.ModifyData"
                End If
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.AccountNo)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.NameOfAcHolder)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.FatherName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Address)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveMoneyTransaction.PhNo)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.AdultMinor)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.Dob)
                    If SaveMoneyTransaction.Guardian = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.Guardian)
                    End If
                    If SaveMoneyTransaction.DobGur = Empty Then
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DobGur)
                    End If
                    If SaveMoneyTransaction.RelationGur = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveMoneyTransaction.RelationGur)
                    End If
                    If SaveMoneyTransaction.Nominee = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.Nominee)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.Nominee)
                    End If
                    If SaveMoneyTransaction.DOBNom = Empty Then
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.DOBNom)
                    End If
                    If SaveMoneyTransaction.AddressNom = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.AddressNom)
                    End If
                    If SaveMoneyTransaction.RelationNom = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveMoneyTransaction.RelationNom)
                    End If
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.SingleJoint)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.RateOfInterestC)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.RateOfInterestD)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveMoneyTransaction.Occupation)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.StaffPublic)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 9, SaveMoneyTransaction.IntroducedBy)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.CanIntroduce)
                    If SaveMoneyTransaction.ReasonIntro = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveMoneyTransaction.ReasonIntro)
                    End If
                    If SaveMoneyTransaction.Pan = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveMoneyTransaction.Pan)
                    End If
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.ClosedOperative)
                    If SaveMoneyTransaction.ReasonClosed = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveMoneyTransaction.ReasonClosed)
                    End If
                    If SaveMoneyTransaction.Remarks = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveMoneyTransaction.Remarks)
                    End If
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveMoneyTransaction.LedgerNo)
                    If SaveMoneyTransaction.Severally_Joint = Empty Then
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.Severally_Joint)
                    End If
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
                    .Execute
                End With
                Set objCommand = Nothing
              End Sub
Public Sub INSERTModifyJointData(ByRef objDatasource As ADODB.Connection, ByVal doWork As Byte)
    Dim objCommand As ADODB.Command
    Dim i As Integer
            For i = 0 To UBound(SaveMoneyTransaction.JointData) - 1 Step 4
            Set objCommand = New ADODB.Command
            With objCommand
                If SaveMoneyTransaction.JointData(i + 3) = "i" Then
                    .CommandText = "kccBank.PACK_Account_Opening.INSERTJointDATA"
                Else
                    .CommandText = "kccBank.PACK_Account_Opening.ModifyJointData"
                End If
                .CommandType = adCmdStoredProc
                'Send connection
                .ActiveConnection = objDatasource
                'Set command time out
                .CommandTimeout = 0
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.AccountNo))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, (SaveMoneyTransaction.JointData(i + 1)))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, CInt(SaveMoneyTransaction.JointData(i)))
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveMoneyTransaction.JointData(i + 2))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
                .Execute
            End With
            Set objCommand = Nothing
        Next
End Sub
Public Function CheckAccount(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.CheckAccount"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveMoneyTransaction.AccountNo)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
         .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
                CheckAccount = True
            Else
                CheckAccount = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function DeleteDATA(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteData"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.AccountNo))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveMoneyTransaction.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveMoneyTransaction.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveMoneyTransaction.InsertModifyDate)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        DeleteDATA = True
                    Else
                        DeleteDATA = False
                    End If
                End If
        End With
End Function
Public Function PopupLov(ByRef objDatasource As ADODB.Connection) As Boolean
    'Get the fields from t_emp_mst table Using LOV
    'Make the object of LOV
    Dim look As New LookUp
    Dim strProcedure As String
    Dim sql As String
    sql = "Select TYPE_OF_ACCOUNT, ACCOUNT_NO, NAME_OF_AC_HOLDER, ADDRESS, PHONE_NO from kccbank.T_Customer_detail where Delete_Flag=" & "'N'"
   
    'strProcedure = "{KccBank.PACK_Account_Opening.GETACCOUNTNOS()}"
    With look
        'Provide column header for the LOV
        .AddColumnHeaders "Type of Account", "Account No.", "Name Of A/C Holer", "Address", "Ph. No."
        'Provide actual column for the ListView According to the Stored Procedure"
        .AddDisplayFields "TYPE_OF_ACCOUNT", "ACCOUNT_NO", "NAME_OF_AC_HOLDER", "ADDRESS", "PHONE_NO"
        .Connection = objDatasource
        .SetColumnsWidth 1200, 700, 2000, 2500, 1000
        .ProcedureText = sql
        .TotalColumns = 5
        .PopulateList
        If .LOVState = False Then
            SaveMoneyTransaction.TypeOfAccount = (.DisplayValueByName("TYPE_OF_ACCOUNT"))
            SaveMoneyTransaction.AccountNo = CLng(.DisplayValueByName("ACCOUNT_NO"))
            PopupLov = True
            Exit Function
         
         End If
     PopupLov = False
    End With
End Function
Public Function GetACCOUNTDETAIL(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
                                
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Account_Opening.GetACCOUNTDETAIL"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.AccountNo))
            'Execute to fetch records
            Set GetACCOUNTDETAIL = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GETJOINTDATA(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Account_Opening.GETJOINTDATA"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.AccountNo))
            'Execute to fetch records
            Set GETJOINTDATA = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GetOccupation(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "{call KccBank.PACK_ACCOUNT_OPENING.GetOccupation()}"
            .CommandType = adCmdText
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            'Execute to fetch records
            Set GetOccupation = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GETAccNo(ByRef objDatasource As ADODB.Connection) _
                                As Long
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = " KccBank.PACK_ACCOUNT_OPENING.GETAccNo"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    GETAccNo = objCommand("Param1").Value
                End If
         End With
    Set objCommand = Nothing
End Function

Public Function DeleteJointSinleRecord(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteJointSinleRecord"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.AccountNo))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, CInt(SaveMoneyTransaction.JointKey))
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        DeleteJointSinleRecord = True
                    Else
                        DeleteJointSinleRecord = False
                    End If
                End If
        End With
End Function
Public Function DeleteJointAllRecords(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteJointAllRecords"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveMoneyTransaction.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveMoneyTransaction.AccountNo))
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        DeleteJointAllRecords = True
                    Else
                        DeleteJointAllRecords = False
                    End If
                End If
        End With
End Function



