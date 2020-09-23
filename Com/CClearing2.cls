VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClearing2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       CLASS      :CClearing
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :HARISH CHAUHAN
'       MADE DATE  :15-04-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
 
  Private Type GroupData
        ClearingKey                As Long
        TypeOfAccount              As String
        AccountNo                  As Long
        ClearingDate               As Date
        ChequeNo                   As Long
        ChequeDate                 As Date
        BankKey                    As Long
        ChequeAmount               As Double
        ReturnBack                 As String
        Remarks                    As String
        m_blnInCorrectRecord       As Boolean
        TerminalName               As String
        UserName                   As String
        InsertDate                 As Date
        ModifyDate                 As Date
 End Type
        Public Enum DBState
        DBInsert = 1
        DBModify = 2
        DBDelete = 3
    End Enum
    Private DBData                  As GroupData
    
        Public Function GetClearingList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
'                                 ByVal hwnd As Long)_
'                                 As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
'        Dim objOLEServer As New KccbUtils.OLERemoteServer
        With objCommand
             .CommandText = "PACK_CLEARING_MST.GETCLEARINGDATA()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
             Set .ActiveConnection = objDatasource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetClearingList = objRecordset

        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing

End Function
Public Function GetAccountType(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "{call KccBank.PACK_CLEARING_MST.GetAccountType()}"
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


    Public Property Let ClearingKey(ByVal v_Value As Long)
        DBData.ClearingKey = v_Value
    End Property
    Public Property Let Remarks(ByVal v_Value As String)
        DBData.Remarks = v_Value
    End Property

    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property

    Public Property Let AccountNo(ByVal v_Value As Long)
       DBData.AccountNo = v_Value
    End Property

    Public Property Let ClearingDate(ByVal v_Value As Date)
        DBData.ClearingDate = v_Value
    End Property

    Public Property Let ChequeNo(ByVal v_Value As Long)
        DBData.ChequeNo = v_Value
    End Property
    Public Property Let ChequeDate(ByVal v_Value As Date)
        DBData.ChequeDate = v_Value
    End Property
    Public Property Let BankKey(ByVal v_Value As Long)
        DBData.BankKey = v_Value
    End Property
       
    Public Property Let ChequeAmount(ByVal v_Value As Double)
        DBData.ChequeAmount = v_Value
    End Property
   
    Public Property Let ReturnBack(ByVal v_Value As String)
        DBData.ReturnBack = v_Value
    End Property

    Public Property Let TerminalName(ByVal v_Value As String)
        DBData.TerminalName = v_Value
    End Property

    Public Property Let UserName(ByVal v_Value As String)
        DBData.UserName = v_Value
    End Property

    Public Property Let InsertDate(ByVal v_Value As Date)
        DBData.InsertDate = v_Value
    End Property

    Public Property Let ModifyDate(ByVal v_Value As Date)
        DBData.ModifyDate = v_Value
    End Property

    Public Property Get InCorrectRecord() As Boolean
        IsRecordExists = DBData.m_blnInCorrectRecord
    End Property
    Public Property Let InCorrectRecord(ByVal v_Value As Boolean)
        DBData.m_blnInCorrectRecord = v_Value
    End Property
    
    
    Public Function GetBankList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
'                                 ByVal hwnd As Long)_
'                                 As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
'        Dim objOLEServer As New KccbUtils.OLERemoteServer
        With objCommand
             .CommandText = "PACK_BANK_MASTER.GETBANKDETAILS()"
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = objDatasource
             .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetBankList = objRecordset

        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function


Function GetCustomerName(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
        'This Procedure gets name and Balance of the Account holder
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_CLEARING_MST.GET_CUSTOMER_NAME()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

                Set .ActiveConnection = objDatasource
                .CommandTimeout = 0

                Set objRecordset = .Execute
                Set GetCustomerName = objRecordset
           End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function


Public Sub SaveData(ByRef objDatasource As ADODB.Connection, _
                        ByVal Recordstate As DBState)
    'This Procedure Inserts Or Modify data depending Upon the Mode
    Dim objCommand As New ADODB.Command
    On Error GoTo InsertError:
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_CLEARING_MST.INSERT_CLEARINGDATA"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_CLEARING_MST.MODIFY_CLEARINGDATA"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_CLEARING_MST.DELETE_CLEARINGDATA"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                If DBData.AccountNo = 0 Or DBData.AccountNo = Empty Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, Null)
                Else
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                End If
                If DBData.ChequeNo = 0 Or DBData.ChequeNo = Empty Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, Null)
                Else
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.ChequeNo)
                End If
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ChequeDate)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.BankKey)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, DBData.ChequeAmount)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.Remarks)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
       
            End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                If DBData.AccountNo = 0 Or DBData.AccountNo = Empty Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, Null)
                Else
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                End If
                If DBData.ChequeNo = 0 Or DBData.ChequeNo = Empty Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, Null)
                Else
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.ChequeNo)
                End If

                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ChequeDate)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.BankKey)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, DBData.ChequeAmount)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.Remarks)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ClearingKey)
            End If
            If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ClearingKey)
            End If
                .Execute
    End With
            Set objCommand = Nothing

InsertError:

    Screen.MousePointer = vbDefault
    Dim errNumber As String
    Dim errString As String
    errString = Left(Err.Description, 9)
    errNumber = Right(errString, 5)

     If errNumber = "02291" Then
        MsgBox "Either Type of Account or Account No is Incorrect", vbInformation, "Bank Error"
        InCorrectRecord = True
     Else
        InCorrectRecord = False
    End If

End Sub


        Function VerifyAccountNo(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
            .CommandText = "PACK_CLEARING_MST.VERIFY_ACCOUNTNO()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set VerifyAccountNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

