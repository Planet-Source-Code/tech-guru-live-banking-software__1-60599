VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccLEARING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Private Type GroupData
        ClearingKey                As Long
        TypeOfAccount              As String
        AccountNo                  As Double
        ClearingDate               As Date
        PaymentReceipt             As String
        ChequeNo                   As Double
        ChequeDate                 As Date
        BankNameDrawee             As String
        AccountNoDrawee            As String
        ClReceivedOn               As Variant
        ClRecvdOn                  As Date
        Cleared                    As String
        ChargesNonClearing         As Long
        ChequeAmount               As Double
        ReturnBack                 As String
        TerminalName               As String
        UserName                   As String
        InsertDate                 As Date
        ModifyDate                 As Date
        PaymentOrReceipt           As String
    End Type
        Public Enum DBState
        DBInsert = 1
        DBModify = 2
        DBDelete = 3
    End Enum
    Private DBData                  As GroupData

    Public Function GetClearingList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
'                                 ByVal hwnd As Long)_
'                                 As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
'        Dim objOLEServer As New KccbUtils.OLERemoteServer
        With objCommand
             .CommandText = "PACK_CLEARING_MST.GETCLEARINGDATA()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetClearingList = objRecordset

'            If objRecordset.RecordCount > 0 Then
'                Set GetGroupList = objRecordset
'            Else
'                MsgBox "No Group Exists in database. Quiting!", vbOKOnly + vbCritical, "KCCB Error"
'                objOLEServer.ShutdownRemoteOLEServer (hWnd)
'            End If
        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function
    Public Property Let ClearingKey(ByVal v_Value As Long)
        DBData.ClearingKey = v_Value
    End Property
    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property

    Public Property Let AccountNo(ByVal v_Value As Double)
       DBData.AccountNo = v_Value
    End Property

    Public Property Let ClearingDate(ByVal v_Value As Date)
        DBData.ClearingDate = v_Value
    End Property
    Public Property Let PaymentReceipt(ByVal v_Value As String)
        DBData.PaymentReceipt = v_Value
    End Property
     Public Property Get PaymentReceipt() As String
        PaymentReceipt = DBData.PaymentReceipt
    End Property
     
    Public Property Let ChequeNo(ByVal v_Value As Double)
        DBData.ChequeNo = v_Value
    End Property
    Public Property Let ChequeDate(ByVal v_Value As Date)
        DBData.ChequeDate = v_Value
    End Property
    Public Property Let BankNameDrawee(ByVal v_Value As String)
        DBData.BankNameDrawee = v_Value
    End Property
    Public Property Let AccountNoDrawee(ByVal v_Value As String)
        DBData.AccountNoDrawee = v_Value
    End Property
   Public Property Let ClReceivedOn(ByVal v_Value As Variant)
        DBData.ClReceivedOn = v_Value
    End Property
   Public Property Let ClRecvdOn(ByVal v_Value As Date)
        DBData.ClRecvdOn = v_Value
    End Property
   Public Property Let Cleared(ByVal v_Value As String)
        DBData.Cleared = v_Value
    End Property
   Public Property Let ChargesNonClearing(ByVal v_Value As Long)
        DBData.ChargesNonClearing = v_Value
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

    Public Sub SaveData(ByRef objDataSource As ADODB.Connection, _
                        ByVal Recordstate As DBState)
    Dim objCommand As New ADODB.Command
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_CLEARING_MST.INSERT_CLEARINGDATA"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_CLEARING_MST.MODIFY_CLEARINGDATA"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.PaymentReceipt)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.ChequeNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ChequeDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.BankNameDrawee)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.AccountNoDrawee)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, DBData.ChequeAmount)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.Cleared)
'                If DBData.ClReceivedOn = "" Then
                    
'                Else
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClReceivedOn)
'                End If
'
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, DBData.ChargesNonClearing)
               
                
                
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                
                
                
                 




'                        If DBData.PaymentReceipt = "R" Then
'                                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'                                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'                                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'                        ElseIf DBData.PaymentReceipt = "P" Then
'                                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'                                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'                                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'
'                        End If
            
            End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.PaymentReceipt)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.ChequeNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ChequeDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.BankNameDrawee)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.AccountNoDrawee)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClReceivedOn)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.Cleared)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, DBData.ChargesNonClearing)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, DBData.ChequeAmount)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ClearingKey)

            End If
               .Execute
    End With
            Set objCommand = Nothing
End Sub

Function GetNameAndBalance(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset

        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_CLEARING_MST.GETNAMEANDBALANCE()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetNameAndBalance = objRecordset
        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function


