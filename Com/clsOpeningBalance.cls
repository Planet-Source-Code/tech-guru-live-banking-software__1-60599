VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type StatusMst
    Account_No              As Long
    Type_Of_Account        As String
    Opening_Balance        As Double
    TerminalName            As String
    UserName                As String
    InsertModifyDate        As Date
End Type
Private SaveOpeningBalance     As StatusMst
Property Let Account_No(v_Value As Long)
    SaveOpeningBalance.Account_No = v_Value
End Property
Property Get Account_No() As Long
    Account_No = SaveOpeningBalance.Account_No
End Property
Property Let Type_Of_Account(v_Value As String)
    SaveOpeningBalance.Type_Of_Account = v_Value
End Property
Property Get Type_Of_Account() As String
    Type_Of_Account = SaveOpeningBalance.Type_Of_Account
End Property
Property Let Opening_Balance(v_Value As Double)
    SaveOpeningBalance.Opening_Balance = v_Value
End Property
Property Get Opening_Balance() As Double
    Opening_Balance = SaveOpeningBalance.Opening_Balance
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveOpeningBalance.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveOpeningBalance.UserName = v_Value
End Property
Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveOpeningBalance.InsertModifyDate = v_Value
End Property
Public Property Get TerminalName() As String
    TerminalName = SaveOpeningBalance.TerminalName
End Property
Public Property Get UserName() As String
    UserName = SaveOpeningBalance.UserName
End Property
Public Property Get InsertModifyDate() As Date
 InsertModifyDate = SaveOpeningBalance.InsertModifyDate
End Property
Public Function GETDATA(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "PACK_OPENING_BALANCE.GETDATA()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GETDATA = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Sub InsertData(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "PACK_OPENING_BALANCE.InsertData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveOpeningBalance.Type_Of_Account)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveOpeningBalance.Account_No)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, SaveOpeningBalance.Opening_Balance)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveOpeningBalance.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveOpeningBalance.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveOpeningBalance.InsertModifyDate)
                .Execute
            End With
            Set objCommand = Nothing
End Sub
Public Function ModifyData(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "PACK_OPENING_BALANCE.ModifyData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveOpeningBalance.Type_Of_Account)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveOpeningBalance.Account_No)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, SaveOpeningBalance.Opening_Balance)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveOpeningBalance.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveOpeningBalance.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveOpeningBalance.InsertModifyDate)
                    .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
                     .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        ModifyData = True
                    Else
                        ModifyData = False
                    End If
                End If
            End With
            Set objCommand = Nothing
End Function
Public Function DeleteData(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    'Delete Record from database
    Dim objCommand As New ADODB.Command
    With objCommand
            .CommandText = "PACK_OPENING_BALANCE.DeleteRecord"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveOpeningBalance.Type_Of_Account)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveOpeningBalance.Account_No)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
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
End Function
Public Function CHECKACCOUNTNO(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "PACK_OPENING_BALANCE.CHECKACCOUNTNO"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveOpeningBalance.Type_Of_Account)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveOpeningBalance.Account_No)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        CHECKACCOUNTNO = True
                    Else
                        CHECKACCOUNTNO = False
                    End If
                End If
        End With
End Function
