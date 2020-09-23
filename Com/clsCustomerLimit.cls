VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomerLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type LimitMst
    AccountNo              As String
    LimitAmount            As Double
    vdate                  As Date
End Type
Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
    DBDebitDelete = 4
    DBCreditDelete = 5
    DBValuesDelete = 6
    End Enum
Private SaveLimitMst     As LimitMst
Property Let AccountNo(V_value As String)
    SaveLimitMst.AccountNo = V_value
End Property
Property Get AccountNo() As String
    AccountNo = SaveLimitMst.AccountNo
End Property
Property Let LimitAmount(V_value As Double)
    SaveLimitMst.LimitAmount = V_value
End Property
Property Get LimitAmount() As Double
    LimitAmount = SaveLimitMst.LimitAmount
End Property
Property Let vdate(V_value As Date)
    SaveLimitMst.vdate = V_value
End Property
Property Get vdate() As Date
    vdate = SaveLimitMst.vdate
End Property
Public Function GETACCNo(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "Natc.Pack_NATC_Limit.GetAccNo()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GETACCNo = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Function GETData(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "Natc.Pack_NATC_Limit.GetData()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GETData = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Sub InsertData(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "Natc.Pack_NATC_Limit.InsertData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, SaveLimitMst.AccountNo)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveLimitMst.LimitAmount)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveLimitMst.vdate)
                    .Execute
            End With
            Set objCommand = Nothing
End Sub
Public Function ModifyData(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "Natc.Pack_NATC_Limit.ModifyData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, SaveLimitMst.AccountNo)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveLimitMst.LimitAmount)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveLimitMst.vdate)
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
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "Natc.Pack_NATC_Limit.DeleteData"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, SaveLimitMst.AccountNo)
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
Public Function CheckAccNo(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "Natc.Pack_NATC_Limit.CheckAccNo"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, SaveLimitMst.AccountNo)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        CheckAccNo = True
                    Else
                        CheckAccNo = False
                    End If
                End If
        End With
End Function


