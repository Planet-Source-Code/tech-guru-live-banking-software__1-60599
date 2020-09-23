VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBankMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type BankMasterdata
    BankKey As Integer
    BankCode As String
    BankName As String
    BankAddress As String
    m_BankKey As Integer
    strBankCode As String
End Type
Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum
Private DBData As BankMasterdata
Private m_blnCheck As Boolean

Public Property Let BankKey(ByVal v_Value As Integer)
    DBData.BankKey = v_Value
End Property
Public Property Let m_BankKey(ByVal v_Value As Integer)
    DBData.m_BankKey = v_Value
End Property
Public Property Let BankCode(ByVal v_Value As String)
    DBData.BankCode = v_Value
End Property
Public Property Let strBankCode(ByVal v_Value As String)
    DBData.strBankCode = v_Value
End Property
Public Property Let BankName(ByVal v_Value As String)
    DBData.BankName = v_Value
End Property
Public Property Let BankAddress(ByVal v_Value As String)
    DBData.BankAddress = v_Value
End Property
Public Function GetBankDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As ADODB.Command
    Set objCmd = New ADODB.Command
        
        With objCmd
            .CommandText = "Pack_Bank_Master.GetBankDetails()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
                .CommandTimeout = 0
            Set GetBankDetails = .Execute
        End With
End Function
Public Sub SaveData(ByRef objDatasource As ADODB.Connection, ByVal Recordstate As DBState)
    Dim param1 As New ADODB.Parameter
    Dim param2 As New ADODB.Parameter
    Dim param3 As New ADODB.Parameter
    Dim param4 As New ADODB.Parameter
    Dim param5 As New ADODB.Parameter
    Dim objCmd As New ADODB.Command
    
    With objCmd
        If Recordstate = 1 Then
            .CommandText = "Pack_Bank_Master.InsertData"
        ElseIf Recordstate = 2 Then
            .CommandText = "Pack_Bank_Master.ModifyData"
        ElseIf Recordstate = 3 Then
            .CommandText = "Pack_Bank_Master.DeleteData"
        End If
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        If Recordstate = 2 Or Recordstate = 3 Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.BankKey)
        End If
        If Recordstate = 1 Or Recordstate = 2 Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.BankCode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.BankName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, DBData.BankAddress)
        End If
        If Recordstate = 3 Then
            .Parameters.Append .CreateParameter("param5", adNumeric, adParamOutput)
        End If
        .Execute
        If Recordstate = 3 Then
            If param5.Value <> 0 Then
                MsgBox "Cannot delete this record, as it was used by some Transaction", vbInformation, "Bank Error"
            End If
        End If
    End With
    Set objCmd = Nothing
    Set param5 = Nothing
End Sub
Public Function CheckBankCode(ByRef objDatasource As ADODB.Connection, ByVal Recordstate As DBState) As Boolean
    Dim param1 As New ADODB.Parameter
    Dim objCmd As New ADODB.Command
    Dim objRS As New ADODB.Recordset
    
    With objCmd
        .CommandText = "Pack_Bank_Master.CheckBankCode"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.BankCode)
        Set objRS = .Execute
    End With
    
    With objRS
        If .RecordCount <> 0 Then
            If Recordstate = 1 Then
                MsgBox "Bank Code already exists, Enter Unique Bank Code", vbInformation, "Bank Error"
                CheckBankCode = False
            ElseIf Recordstate = 2 Then
                If DBData.strBankCode = DBData.BankCode Then
                    CheckBankCode = True
                Else
                    MsgBox "Bank Code already exists, Enter Unique Bank Code", vbInformation, "Bank Error"
                    CheckBankCode = False
                End If
            Else
                CheckBankCode = True
            End If
        Else
            CheckBankCode = True
        End If
    End With
    
    Set objCmd = Nothing
    Set objRS = Nothing
End Function
