VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUsedChequeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type UsedCheque
    TrnKey As Integer
    TypeOfAccount As String
    AccountNo As Integer
    DateOfEntry As Date
    ChequeNo As Integer
    Status As String
    BankCharges As Double
    IsDeleted As String
    TerminalName As String
    UserName As String
    TrnDate As Date
    IsChequeNoExists As Boolean
End Type

Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum
Private m_blnSuccess As Boolean
Private DBData As UsedCheque
Public Property Let TrnKey(ByVal v_value As Integer)
    DBData.TrnKey = v_value
End Property
Public Property Let TypeOfAccount(ByVal v_value As String)
    DBData.TypeOfAccount = v_value
End Property
Public Property Let AccountNo(ByVal v_value As Integer)
    DBData.AccountNo = v_value
End Property
Public Property Let DateOfEntry(ByVal v_value As Date)
    DBData.DateOfEntry = v_value
End Property
Public Property Let ChequeNo(ByVal v_value As Integer)
    DBData.ChequeNo = v_value
End Property
Public Property Let BankCharges(ByVal v_value As Double)
    DBData.BankCharges = v_value
End Property
Public Property Let IsDeleted(ByVal v_value As String)
    DBData.IsDeleted = v_value
End Property
Public Property Get IsChequeNoExists() As Boolean
    IsChequeNoExists = m_blnSuccess
End Property
Public Property Let TerminalName(ByVal v_value As String)
    DBData.TerminalName = v_value
End Property
Public Property Let UserName(ByVal v_value As String)
    DBData.UserName = v_value
End Property
Public Property Let TrnDate(ByVal v_value As Date)
    DBData.TrnDate = v_value
End Property
Public Property Let Status(ByVal v_value As String)
    DBData.Status = v_value
End Property
Public Function GetUsedChequeDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.GetUsedChequeDetails"
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        Set GetUsedChequeDetails = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Function CheckChequeNo(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.CheckChequeNo"
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNo)
        Set CheckChequeNo = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Sub ChequeNoExists(ByRef objDatasource As ADODB.Connection)
    Dim objCmd As New ADODB.Command
    Dim objRS As New ADODB.Recordset
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.IsChequeNoExists"
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNo)
        Set objRS = .Execute
    End With
    
    With objRS
        If .RecordCount <> 0 Then
            m_blnSuccess = True
        Else
            m_blnSuccess = False
        End If
    End With
    
    Set objCmd = Nothing
    Set objRS = Nothing
End Sub
Public Sub SaveData(ByRef objDatasource As ADODB.Connection, ByVal Recordstate As DBState)
    Dim objCmd As New ADODB.Command
    
    With objCmd
        
        If Recordstate = DBInsert Then
            .CommandText = "Pack_UsedCheque_Master.Insertdata"
        ElseIf Recordstate = DBModify Then
            .CommandText = "Pack_UsedCheque_Master.Modifydata"
        ElseIf Recordstate = DBDelete Then
            .CommandText = "Pack_UsedCheque_Master.DeleteData"
        End If
        
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .CommandType = adCmdStoredProc
        
        If Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.TrnKey)
        End If
        If Recordstate = DBInsert Or Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        End If
        If Recordstate = DBInsert Or Recordstate = DBModify Then
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.DateOfEntry)
        End If
        If Recordstate = DBInsert Or Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNo)
        End If
        If Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, DBData.IsDeleted)
        End If
        If Recordstate = DBInsert Or Recordstate = DBModify Then
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, DBData.Status)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.BankCharges)
        End If
        If Recordstate = DBInsert Or Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.TrnDate)
        End If
        
        .Execute
    End With
    
    Set objCmd = Nothing
End Sub
Public Function CheckAccountNo(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.CheckAccountNo"
        Set .ActiveConnection = objDatasource
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        Set CheckAccountNo = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Function GetAccountCode(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.GetAccountCode()"
        Set .ActiveConnection = objDatasource
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set GetAccountCode = .Execute
    End With
    
    Set objCmd = Nothing
End Function
