VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccountClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type AccountClosing
    TrnNo As Integer
    TypeOfAccount As String
    AccountNo As Integer
    DateOfClosing As String
    TerminalName As String
    UserName As String
    ClosingDate As Date
    IsAccountClose As Integer
    NoOfIssuedCheques As Integer
    NoOfUsedCheques As Integer
End Type

Private DBData As AccountClosing
Private intAccountClose As Integer
Private intNoOfIssuedCheques As Integer
Private intNoOfUsedCheques As Integer
Public Property Let TrnNo(ByVal v_Value As Integer)
    DBData.TrnNo = v_Value
End Property
Public Property Let TypeOfAccount(ByVal v_Value As String)
    DBData.TypeOfAccount = v_Value
End Property
Public Property Let AccountNo(ByVal v_Value As Integer)
    DBData.AccountNo = v_Value
End Property
Public Property Let DateOfClosing(ByVal v_Value As Date)
    DBData.DateOfClosing = v_Value
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    DBData.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    DBData.UserName = v_Value
End Property
Public Property Let ClosingDate(ByVal v_Value As Date)
    DBData.ClosingDate = v_Value
End Property
Public Property Get IsAccountClose() As Integer
    DBData.IsAccountClose = intAccountClose
End Property
Public Property Get NoOfIssuedCheques() As Integer
    DBData.NoOfIssuedCheques = intNoOfIssuedCheques
End Property
Public Property Get NoOfUsedCheques() As Integer
    DBData.NoOfUsedCheques = intNoOfUsedCheques
End Property

Public Function GetAccountCode(ByRef objdatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.GetAccountCode()"
        Set .ActiveConnection = objdatasource
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set GetAccountCode = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Function CheckAccountNo(ByRef objdatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_UsedCheque_Master.CheckAccountNo"
        Set .ActiveConnection = objdatasource
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        Set CheckAccountNo = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Function CheckStandingInstrs(ByRef objdatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_Account_Closing.CheckStandingInstrs"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = objdatasource
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        Set CheckStandingInstrs = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Sub CheckIssuedCheques(ByRef objdatasource As ADODB.Connection)
    Dim objCmd As New ADODB.Command
    Dim param1 As New ADODB.Parameter
    Dim param2 As New ADODB.Parameter
    Dim param3 As New ADODB.Parameter
    With objCmd
        .CommandText = "Pack_Account_Closing.CheckIssuedCheques"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = objdatasource
        Set param1 = .CreateParameter("param1", adChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append param1
        Set param2 = .CreateParameter("param2", adNumeric, adParamInput, , DBData.AccountNo)
            .Parameters.Append param2
        Set param3 = .CreateParameter("param3", adNumeric, adParamOutput)
            .Parameters.Append param3
        .Execute
        
        intNoOfIssuedCheques = param3.Value
    End With
    
    Set objCmd = Nothing
    Set param3 = Nothing
End Sub
Public Sub CheckUsedCheques(ByRef objdatasource As ADODB.Connection)
    Dim objCmd As New ADODB.Command
    Dim param1 As New ADODB.Parameter
    Dim param2 As New ADODB.Parameter
    Dim param3 As New ADODB.Parameter
    With objCmd
        .CommandText = "Pack_Account_Closing.CheckUsedCheques"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = objdatasource
        Set param1 = .CreateParameter("param1", adChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append param1
        Set param2 = .CreateParameter("param2", adNumeric, adParamInput, , DBData.AccountNo)
            .Parameters.Append param2
        Set param3 = .CreateParameter("param3", adNumeric, adParamOutput)
            .Parameters.Append param3
        .Execute
        
        intNoOfUsedCheques = param3.Value
    End With
    
    Set objCmd = Nothing
    Set param3 = Nothing
End Sub
Public Function CheckBalance(ByRef objdatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_Account_Closing.CheckBal"
        Set .ActiveConnection = objdatasource
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        Set CheckBalance = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Function CheckClearing(ByRef objdatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As New ADODB.Command
    
    With objCmd
        .CommandText = "Pack_Account_Closing.CheckClearing"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = objdatasource
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, DBData.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
        Set CheckClearing = .Execute
    End With
    
    Set objCmd = Nothing
End Function
Public Sub AccountClose(ByRef objdatasource As ADODB.Connection)
    Dim objCmd As New ADODB.Command
    Dim param1 As New ADODB.Parameter
    Dim param2 As New ADODB.Parameter
    Dim param3 As New ADODB.Parameter
    With objCmd
        .CommandText = "Pack_Account_Closing.AccountClose"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set param1 = .CreateParameter("param1", adChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append param1
        Set param2 = .CreateParameter("param2", adNumeric, adParamInput, , DBData.AccountNo)
            .Parameters.Append param2
        Set param3 = .CreateParameter("param3", adNumeric, adParamOutput)
            .Parameters.Append param3
        .Execute
        intAccountClose = param3.Value
    End With
    
    Set objCmd = Nothing
    Set param3 = Nothing
End Sub
