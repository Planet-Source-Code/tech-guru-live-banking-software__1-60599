VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockHandDp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
'       PROJECT    :KCCRATEOFINTEREST
'       MODULE     :ESTABLISHMENT
'       PRUPOSE    :A genric solution for the stockhand drawing power
'
'       COM        :STOCK HAND DP
'       FORM
'       REFERRED   :STOCK HAND DP
'       FORM CODE  :KCCB EST 0004
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :GURINDER MANN
'       MADE DATE  :21-03-2002
'       REVISIONS  :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------

Private Type StockHandData
    AccountNO           As Double
    MonthDp             As Long
    DateOfSubmission    As Date
    ValueOfStock        As Double
    Margin              As Double
    LimitSanctioned     As Double
    TerminalName        As String
    UserName            As String
    InsertDate          As Date
    ModifyDate          As Date
End Type

Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum

Private DBData As StockHandData
Private m_blnCheck As Boolean

Public Property Let AccountNO(ByVal v_Value As Double)
    DBData.AccountNO = v_Value
End Property

Public Property Let MonthDp(ByVal v_Value As Long)
    DBData.MonthDp = v_Value
End Property

Public Property Let DateOfSubmission(ByVal v_Value As Date)
    DBData.DateOfSubmission = v_Value
End Property

Public Property Let ValueOfStock(ByVal v_Value As Double)
    DBData.ValueOfStock = v_Value
End Property

Public Property Let Margin(ByVal v_Value As Double)
    DBData.Margin = v_Value
End Property

Public Property Let LimitSanctioned(ByVal v_Value As Double)
    DBData.LimitSanctioned = v_Value
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

Public Property Get IsUniqueCheck() As Boolean
    IsUniqueCheck = m_blnCheck
End Property

'Checks weather account no exists in customer detail or not
Public Function CheckACNo(ByRef objDatasource As ADODB.Connection)
    Dim objCmd As New ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
        With objCmd
            .ActiveConnection = objDatasource
            
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.AccountNO)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adNumeric, adParamOutput)
                .Parameters.Append Param2
                
            .CommandText = "{call Pack_StockHandDp.CheckAccounNo(?,?)}"
            
            .Execute
                
                If Param2.Value = 0 Then
                    m_blnCheck = False
                Else
                    m_blnCheck = True
                End If
        End With
    Set objCmd = Nothing
End Function

'Checks weather Data exists for particular account no or not
Public Function CheckDPExists(ByRef objDatasource As ADODB.Connection)
    Dim objCmd As New ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
        With objCmd
            .ActiveConnection = objDatasource
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.AccountNO)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adNumeric, adParamInput, , DBData.MonthDp)
                .Parameters.Append Param2
            Set Param3 = .CreateParameter("Param3", adNumeric, adParamOutput)
                .Parameters.Append Param3
            .CommandText = "{call Pack_StockHandDp.CheckDPExists(?,?,?)}"
            
            .Execute
            
                If Param3.Value = 0 Then
                    m_blnCheck = True
                Else
                    m_blnCheck = False
                End If
        End With
    Set objCmd = Nothing
End Function

'To Save/Modify/Delete data in database
Public Sub SaveData(ByRef objDatasource As ADODB.Connection, ByVal Recordstate As DBState)
    Dim objCmd As ADODB.Command
    Set objCmd = New ADODB.Command
        With objCmd
            If Recordstate = DBInsert Then
                .CommandText = "Pack_StockHandDp.InsertData"
            ElseIf Recordstate = DBModify Then
                .CommandText = "Pack_StockHandDp.ModifyData"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "Pack_StockHandDp.DeleteData"
            End If
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
                .CommandTimeout = 0
                
                If Recordstate = DBInsert Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNO)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.MonthDp)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.DateOfSubmission, "dd-mm-yyyy")))
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ValueOfStock)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Margin)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.LimitSanctioned)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                ElseIf Recordstate = DBModify Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNO)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.MonthDp)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.DateOfSubmission, "dd-mm-yyyy")))
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ValueOfStock)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Margin)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.LimitSanctioned)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.ModifyDate, "dd-mm-yyyy")))
                ElseIf Recordstate = DBDelete Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNO)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.MonthDp)
                End If
                .Execute

        End With
    Set objCmd = Nothing
End Sub

'To retrieve list of stock hand exists in database for particular account no
Public Function GetStockHandList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Set objCmd = New ADODB.Command
        
        With objCmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.AccountNO)
                .Parameters.Append Param1
            .CommandText = "Pack_StockHandDp.GetStockHandDp()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetStockHandList = .Execute
        End With
    Set objCmd = Nothing
End Function

'To retreive account name and address for particular account no
Public Function GetAccountName(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim objCmd As ADODB.Command
    Set objCmd = New ADODB.Command
    
        With objCmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.AccountNO)
                .Parameters.Append Param1
            .CommandText = "Pack_StockHandDp.GetAccountDetail()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetAccountName = .Execute
        End With
    Set objCmd = Nothing
End Function
