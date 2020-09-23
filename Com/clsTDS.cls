VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type TDSMst
    TDS_Key                    As Long
    Status_Code                As Long
    TDS_Rate                   As Double
    Effective_Date             As Date
    TerminalName               As String
    UserName                   As String
    InsertModifyDate           As Date

End Type
Private SaveTDSMst     As TDSMst
Property Let TDS_Key(v_Value As Long)
    SaveTDSMst.TDS_Key = v_Value
End Property
Property Get TDS_Key() As Long
    TDS_Key = SaveTDSMst.TDS_Key
End Property
Property Let Status_Code(v_Value As Long)
    SaveTDSMst.Status_Code = v_Value
End Property
Property Get Status_Code() As Long
    Status_Code = SaveTDSMst.Status_Code
End Property
Property Let TDS_Rate(v_Value As Double)
    SaveTDSMst.TDS_Rate = v_Value
End Property
Property Get TDS_Rate() As Double
    TDS_Rate = SaveTDSMst.TDS_Rate
End Property
Property Let Effective_Date(v_Value As Date)
    SaveTDSMst.Effective_Date = v_Value
End Property
Property Get Effective_Date() As Date
    Effective_Date = SaveTDSMst.Effective_Date
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveTDSMst.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveTDSMst.UserName = v_Value
End Property
Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveTDSMst.InsertModifyDate = v_Value
End Property
Public Property Get TerminalName() As String
    TerminalName = SaveTDSMst.TerminalName
End Property
Public Property Get UserName() As String
    UserName = SaveTDSMst.UserName
End Property
Public Property Get InsertModifyDate() As Date
 InsertModifyDate = SaveTDSMst.InsertModifyDate
End Property
Public Function GETDATA(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "PACK_TDS_Master.GETDATA()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GETDATA = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Function GETStatusCode(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "PACK_TDS_Master.GETStatusCode()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GETStatusCode = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Sub InsertData(ByRef objDatasource As ADODB.Connection)
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "PACK_TDS_Master.InsertData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveTDSMst.Status_Code)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 5, SaveTDSMst.TDS_Rate)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveTDSMst.Effective_Date)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveTDSMst.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveTDSMst.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveTDSMst.InsertModifyDate)
                .Execute
            End With
            Set objCommand = Nothing
End Sub
Public Function ModifyData(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "PACK_TDS_Master.ModifyData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveTDSMst.TDS_Key)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveTDSMst.Status_Code)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 5, SaveTDSMst.TDS_Rate)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveTDSMst.Effective_Date)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveTDSMst.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveTDSMst.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveTDSMst.InsertModifyDate)
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
            .CommandText = "PACK_TDS_Master.DeleteRecord"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveTDSMst.TDS_Key)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveTDSMst.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveTDSMst.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveTDSMst.InsertModifyDate)
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


