VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type StatusMst
    Name              As String
    Description        As String
    Status_Code        As Long
    TerminalName            As String
    UserName                As String
    InsertModifyDate        As Date

End Type
Private SaveStatusMst     As StatusMst
Property Let Name(v_Value As String)
    SaveStatusMst.Name = v_Value
End Property
Property Get Name() As String
    Name = SaveStatusMst.Name
End Property
Property Let Description(v_Value As String)
    SaveStatusMst.Description = v_Value
End Property
Property Get Description() As String
    Description = SaveStatusMst.Description
End Property
Property Let Status_Code(v_Value As Long)
    SaveStatusMst.Status_Code = v_Value
End Property
Property Get Status_Code() As Long
    Status_Code = SaveStatusMst.Status_Code
End Property

Public Property Let TerminalName(ByVal v_Value As String)
    SaveStatusMst.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveStatusMst.UserName = v_Value
End Property
Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveStatusMst.InsertModifyDate = v_Value
End Property
Public Property Get TerminalName() As String
    TerminalName = SaveStatusMst.TerminalName
End Property
Public Property Get UserName() As String
    UserName = SaveStatusMst.UserName
End Property
Public Property Get InsertModifyDate() As Date
 InsertModifyDate = SaveStatusMst.InsertModifyDate
End Property
Public Function GETDATA(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "PACK_Status_Master.GETDATA()"
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
                    .CommandText = "PACK_Status_Master.InsertData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveStatusMst.Status_Code)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveStatusMst.Name)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveStatusMst.Description)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveStatusMst.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveStatusMst.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveStatusMst.InsertModifyDate)
                .Execute
            End With
            Set objCommand = Nothing
End Sub
Public Function ModifyData(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                    .CommandText = "PACK_Status_Master.ModifyData"
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveStatusMst.Status_Code)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveStatusMst.Name)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveStatusMst.Description)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveStatusMst.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveStatusMst.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveStatusMst.InsertModifyDate)
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
            .CommandText = "PACK_Status_Master.DeleteRecord"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveStatusMst.Status_Code)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveStatusMst.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveStatusMst.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveStatusMst.InsertModifyDate)
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
Public Function CheckCode(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "PACK_Status_Master.CheckCode"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, SaveStatusMst.Status_Code)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        CheckCode = True
                    Else
                        CheckCode = False
                    End If
                End If
        End With
End Function


