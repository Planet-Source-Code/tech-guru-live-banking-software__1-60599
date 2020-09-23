VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGrpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    Private Type GroupData
        GroupKey                As Long
        GroupCode               As String
        GroupDescription        As String
        TerminalName            As String
        UserName                As String
        InsertDate              As Date
        ModifyDate              As Date
        GroupStatus             As String
    End Type
    
    Public Enum DBState
        DBInsert = 1
        DBModify = 2
        DBDelete = 3
    End Enum
    
    Private DBData                  As GroupData
    
    Public Function GetGroupList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
                         '        ByVal hwnd As Long)_
                         '        As ADODB.Recordset
                                 
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
        Dim objOLEServer As New KccbUtils.OLERemoteServer
        
        With objCommand
            .CommandText = "kccbank.Pack_Group_mst_Data.GetGroupData()"
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetGroupList = objRecordset


'            If objRecordset.RecordCount > 0 Then
'                Set GetGroupList = objRecordset
'            Else
'                MsgBox "No Group Exists in database. Quiting!", vbOKOnly + vbCritical, "KCCB Error"
'                objOLEServer.ShutdownRemoteOLEServer (hwnd)
'            End If
        End With
            

        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing

        Exit Function
                                 
    End Function
    
    Public Property Let GroupCode(ByVal v_Value As String)
        DBData.GroupCode = v_Value
    End Property
    
    Public Property Let GroupStatus(ByVal v_Value As String)
        DBData.GroupStatus = v_Value
    End Property
    
    Public Property Let GroupDescription(ByVal v_Value As String)
        DBData.GroupDescription = v_Value
    End Property
    
    Public Property Let GroupKey(ByVal v_Value As Long)
        DBData.GroupKey = v_Value
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
                .CommandText = "kccbank.Pack_Group_mst_Data.Insert_Groupdata"
            ElseIf Recordstate = DBModify Then
                .CommandText = "kccbank.Pack_Group_mst_Data.Modify_Groupdata"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "kccbank.Pack_Group_Data.DeleteGroupData"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.GroupCode)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 60, DBData.GroupDescription)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.GroupStatus)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
            
            End If
            
            If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.GroupCode)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 60, DBData.GroupDescription)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.GroupStatus)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, CLng(DBData.GroupKey))
            
           
            End If
        
             .Execute
        
        
        End With
        Set objCommand = Nothing
        
    End Sub
