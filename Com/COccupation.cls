VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COccupation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Private Type GroupData
        Occupationkey              As Integer
        OccupationCode             As String
        OccupationDesc             As String
        TerminalName               As String
        UserName                   As String
        InsertDate                 As Date
        ModifyDate                 As Date
    End Type
        Public Enum DBState
        DBInsert = 1
        DBModify = 2
        DBDelete = 3
    End Enum
    Private DBData                  As GroupData
    
    Public Function GetOccupationList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
                         '        ByVal hwnd As Long)_
                         '        As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
     '   Dim objOLEServer As New KccbUtils.OLERemoteServer
        With objCommand
            .CommandText = "PACK_OCCUPATION_MST.GETOCCUPATIONDATA()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetOccupationList = objRecordset

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
    
    Public Property Let Occupationkey(ByVal v_Value As Long)
        DBData.Occupationkey = v_Value
    End Property
    
    Public Property Let OccupationCode(ByVal v_Value As String)
       DBData.OccupationCode = v_Value
    End Property
    
    Public Property Let OccupationDesc(ByVal v_Value As String)
        DBData.OccupationDesc = v_Value
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
    
    On Error GoTo deleteError:
    
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_OCCUPATION_MST.INSERT_OCCUPATIONDATA"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_OCCUPATION_MST.MODIFY_OCCUPATIONDATA"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_OCCUPATION_MST.DELETE_OCCUPATIONDATA"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.OccupationCode)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, DBData.OccupationDesc)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
            End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.OccupationCode)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 40, DBData.OccupationDesc)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , CLng(DBData.Occupationkey))
           End If
           If Recordstate = DBDelete Then
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , CLng(DBData.Occupationkey))
           End If
             .Execute
    End With
            Set objCommand = Nothing

deleteError:
        Screen.MousePointer = vbDefault
        Dim errNumber As String
        Dim errString As String
        errString = Left(Err.Description, 9)
        errNumber = Right(errString, 5)
           If errNumber = "02292" Then
                  MsgBox "This Record is being used so it cannot be deleted", vbInformation, "Record Deletion"
           End If
        On Error GoTo 0:
End Sub
