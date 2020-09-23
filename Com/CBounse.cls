VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBounse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       CLASS      :CBounceApproval
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :HARISH CHAUHAN
'       MADE DATE  :17-04-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------

Private Type GroupData
        BounseKey                  As Long
        ClearingDate               As Date
        ChequeNo                   As Long
        ChequeDate                 As Date
        Reason                     As String
        Charges                    As Double
        TypeOfAccount              As String
        AccountNumber              As Long
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
    
        Public Function GetClearingList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
'                                 ByVal hwnd As Long)_
'                                 As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
'        Dim objOLEServer As New KccbUtils.OLERemoteServer
        With objCommand
             .CommandText = "PACK_BOUNSE_APPROVAL.GETCLEARINGDATA()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
             Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetClearingList = objRecordset

        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing

End Function

    Public Property Let BounseKey(ByVal v_Value As Long)
        DBData.BounseKey = v_Value
    End Property
    
    Public Property Let ClearingDate(ByVal v_Value As Date)
        DBData.ClearingDate = v_Value
    End Property

    Public Property Let ChequeNo(ByVal v_Value As Long)
        DBData.ChequeNo = v_Value
    End Property
    Public Property Let ChequeDate(ByVal v_Value As Date)
        DBData.ChequeDate = v_Value
    End Property
       
    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property
    
    Public Property Let AccountNumber(ByVal v_Value As Long)
        DBData.AccountNumber = v_Value
    End Property
    
    
    Public Property Let Reason(ByVal v_Value As String)
        DBData.Reason = v_Value
    End Property

    Public Property Let Charges(ByVal v_Value As Double)
        DBData.Charges = v_Value
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
    'This Procedure Inserts Or Modify data depending Upon the Mode
    Dim objCommand As New ADODB.Command
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_BOUNSE_APPROVAL.INSERT_BOUNSE_APPROVAL"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_BOUNSE_APPROVAL.MODIFY_BOUNSE_APPROVAL"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_BOUNSE_APPROVAL.DELETE_BOUNSE_APPROVAL"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.ChequeNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ChequeDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, DBData.Reason)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, DBData.Charges)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
            End If
'           If Recordstate = DBModify Then
'                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ClearingDate)
'                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
'                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
'                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.ChequeNo)
'                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.ChequeDate)
'                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.BankKey)
'                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, DBData.ChequeAmount)
'                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.OutstationLocal)
'                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, DBData.City)
'                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
'                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
'                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
'                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ClearingKey)
'            End If
'            If Recordstate = DBDelete Then
'                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ClearingKey)
'            End If
                .Execute
    End With
            Set objCommand = Nothing
End Sub

    Function GetGlNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_BOUNSE_APPROVAL.GET_GL_NO()"
            .CommandType = adCmdStoredProc
             .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, DBData.TypeOfAccount)
             Set .ActiveConnection = objDataSource
             .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetGlNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

