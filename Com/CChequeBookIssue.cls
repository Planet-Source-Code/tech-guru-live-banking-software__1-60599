VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChequeBookIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  '------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       CLASS      :CChequeBookIssue
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :HARISH CHAUHAN
'       MADE DATE  :19-03-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
  
  Private Type GroupData
        ChequebookKey              As Long
        TypeOfAccount              As String
        AccountNo                  As Long
        LooseCheque                As String
        ChequeNoFrom               As Long
        ChequeNoTo                 As Long
        MoneyCharged               As Long
        IssueDate                  As Date
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
    
    Public Function GetChequeBookIssueList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function gets data from ChequeBook master for the corrseponding Account
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
    
        With objCommand
             .CommandText = "PACK_CHEQUEBOOK_ISSUE_MST.GETCHEQUEBOOKDATA()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)
            
            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetChequeBookIssueList = objRecordset

        End With
        
            Set objRecordset = Nothing
            Set objOLEServer = Nothing
            Set objCommand = Nothing
        Exit Function
End Function
    Public Property Let ChequebookKey(ByVal v_Value As String)
        DBData.ChequebookKey = v_Value
    End Property
    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property
    
    Public Property Let AccountNo(ByVal v_Value As Long)
       DBData.AccountNo = v_Value
    End Property
    
    Public Property Let LooseCheque(ByVal v_Value As String)
        DBData.LooseCheque = v_Value
    End Property
    Public Property Let ChequeNoFrom(ByVal v_Value As Long)
        DBData.ChequeNoFrom = v_Value
    End Property
    Public Property Let ChequeNoTo(ByVal v_Value As Long)
        DBData.ChequeNoTo = v_Value
    End Property
    Public Property Let MoneyCharged(ByVal v_Value As Long)
        DBData.MoneyCharged = v_Value
    End Property
    Public Property Let IssueDate(ByVal v_Value As Date)
        DBData.IssueDate = v_Value
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
        'This procedure Inserts and Modify data depending Upon the Mode
         Dim objCommand As New ADODB.Command
        
        With objCommand
             If Recordstate = DBInsert Then
                 .CommandText = "PACK_CHEQUEBOOK_ISSUE_MST.INSERT_CHEQUEBOOKDATA"
             ElseIf Recordstate = DBModify Then
                 .CommandText = "PACK_CHEQUEBOOK_ISSUE_MST.MODIFY_CHEQUEBOOKDATA"
             ElseIf Recordstate = DBDelete Then
                 .CommandText = "PACK_CHEQUEBOOK_ISSUE_MST.DELETE_CHEQUEBOOKDATA"
             End If
             .CommandType = adCmdStoredProc
    
             Set .ActiveConnection = objDataSource
             .CommandTimeout = 0
             If Recordstate = DBInsert Then
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.LooseCheque)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNoFrom)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNoTo)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.MoneyCharged)
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.IssueDate)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
             End If
            If Recordstate = DBModify Then
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.LooseCheque)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNoFrom)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequeNoTo)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.MoneyCharged)
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.IssueDate)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequebookKey)
                
             End If
            If Recordstate = DBDelete Then
                  .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ChequebookKey)
            End If
            .Execute
        End With
        Set objCommand = Nothing
        
End Sub

Function GetNameAndBalance(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function gets name and balance of the Account Holder
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_CHEQUEBOOK_ISSUE_MST.GETNAMEANDBALANCE()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)
            
            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetNameAndBalance = objRecordset
        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

Function VerifyAccountNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_CLEARING_MST.VERIFY_ACCOUNTNO()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set VerifyAccountNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objOLEServer = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

