VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFdloan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    
'------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       CLASS      :CFDLOAN MASTER
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :HARISH CHAUHAN
'       MADE DATE  :07-04-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
Private Type GroupData
       
        TypeOfAccount              As String
        FdAccountNo                As Long
        FdLoanAccountNo            As Long
        FdNo                       As String
        LoanAmount                 As Double
        Rof                        As Double
        LoanDate                   As Date
        LoanSanctioned             As Double
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
    'This Function gets data from FD Details of the corrseponding Account
    Public Function GetFdList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset

        With objCommand
            .CommandText = "PACK_FD_LOAN.GET_FD_DETAILS()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.FdAccountNo)
                Set .ActiveConnection = objDataSource
                    .CommandTimeout = 0
                Set objRecordset = .Execute
                Set GetFdList = objRecordset
        End With
                Set objRecordset = Nothing
                Set objOLEServer = Nothing
                Set objCommand = Nothing
        Exit Function
End Function

    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property

    Public Property Let FdAccountNo(ByVal v_Value As Long)
       DBData.FdAccountNo = v_Value
    End Property
     
    Public Property Let FdLoanAccountNo(ByVal v_Value As Long)
       DBData.FdLoanAccountNo = v_Value
    End Property
    
    Public Property Let FdNo(ByVal v_Value As String)
        DBData.FdNo = v_Value
    End Property
    
    
    Public Property Let LoanAmount(ByVal v_Value As Double)
       DBData.LoanAmount = v_Value
    End Property

    Public Property Let Rof(ByVal v_Value As Integer)
       DBData.Rof = v_Value
    End Property

    Public Property Let LoanDate(ByVal v_Value As Date)
       DBData.LoanDate = v_Value
    End Property

    Public Property Let LoanSanctioned(ByVal v_Value As Double)
       DBData.LoanSanctioned = v_Value
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
        
Public Sub SaveDataMst(ByRef objDataSource As ADODB.Connection, _
                        ByVal Recordstate As DBState)
    'This Procedure Inserts and Modify Data depending upon the Mode
    Dim objCommand As New ADODB.Command
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_FD_LOAN.INSERT_FD_LOAN_MST"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_FD_LOAN.MODIFY_FD_LOAN_MST"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_FD_LOAN.DELETE_FD_LOAN_MST"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdLoanAccountNo)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdAccountNo)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.LoanAmount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Rof)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(DBData.LoanDate))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
             End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.LoanAmount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Rof)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.LoanDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdLoanAccountNo)
           End If
           If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdLoanAccountNo)
           End If

            .Execute
    End With
            Set objCommand = Nothing
End Sub

    Public Sub SaveDataTrn(ByRef objDataSource As ADODB.Connection, _
                        ByVal Recordstate As DBState)
    'This Procedure Inserts and Modify Data depending upon the Mode
    Dim objCommand As New ADODB.Command
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_FD_LOAN.INSERT_FD_LOAN_TRN"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_FD_LOAN.MODIFY_FD_LOAN_DETAIL"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_FD_LOAN.DELETE_FD_LOAN_DETAIL"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdLoanAccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.FdNo)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.LoanSanctioned)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
             End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.LoanSanctioned)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdLoanAccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.FdNo)
           End If
           If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.FdLoanAccountNo)
           End If
                .Execute
    End With
            Set objCommand = Nothing
End Sub

Function VerifyFdAccountNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_FD_LOAN.VERIFY_FD_ACCOUNT()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.FdAccountNo)

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set VerifyFdAccountNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

Function GetLastFdAccountNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_FD_LOAN.GET_LAST_ACCOUNT_NO()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetLastFdAccountNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objCommand = Nothing
        Exit Function
End Function
