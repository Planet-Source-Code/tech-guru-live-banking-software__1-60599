VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRdLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    
'------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       CLASS      :CFD MASTER
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :HARISH CHAUHAN
'       MADE DATE  :01-04-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
Private Type GroupData
       
        TypeOfAccount              As String
        RDAccountNo                As Long
        RDLoanAccountNo            As Long
        Amount                     As Double
        Rof                        As Double
        LoanDate                   As Date
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
    'This Function gets data from FD Master of the corrseponding Account
    Public Function GetRdList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset

        With objCommand
            .CommandText = "PACK_RD_LOAN_MST.GET_RD_DATA()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.RDAccountNo)

                Set .ActiveConnection = objDataSource
                    .CommandTimeout = 0
                Set objRecordset = .Execute
                Set GetRdList = objRecordset
        End With
                Set objRecordset = Nothing
                Set objOLEServer = Nothing
                Set objCommand = Nothing
        Exit Function
End Function

    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property

    Public Property Let RDAccountNo(ByVal v_Value As Long)
       DBData.RDAccountNo = v_Value
    End Property
     
    Public Property Let RDLoanAccountNo(ByVal v_Value As Long)
       DBData.RDLoanAccountNo = v_Value
    End Property
    
    Public Property Let Amount(ByVal v_Value As Double)
       DBData.Amount = v_Value
    End Property

    Public Property Let Rof(ByVal v_Value As Integer)
       DBData.Rof = v_Value
    End Property

    Public Property Let LoanDate(ByVal v_Value As Date)
       DBData.LoanDate = v_Value
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
    'This Procedure Inserts and Modify Data depending upon the Mode
    Dim objCommand As New ADODB.Command
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_RD_LOAN_MST.INSERT_RDDATA"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_RD_LOAN_MST.MODIFY_RDDATA"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_RD_LOAN_MST.DELETE_RDDATA"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.RDAccountNo)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.RDLoanAccountNo)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Amount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Rof)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(DBData.LoanDate))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
             End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Amount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.Rof)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.LoanDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.RDLoanAccountNo)
           End If
           If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.RDLoanAccountNo)
            End If
               
            .Execute
    End With
            Set objCommand = Nothing
End Sub


Function VerifyAccountNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_RD_LOAN_MST.VERIFY_RD_ACCOUNT()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.RDAccountNo)

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set VerifyAccountNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

Function GetNameAndAddress(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Procedure gets name and Balance of the Account holder
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
         With objCommand
              .CommandText = "PACK_RD_LOAN_MST.GET_NAME_ADDRESS()"
              .CommandType = adCmdStoredProc
              .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
              .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.RDAccountNo)

              Set .ActiveConnection = objDataSource
              .CommandTimeout = 0

              Set objRecordset = .Execute
              Set GetNameAndAddress = objRecordset
         End With
              Set objRecordset = Nothing
              Set objCommand = Nothing
              Exit Function
End Function

Function GetLastRdAccountNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_RD_LOAN_MST.GET_LAST_ACCOUNT_NO()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetLastRdAccountNo = objRecordset
        End With
        Set objRecordset = Nothing
        Set objCommand = Nothing
        Exit Function
End Function

Function GetTotalRdAmount(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        'This Function Verify that if the Entered account exists or not
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset
           With objCommand
             .CommandText = "PACK_RD_LOAN_MST.GET_TOTAL_RD_AMOUNT()"
            .CommandType = adCmdStoredProc
             .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
             .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.RDAccountNo)
             Set .ActiveConnection = objDataSource
            .CommandTimeout = 0

            Set objRecordset = .Execute
                Set GetTotalRdAmount = objRecordset
        End With
        Set objRecordset = Nothing
        Set objCommand = Nothing
        Exit Function
End Function
