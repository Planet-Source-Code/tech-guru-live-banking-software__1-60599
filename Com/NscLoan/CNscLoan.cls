VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNscLoan"
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
'       MADE DATE  :01-07-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
Private Type GroupData
       
        TypeOfAccount              As String
        PublicStaff                As String
        AccountNo                  As Long
        TrnKey                     As Long
        LoanDate                   As Date
        Name                       As String
        FatherName                 As String
        Address                    As String
        AmountSanctioned           As Double
        NscNo                      As String
        RegdNo                     As String
        PurchaseDate               As Date
        MaturityDate               As Date
        NscAmount                  As Double
        MaturityAmount             As Double
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
'    'This Function gets data from FD Master of the corrseponding Account
'    Public Function GetRdList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
'        Dim objCommand As New ADODB.Command
'        Dim objRecordset As ADODB.Recordset
'
'        With objCommand
'            .CommandText = "PACK_RD_LOAN_MST.GET_RD_DATA()"
'            .CommandType = adCmdStoredProc
'            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.RDAccountNo)
'
'                Set .ActiveConnection = objDataSource
'                    .CommandTimeout = 0
'                Set objRecordset = .Execute
'                Set GetRdList = objRecordset
'        End With
'                Set objRecordset = Nothing
'                Set objOLEServer = Nothing
'                Set objCommand = Nothing
'        Exit Function
'End Function

    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property

    Public Property Let PublicStaff(ByVal v_Value As String)
        DBData.PublicStaff = v_Value
    End Property
    
    Public Property Let AccountNo(ByVal v_Value As Long)
       DBData.AccountNo = v_Value
    End Property
     
    Public Property Let TrnKey(ByVal v_Value As Long)
       DBData.TrnKey = v_Value
    End Property
    
    Public Property Let LoanDate(ByVal v_Value As Date)
       DBData.LoanDate = v_Value
    End Property

    Public Property Let Name(ByVal v_Value As String)
        DBData.Name = v_Value
    End Property
    
    Public Property Let FatherName(ByVal v_Value As String)
        DBData.FatherName = v_Value
    End Property
    
    Public Property Let Address(ByVal v_Value As String)
        DBData.Address = v_Value
    End Property
    
    Public Property Let AmountSanctioned(ByVal v_Value As Double)
        DBData.AmountSanctioned = v_Value
    End Property
    
    Public Property Let NscNo(ByVal v_Value As String)
        DBData.NscNo = v_Value
    End Property
    
    Public Property Let RegdNo(ByVal v_Value As String)
        DBData.RegdNo = v_Value
    End Property
    
    Public Property Let PurchaseDate(ByVal v_Value As Date)
        DBData.PurchaseDate = v_Value
    End Property
    
    Public Property Let MaturityDate(ByVal v_Value As Date)
        DBData.MaturityDate = v_Value
    End Property
    
    Public Property Let NscAmount(ByVal v_Value As Double)
        DBData.NscAmount = v_Value
    End Property
    
    Public Property Let MaturityAmount(ByVal v_Value As Double)
        DBData.MaturityAmount = v_Value
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
                .CommandText = "PACK_NSC_DATA.INSERT_NSC_MST"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_NSC_DATA.MODIFY_NSC_MST"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_NSC_DATA.DELETE_NSC_MST"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.PublicStaff)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(DBData.LoanDate))
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Name)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.FatherName)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.Address)
                 .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AmountSanctioned)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                 .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                 .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
             End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.LoanDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Name)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.FatherName)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.Address)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AmountSanctioned)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                
           End If
           If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
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
                .CommandText = "PACK_NSC_DATA.INSERT_NSC_DETAIL"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_NSC_DATA.MODIFY_NSC_DETAIL"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_NSC_DATA.DELETE_NSC_DETAIL"
            End If
            .CommandType = adCmdStoredProc

            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.NscNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.RegdNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.PurchaseDate)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.MaturityDate)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.NscAmount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.MaturityAmount)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
             End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.NscNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.RegdNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.PurchaseDate)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.MaturityDate)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.NscAmount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.TrnKey)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
           End If
           If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.TrnKey)
           End If
                .Execute
    End With
            Set objCommand = Nothing
End Sub

    Public Function GetLastAccountNo(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset

        With objCommand
            .CommandText = "PACK_NSC_DATA.GET_MAX_ACCOUNT_NO()"
            .CommandType = adCmdStoredProc


                Set .ActiveConnection = objDataSource
                    .CommandTimeout = 0
                Set objRecordset = .Execute
                Set GetLastAccountNo = objRecordset
        End With
                Set objRecordset = Nothing
                Set objOLEServer = Nothing
                Set objCommand = Nothing
        Exit Function
End Function

Public Function GetNscTransactions(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        Dim objRecordset As ADODB.Recordset

        With objCommand
            .CommandText = "PACK_NSC_DATA.GET_NSC_TRANSACTIONS()"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

                Set .ActiveConnection = objDataSource
                    .CommandTimeout = 0
                Set objRecordset = .Execute
                Set GetNscTransactions = objRecordset
        End With
                Set objRecordset = Nothing
                Set objOLEServer = Nothing
                Set objCommand = Nothing
        Exit Function
End Function

