VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApproved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'------------------------------------------------------------------
'       PROJECT    :KCCAPPROVAL
'       MODULE     :ESTABLISHMENT
'       PRUPOSE    :A genric solution for the stockhand drawing power
'
'       COM        :APPROVAL
'       FORM
'       REFERRED   :
'       FORM CODE  :KCCB EST 0004
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :GURINDER MANN
'       MADE DATE  :
'       REVISIONS  :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
                                                                                                                                                                               
Private Type AppovedTo
    m_Isapprove         As String
    m_TypeOfAccount     As String
    m_TypeOfEntry       As String
    m_AppGroupKey       As Long
    m_ApproveAmount     As Long
    m_AppCriteria       As String
End Type

Private Type DoApprove
    m_TableName         As String
    m_AppTerminalCol    As String
    m_AppTerminalValue  As String
    m_AppUserCol        As String
    m_AppUserName       As String
    m_AppDateCol        As String
    m_ApprovalDate      As Date
    m_AppCol            As String
    m_App_Date_Col      As String
    m_AppColValue       As String
    m_IsApprovedBy      As String
End Type

Private Type InsertMoneyTransDDPo
    m_Type_of_acc       As String
    m_Account_NO        As Double
    m_TransType         As String
    m_PayRec            As String
    m_TrnDate           As Date
    m_DDPONo            As Double
    m_DDPoDate          As Date
    m_Remarks           As String
    m_DR1Amount         As Double
    m_Terminal          As String
    m_User              As String
    m_Date              As Date
    m_2Type_of_acc      As String
    m_2Account_NO       As Double
    m_CR1Amount         As Double
    m_3Type_of_acc      As String
    m_3Account_NO       As Double
    m_DR2Amount         As Double
    m_4Type_of_acc      As String
    m_4Account_NO       As Double
    m_CR2Amount         As Double
    m_chequedate        As Variant
End Type
                                                                                                                                                                               
Private DApprove    As DoApprove
Private Approve     As AppovedTo
Private DDPOIssue   As InsertMoneyTransDDPo


Public Property Let ChequeDateSI(ByVal v_Value As Variant)
    DDPOIssue.m_chequedate = v_Value
End Property

 Public Property Let TerminalName(ByVal v_Value As String)
    DDPOIssue.m_Terminal = v_Value
End Property

Public Property Let IsApprovedNo(ByVal v_Value As String)
    DApprove.m_IsApprovedBy = v_Value
End Property
 Public Property Let DDUserName(ByVal v_Value As String)
    DDPOIssue.m_User = v_Value
End Property
Public Property Let App_date_Col(ByVal v_Value As String)
    DApprove.m_App_Date_Col = v_Value
End Property

Public Property Let TerminalDate(ByVal v_Value As Date)
    DDPOIssue.m_Date = v_Value
End Property

 Public Property Let Remarks(ByVal v_Value As String)
    DDPOIssue.m_Remarks = v_Value
End Property

Public Property Let DDPODate(ByVal v_Value As Date)
    DDPOIssue.m_DDPoDate = v_Value
End Property

Public Property Let DDPoNUmber(ByVal v_Value As Double)
    DDPOIssue.m_DDPONo = v_Value
End Property

Public Property Let TRansDate(ByVal v_Value As Date)
    DDPOIssue.m_TrnDate = v_Value
End Property

Public Property Let PaymentReceived(ByVal v_Value As String)
    DDPOIssue.m_PayRec = v_Value
End Property

Public Property Let TypeOfTrans(ByVal v_Value As String)
    DDPOIssue.m_TransType = v_Value
End Property

Public Property Let TypeOfAcc(ByVal v_Value As String)
    DDPOIssue.m_Type_of_acc = v_Value
End Property
Public Property Let Accountno(ByVal v_Value As Double)
    DDPOIssue.m_Account_NO = v_Value
End Property
Public Property Let DR1Amount(ByVal v_Value As Double)
    DDPOIssue.m_DR1Amount = v_Value
End Property

Public Property Let TypeOfAcc2(ByVal v_Value As String)
    DDPOIssue.m_2Type_of_acc = v_Value
End Property
Public Property Let Accountno2(ByVal v_Value As Double)
    DDPOIssue.m_2Account_NO = v_Value
End Property
Public Property Let CR1Amount(ByVal v_Value As Double)
    DDPOIssue.m_CR1Amount = v_Value
End Property



Public Property Let TypeOfAcc3(ByVal v_Value As String)
    DDPOIssue.m_3Type_of_acc = v_Value
End Property
Public Property Let Accountno3(ByVal v_Value As Double)
    DDPOIssue.m_3Account_NO = v_Value
End Property
Public Property Let DR2Amount(ByVal v_Value As Double)
    DDPOIssue.m_DR2Amount = v_Value
End Property

Public Property Let TypeOfAcc4(ByVal v_Value As String)
    DDPOIssue.m_4Type_of_acc = v_Value
End Property
Public Property Let Accountno4(ByVal v_Value As Double)
    DDPOIssue.m_4Account_NO = v_Value
End Property
Public Property Let CR2Amount(ByVal v_Value As Double)
    DDPOIssue.m_CR2Amount = v_Value
End Property
Public Property Let ApprovalColName(ByVal v_AppColName As String)
    DApprove.m_AppCol = v_AppColName
End Property

Public Property Let AppColValue(ByVal v_AppColValue As String)
    DApprove.m_AppColValue = v_AppColValue
End Property

Public Property Let TableName(ByVal v_TableName As String)
    DApprove.m_TableName = v_TableName
End Property

Public Property Let AppDateCol(ByVal v_AppdateCol As String)
    DApprove.m_AppDateCol = v_AppdateCol
End Property

Public Property Let ApprovalDate(ByVal v_ApprovalDate As Date)
    DApprove.m_ApprovalDate = v_ApprovalDate
End Property

Public Property Let TerminalColName(ByVal v_TerminalColName As String)
    DApprove.m_AppTerminalCol = v_TerminalColName
End Property

Public Property Let ApprovalCriteria(ByVal v_Value As String)
    Approve.m_AppCriteria = v_Value
End Property


Public Property Let TerminalValue(ByVal v_TerminalName As String)
    DApprove.m_AppTerminalValue = v_TerminalName
End Property

Public Property Let UserCol(ByVal v_UserCol As String)
    DApprove.m_AppUserCol = v_UserCol
End Property

Public Property Let UserName(ByVal v_UserName As String)
    DApprove.m_AppUserName = v_UserName
End Property

Public Property Let AppGroupKey(ByVal v_Groupkey As Long)
    Approve.m_AppGroupKey = v_Groupkey
End Property

Public Property Let AppAmount(ByVal v_AppAmount As Long)
    Approve.m_ApproveAmount = v_AppAmount
End Property

Public Property Let Approval(ByVal v_Approve As String)
    Approve.m_Isapprove = v_Approve
End Property

Public Property Let TypeOfAccount(ByVal v_TypeOfAccount As String)
    Approve.m_TypeOfAccount = v_TypeOfAccount
End Property

Public Property Let TypeOfEntry(ByVal v_TypeOfEntry As String)
    Approve.m_TypeOfEntry = v_TypeOfEntry
End Property


'To check is already approved or not
Public Function CheckApproval(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd  As ADODB.Command
    Dim Param1  As New ADODB.Parameter
    Dim Param2  As New ADODB.Parameter
    Dim Param3  As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adVarChar, adParamInput, 30, DApprove.m_TableName)
                .Parameters.Append Param1
            
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 20, DApprove.m_AppCol)
                .Parameters.Append Param2
                
            Set Param3 = .CreateParameter("Param3", adVarChar, adParamInput, 9, DApprove.m_AppColValue)
                .Parameters.Append Param3
                
            .CommandText = "kccbank.PACK_MDIAPPROVED.CheckApproval()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            Set CheckApproval = .Execute
            
        End With
End Function



'To checkJV is already approved or not
Public Function CheckJVApproval(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd  As ADODB.Command
    Dim Param1  As New ADODB.Parameter
    Dim Param2  As New ADODB.Parameter
    Dim Param3  As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adVarChar, adParamInput, 30, DApprove.m_TableName)
                .Parameters.Append Param1
            
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 20, DApprove.m_AppCol)
                .Parameters.Append Param2
                
            Set Param3 = .CreateParameter("Param3", adVarChar, adParamInput, 9, DApprove.m_AppColValue)
                .Parameters.Append Param3
                
            Set Param4 = .CreateParameter("Param4", adDate, adParamInput, 8, CDate(Format(DApprove.m_ApprovalDate, "dd-mm-yyyy")))
                .Parameters.Append Param4
            
            .CommandText = "kccbank.PACK_MDIAPPROVED.CheckJvApproval()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            Set CheckJVApproval = .Execute
            
        End With
End Function


'To Approve any table
Public Function ApproveTable(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objcmd  As ADODB.Command
    Dim Param1  As New ADODB.Parameter
    Dim Param2  As New ADODB.Parameter
    Dim Param3  As New ADODB.Parameter
    Dim Param4  As New ADODB.Parameter
    Dim Param5  As New ADODB.Parameter
    Dim Param6  As New ADODB.Parameter
    Dim Param7  As New ADODB.Parameter
    Dim Param8  As New ADODB.Parameter
    Dim Param9  As New ADODB.Parameter
    Dim Param10 As New ADODB.Parameter
    
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adVarChar, adParamInput, 30, DApprove.m_TableName)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 30, DApprove.m_AppTerminalCol)
                .Parameters.Append Param2
            Set Param3 = .CreateParameter("Param3", adVarChar, adParamInput, 30, DApprove.m_AppUserCol)
                .Parameters.Append Param3
            Set Param4 = .CreateParameter("Param4", adVarChar, adParamInput, 30, DApprove.m_AppDateCol)
                .Parameters.Append Param4
            Set Param5 = .CreateParameter("Param5", adVarChar, adParamInput, 50, DApprove.m_AppTerminalValue)
                .Parameters.Append Param5
            Set Param6 = .CreateParameter("Param6", adVarChar, adParamInput, 100, DApprove.m_AppUserName)
                .Parameters.Append Param6
            Set Param7 = .CreateParameter("Param7", adDate, adParamInput, 8, CDate(Format(DApprove.m_ApprovalDate, "dd-mm-yyyy")))
                .Parameters.Append Param7
            Set Param8 = .CreateParameter("Param8", adVarChar, adParamInput, 20, DApprove.m_AppCol)
                .Parameters.Append Param8
            
            Set Param9 = .CreateParameter("Param9", adVarChar, adParamInput, 9, DApprove.m_AppColValue)
                .Parameters.Append Param9
                
            Set Param10 = .CreateParameter("Param10", adVarChar, adParamInput, 12, DApprove.m_IsApprovedBy)
                .Parameters.Append Param10
            
            .CommandText = "kccbank.PACK_MDIAPPROVED.ApprovedTable()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            .Execute
            ApproveTable = True
        End With
End Function

'To save data in MoneyTransaction
Public Function InsertMoneyTransDDPo(ByRef objDatasource As ADODB.Connection)
    Dim objcmd As New ADODB.Command
        With objcmd
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_Type_of_acc)
            If DDPOIssue.m_Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_TransType)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_PayRec)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_TrnDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_DDPoDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, DDPOIssue.m_Remarks)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DR1Amount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DDPOIssue.m_Terminal)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DDPOIssue.m_User)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_Date, "dd-mm-yyyy")))
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_2Type_of_acc)
            If DDPOIssue.m_2Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_2Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_CR1Amount)
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_3Type_of_acc)
            If DDPOIssue.m_3Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_3Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DR2Amount)
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_4Type_of_acc)
            If DDPOIssue.m_4Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_4Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_CR2Amount)
            

            .CommandText = "kccbank.PACK_DDPO_DATA.InsertMoneyTransDDPO()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            .Execute
            
        End With
End Function


'To save data in MoneyTransactionb
Public Function InsertMoneyDDPayment(ByRef objDatasource As ADODB.Connection)
        Dim objcmd As New ADODB.Command
        With objcmd
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_Type_of_acc)
            If DDPOIssue.m_Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_TransType)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_PayRec)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_TrnDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_DDPoDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, DDPOIssue.m_Remarks)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DR1Amount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DDPOIssue.m_Terminal)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DDPOIssue.m_User)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_Date, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_2Type_of_acc)
            If DDPOIssue.m_2Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_2Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_CR1Amount)
            
            .CommandText = "kccbank.PACK_DDPayment_DATA.InsertMoneyDDPayment()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            .Execute
            
        End With
End Function


'To save data in MoneyTransaction
Public Function InsertMoneyPoPayment(ByRef objDatasource As ADODB.Connection)
    Dim objcmd As New ADODB.Command
        With objcmd
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_Type_of_acc)
            If DDPOIssue.m_Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_TransType)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_PayRec)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_TrnDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_DDPoDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, DDPOIssue.m_Remarks)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DR1Amount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DDPOIssue.m_Terminal)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DDPOIssue.m_User)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_Date, "dd-mm-yyyy")))
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_2Type_of_acc)
            If DDPOIssue.m_2Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_2Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_CR1Amount)
            
            .CommandText = "kccbank.PACK_DDPO_DATA.InsertMoneyDDPOPayment()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            .Execute
            
        End With
End Function

'To retrieve list of stock hand exists in database for particular account no
Public Function GetStockHandList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd As ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Set objcmd = New ADODB.Command
         With objcmd
            Set Param1 = .CreateParameter("Param1", adVarChar, adParamInput, 15, Approve.m_Isapprove)
                .Parameters.Append Param1
            .CommandText = "kccbank.PACK_MDIAPPROVED.GetStockHandDp()"
            
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            Set GetStockHandList = .Execute
       End With
    Set objcmd = Nothing
End Function



'To retreive detail of customer for particular account no and type of account which will be passed from verifyform
Public Function GetCustomerList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
        
        With objcmd
            Set Param1 = .CreateParameter("Param1", adVarChar, adParamInput, 12, Approve.m_Isapprove)
                .Parameters.Append Param1
            .CommandText = "PACK_MDIAPPROVED.GETALLACCOUNTS()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetCustomerList = .Execute
        End With
    Set objcmd = Nothing
End Function
 
'To retrieve detail of clearing account details
Public Function GetClearingDetail(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
  
    Set objcmd = New ADODB.Command
        
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIAPPROVED.GETCLEARINGDATA()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetClearingDetail = .Execute
        End With
    Set objcmd = Nothing

End Function

'To retrieve detail of cheque book issue details
Public Function GetChequeBookDetail(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset


    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command

        With objcmd
            .CommandText = "PACK_MDIAPPROVED.GETCHEQUEBOOKDATA()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0

            Set GetChequeBookDetail = .Execute
        End With
    Set objcmd = Nothing

End Function

'To retrieve approval no of user according to his group key
Public Function GetApprovalNo(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd As ADODB.Command
  
    Set objcmd = New ADODB.Command
        
        With objcmd
            .CommandText = "PACK_MDIAPPROVED.GetApprovalNo()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetApprovalNo = .Execute
        End With
    Set objcmd = Nothing

End Function

'To retreive moneytransaction data for approval
Public Function GetMoneyTransaction(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "Pack_MDIApproved.GetMoneyTransaction()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            Set GetMoneyTransaction = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive jvMaster enteries for approval
Public Function GetJVMasterdetail(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            .CommandText = "Pack_MDIApproved.GETJVDETAILS()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            Set GetJVMasterdetail = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive ChequeUsed enteries for approval
Public Function GetUsedChequeDetail(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            .CommandText = "Pack_MDIApproved.GetUsedChequesDetail()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            Set GetUsedChequeDetail = .Execute
        End With
        
        Set objcmd = Nothing
End Function


'To retreive RDdetail enteries for approval
Public Function GetRDDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETRDDETAIL()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetRDDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive Bounce details for approval
Public Function GetBounceDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
                .CommandText = "PACK_MDIApproved.GETBounceDATA()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetBounceDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive FDMASTER details for approval
Public Function GetFDDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            .CommandText = "PACK_MDIApproved.GetFDMasterdetail()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetFDDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function


'To retreive NscLoan details for approval
Public Function GetNscDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            .CommandText = "PACK_MDIApproved.GetNSCDetails()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetNscDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function



'To retreive RDLOANMASTER details for approval
Public Function GetRDLoanDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETRDLOANDETAIL()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetRDLoanDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive FDLOANMASTER details for approval
Public Function GetFDLoanDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETFDLOANDETAIL()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetFDLoanDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive DDIssuedetails for approval
Public Function GetDDIssueDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETDDISSUE()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetDDIssueDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

'To retreive DDPaymnet Details for approval
Public Function GetDDPaymnetDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETDDPAYMENT()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            Set GetDDPaymnetDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function


'To retreive POPaymentdetails for approval
Public Function GetPOPaymentDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETPOPAYMENT()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetPOPaymentDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

Public Function GetPOIssueDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETPOISSUE()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetPOIssueDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function


Public Function GetDDCancelDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETDDCancel()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetDDCancelDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

Public Function GetPORevalidDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETPORevalid()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetPORevalidDetails = .Execute
        End With
        
        Set objcmd = Nothing

End Function

Public Function GetDDRevalidDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETDDRevalid()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetDDRevalidDetails = .Execute
        End With
        
        Set objcmd = Nothing

End Function


Public Function GetPOCancelDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GETPOCancel()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetPOCancelDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

Public Function GetPOStopPaymentDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GetPOStopPayment()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetPOStopPaymentDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

Public Function GetPOStopPaymentReverseDetails(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GetPOStopPaymentRev()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetPOStopPaymentReverseDetails = .Execute
        End With
        
        Set objcmd = Nothing
End Function

Public Function GetFDDetailsApp(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
             Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GetFDDetails()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetFDDetailsApp = .Execute
        End With
        
        Set objcmd = Nothing
End Function


Public Function GetAppStandingInstDetail(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , Approve.m_ApproveAmount)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adVarChar, adParamInput, 2, Approve.m_AppCriteria)
                .Parameters.Append Param2
            .CommandText = "PACK_MDIApproved.GetAppStandingAll()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetAppStandingInstDetail = .Execute
        End With
        
        Set objcmd = Nothing
End Function

Public Function GetAgriLoanDetail(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objcmd As ADODB.Command
    Set objcmd = New ADODB.Command
    
        With objcmd
            .CommandText = "PACK_MDIApproved.GetAgriLoanDetail()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetAgriLoanDetail = .Execute
        End With
        
        Set objcmd = Nothing
End Function



'Public Function DisplayMoneyTransaction(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
'    Dim param1 As New ADODB.Parameter
'    Dim param2 As New ADODB.Parameter
'    Dim param3 As New ADODB.Parameter
'    Dim objCmd As New ADODB.Command
'
'    With objCmd
'        Set param1 = .CreateParameter("param1", adVarChar, adParamInput, 12, Approve.m_Isapprove)
'            .Parameters.Append param1
'        Set param2 = .CreateParameter("param2", adChar, adParamInput, 2, Approve.m_TypeOfAccount)
'            .Parameters.Append param2
'        Set param3 = .CreateParameter("param3", adChar, adParamInput, 1, Approve.m_TypeOfEntry)
'            .Parameters.Append param3
'        .CommandText = "Pack_MDIApproved.GetMoneyTransaction()"
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = objDatasource
    '        .CommandTimeout = 0
'        Set DisplayMoneyTransaction = .Execute
'    End With
'End Function


Public Function InsertMoneyStandInstruct(ByRef objDatasource As ADODB.Connection)
    Dim objcmd As New ADODB.Command
        With objcmd
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_Type_of_acc)
            If DDPOIssue.m_Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_TransType)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DDPOIssue.m_PayRec)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_TrnDate, "dd-mm-yyyy")))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DDPONo)
            If DDPOIssue.m_chequedate = Empty Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_chequedate, "dd-mm-yyyy")))
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, DDPOIssue.m_Remarks)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_DR1Amount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DDPOIssue.m_Terminal)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DDPOIssue.m_User)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , CDate(Format(DDPOIssue.m_Date, "dd-mm-yyyy")))
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DDPOIssue.m_2Type_of_acc)
            If DDPOIssue.m_2Account_NO = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_2Account_NO)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DDPOIssue.m_CR1Amount)
            
            .CommandText = "kccbank.PACK_DDPO_DATA.InsertMoneyDDPOPayment()"
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            .CommandType = adCmdStoredProc
            .Execute
            
        End With
End Function

