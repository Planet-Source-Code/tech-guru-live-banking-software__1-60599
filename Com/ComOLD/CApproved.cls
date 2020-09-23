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
                                                                                                                                                                                       
'To retrieve list of stock hand exists in database for particular account no
Public Function GetStockHandList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Set objCmd = New ADODB.Command
        
        With objCmd
            'Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.AccountNO)
             '   .Parameters.Append Param1
            .CommandText = "PACK_MDIAPPROVED.GetStockHandDp()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetStockHandList = .Execute
        End With
    Set objCmd = Nothing
End Function
 
Public Function GetCustomerList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Set objCmd = New ADODB.Command
        
        With objCmd
            'Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.AccountNO)
             '   .Parameters.Append Param1
            .CommandText = "PACK_MDIAPPROVED.GETALLACCOUNTS()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
            .CommandTimeout = 0
            
            Set GetCustomerList = .Execute
        End With
    Set objCmd = Nothing
End Function
 
