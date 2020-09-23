VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_DDPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**********************************************
'Class Name         :- C_DDPO.Cls
'Pupose             :- Class created for DDPOISSue/Payment/Cancellation/Revalidation
'Referred dll       :-
'Date of Creation   :- 22nd April'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************


'Variables declared for saving on Deletion,Insertion,Modifcation
Private Type DDPOIssueData
    SecurityNo          As Double
    DD_PO               As String
    TransType           As String
    DDPONo              As Double
    DDDate              As Date
    FavourOF            As String
    DDAmount            As Single
    DepositorName       As String
    Commission          As Single
    CityName            As String
    AcType              As String
    AcNo                As Long
    Status              As String
    PaymentMode         As String
    PaymentDt           As Date
    CancelDt            As Date
    CancelCharges       As Single
'    ISRevalidated       As String
'    RevalidCharge       As Single
'    ISStopPayment       As String
'    StopPayCharge       As Single
'    StopPayRevCharge    As Single
    ISdeleted           As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
End Type

'Enumeration to define different stages of insert/modify and delete
Public Enum DBStateDD
    DBInsertDD = 1
    DBModifyDD = 2
    DBDeleteDD = 3
    DBInsertPOPayment = 4
End Enum

Private SaveDDPOIssue                As DDPOIssueData
Public Property Let PaymentMode(ByVal v_Value As String)
    SaveDDPOIssue.PaymentMode = v_Value
End Property

Public Property Let PaymentDate(ByVal v_Value As Date)
    SaveDDPOIssue.PaymentDt = v_Value
End Property
Public Property Let CancelDate(ByVal v_Value As Date)
    SaveDDPOIssue.CancelDt = v_Value
End Property
Public Property Let CancelCharges(ByVal v_Value As Single)
    SaveDDPOIssue.CancelCharges = v_Value
End Property

Public Property Let SecurityNo(ByVal v_Value As Double)
    SaveDDPOIssue.SecurityNo = v_Value
End Property

Public Property Let DD_PO(ByVal v_Value As String)
    SaveDDPOIssue.DD_PO = v_Value
End Property

Public Property Let Cash_Transfer(ByVal v_Value As String)
    SaveDDPOIssue.TransType = v_Value
End Property
    
Public Property Let DDPONo(ByVal v_Value As Double)
    SaveDDPOIssue.DDPONo = v_Value
End Property

Public Property Let IssueDate(ByVal v_Value As Date)
    SaveDDPOIssue.DDDate = v_Value
End Property

Public Property Let InFavourOf(ByVal v_Value As String)
    SaveDDPOIssue.FavourOF = v_Value
End Property

Public Property Let DDAmount(ByVal v_Value As Single)
    SaveDDPOIssue.DDAmount = v_Value
End Property

Public Property Let DepositorName(ByVal v_Value As String)
    SaveDDPOIssue.DepositorName = v_Value
End Property
    
Public Property Let CommissionCharges(ByVal v_Value As Single)
    SaveDDPOIssue.Commission = v_Value
End Property

Public Property Let City(ByVal v_Value As String)
    SaveDDPOIssue.CityName = v_Value
End Property

Public Property Let DDPOStatus(ByVal v_Value As String)
    SaveDDPOIssue.Status = v_Value
End Property

Public Property Let AcNo(ByVal v_Value As Long)
    SaveDDPOIssue.AcNo = v_Value
End Property
Public Property Let AcType(ByVal v_Value As String)
    SaveDDPOIssue.AcType = v_Value
End Property
Public Property Let ISDelete(ByVal v_Value As String)
    SaveDDPOIssue.ISdeleted = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveDDPOIssue.UserName = v_Value
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveDDPOIssue.TerminalName = v_Value
End Property
Public Property Let InsertModDate(ByVal v_Value As Date)
    SaveDDPOIssue.InsertModifyDate = v_Value
End Property

Public Sub DDPOPaymentData(ByRef objDatasource As ADODB.Connection, _
                         ByVal Recordstate As DBStateDD)
Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertPOPayment Then
            .CommandText = "kccbank.PACK_DDPO_Data.InsertPOPayment"
        End If
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
    
        If Recordstate = DBInsertPOPayment Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.PaymentDt)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveDDPOIssue.AcType)
            If SaveDDPOIssue.AcNo = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.AcNo)
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.PaymentMode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
            
            .Execute
        End If
    End With
    Set objCommand = Nothing
End Sub

Public Sub POStopPaymentData(ByRef objDatasource As ADODB.Connection)
Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccbank.PACK_DDPO_Data.InsertDDPOStopPay"
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
    
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.CancelDt)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.CancelCharges)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
        
        .Execute
    End With
    Set objCommand = Nothing

End Sub

Public Sub POStopPaymentRevData(ByRef objDatasource As ADODB.Connection)
Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccbank.PACK_DDPO_Data.InsertPOStopPayRev"
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
    
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.CancelDt)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.CancelCharges)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
        
        .Execute
    End With
    Set objCommand = Nothing
End Sub


Public Sub DDPORevalidationData(ByRef objDatasource As ADODB.Connection)
Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccbank.PACK_DDPO_Data.InsertDDPORevalid"
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
    
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.DD_PO)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.CancelDt)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.CancelCharges)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
        
        .Execute
    End With
    Set objCommand = Nothing

End Sub

Public Sub DDPOCancellationData(ByRef objDatasource As ADODB.Connection)
Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccbank.PACK_DDPO_Data.InsertDDPOCancel"
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
    
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.DD_PO)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.CancelDt)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.CancelCharges)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveDDPOIssue.AcType)
        If SaveDDPOIssue.AcNo = Empty Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.AcNo)
        End If
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
        
        .Execute
    End With
    Set objCommand = Nothing
End Sub

Public Sub DDPOIssueData(ByRef objDatasource As ADODB.Connection, _
                         ByVal Recordstate As DBStateDD)
''To pass on values at the time of Deletion/Insertion/Modification

Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertDD Then
            .CommandText = "kccbank.PACK_DDPO_Data.InsertDDPOIssue"
        ElseIf Recordstate = DBModifyDD Then
            .CommandText = "kccbank.PACK_DDPO_Data.ModifyDDPOIssue"
        ElseIf Recordstate = DBDeleteDD Then
            .CommandText = "kccbank.PACK_DDPO_Data.DeleteDDPO"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If Recordstate = DBInsertDD Or Recordstate = DBModifyDD Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.DD_PO)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.TransType)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, SaveDDPOIssue.FavourOF)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.DDAmount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.DepositorName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveDDPOIssue.CityName)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.Commission)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveDDPOIssue.AcType)
            If SaveDDPOIssue.AcNo = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.AcNo)
            End If
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.Status)
        End If
        
        If Recordstate = DBDeleteDD Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveDDPOIssue.SecurityNo)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.ISdeleted)
        End If
        
        If Recordstate = DBInsertDD Or Recordstate = DBModifyDD Or Recordstate = DBDeleteDD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
        End If
        
        .Execute
        
    End With
    Set objCommand = Nothing
End Sub

'
'Public Function GetFDMasterDetails(ByRef objDatasource As ADODB.Connection) _
'                                   As ADODB.Recordset
''Purpose        :- This function will get Master Details corresponding to given FD No
''                  and populate them in required fields
'
'    Dim objCommand As New ADODB.Command
'
'    With objCommand
'        .CommandText = "kccbank.Pack_FDDetails_Data.GetFDMasterDetails()"
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = objDatasource
'        .CommandTimeout = 0
'        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveFD.FDNO)
'        Set GetFDMasterDetails = .Execute
'    End With
'    Set objCommand = Nothing
'End Function
'
'
'Public Function GetFDDetails(ByRef objDatasource As ADODB.Connection) _
'                               As ADODB.Recordset
''Purpose        :- This function will get Details corresponding to given FD No
''                  and populate them in required fields
'
'    Dim objCommand As New ADODB.Command
'
'    With objCommand
'        .CommandText = "kccbank.Pack_FDDetails_Data.GetFDDetails()"
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = objDatasource
'        .CommandTimeout = 0
'        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveFD.FDNO)
'        Set GetFDDetails = .Execute
'    End With
'    Set objCommand = Nothing
'End Function
'
'
