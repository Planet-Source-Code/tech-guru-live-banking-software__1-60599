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
    DDAmount            As Double
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
    PaymentBank         As String
    PaymentBranch       As String
    Remarks             As String
'    RevalidCharge       As Single
'    ISStopPayment       As String
'    StopPayCharge       As Single
'    StopPayRevCharge    As Single
    ISdeleted           As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
    TrnKey              As Single
    Form6Date           As Variant
End Type

'Enumeration to define different stages of insert/modify and delete
Public Enum DBStateDD
    DBInsertDD = 1
    DBModifyDD = 2
    DBDeleteDD = 3
    DBInsertPOPayment = 4
    DBInsertDDPay = 5
    DBModifyDDPay = 6
    DBDeleteDDPay = 7
    DBInsertForm6 = 8
    DBModifyForm6 = 9
    DBDeleteForm6 = 10
End Enum

Private SaveDDPOIssue                As DDPOIssueData
Private m_TrnKey                     As Single
Public Property Get TrnKey() As Single
    TrnKey = m_TrnKey
End Property
Public Property Let TrnKey(ByVal v_Value As Single)
    SaveDDPOIssue.TrnKey = v_Value
End Property
Public Property Let Form6Date(ByVal v_Value As Variant)
    SaveDDPOIssue.Form6Date = v_Value
End Property

Public Property Let Paymentremarks(ByVal v_Value As String)
    SaveDDPOIssue.Remarks = v_Value
End Property
Public Property Let PaymentBank(ByVal v_Value As String)
    SaveDDPOIssue.PaymentBank = v_Value
End Property

Public Property Let PaymentBranch(ByVal v_Value As String)
    SaveDDPOIssue.PaymentBranch = v_Value
End Property

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

Public Property Let DDAmount(ByVal v_Value As Double)
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

Public Sub Form7Details(ByRef objDatasource As ADODB.Connection)

Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "kccbank.PACK_DDPayment_Data.UpdateForm7Details"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
    
                    
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.CancelDt)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.PaymentMode)
        If SaveDDPOIssue.PaymentMode = "P" Then
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.PaymentDt)
        Else
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , Null)
        End If
        
        .Execute
        
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
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.DD_PO)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)
        
        .Execute
    End With
    Set objCommand = Nothing

End Sub

Public Sub DDPaymentData(ByRef objDatasource As ADODB.Connection, _
                         ByVal Recordstate As DBStateDD)
Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertDDPay Then
            .CommandText = "kccbank.PACK_DDPayment_Data.InsertDDPayment"
        ElseIf Recordstate = DBModifyDDPay Then
            .CommandText = "kccbank.PACK_DDPayment_Data.ModifyDDPayment"
        ElseIf Recordstate = DBDeleteDDPay Then
            .CommandText = "kccbank.PACK_DDPayment_Data.DeleteDDPayment"
        End If
        
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If Recordstate = DBModifyDDPay Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.TrnKey)
        End If
        If Recordstate = DBInsertDDPay Or Recordstate = DBModifyDDPay Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.SecurityNo)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.DDAmount)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.PaymentDt)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, SaveDDPOIssue.FavourOF)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveDDPOIssue.PaymentBank)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.PaymentBranch)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveDDPOIssue.CityName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.Remarks)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.PaymentMode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveDDPOIssue.AcType)
            If SaveDDPOIssue.AcNo = Empty Then
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.AcNo)
            End If
                    
        End If
                                                                
        If Recordstate = DBDeleteDDPay Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.TrnKey)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.ISdeleted)
        End If
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)

        If Recordstate = DBInsertDDPay Then
            .Parameters.Append .CreateParameter("param1", adNumeric, adParamInputOutput)
        End If
        
        .Execute
        
        If Recordstate = DBInsertDDPay Then
            If IsNull(objCommand("param1").Value) = False Then
                m_TrnKey = CLng(objCommand("param1").Value)
            End If
        End If
        
    End With
    Set objCommand = Nothing

End Sub

Public Sub Form6Data(ByRef objDatasource As ADODB.Connection, _
                         ByVal Recordstate As DBStateDD)
Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertForm6 Then
            .CommandText = "kccbank.PACK_DDPayment_Data.InsertForm6"
        ElseIf Recordstate = DBModifyForm6 Then
            .CommandText = "kccbank.PACK_DDPayment_Data.ModifyForm6"
        ElseIf Recordstate = DBDeleteForm6 Then
            .CommandText = "kccbank.PACK_DDPayment_Data.DeleteForm6"
        End If
        
    
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If Recordstate = DBModifyForm6 Or Recordstate = DBDeleteForm6 Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.TrnKey)
        End If
        If Recordstate = DBInsertForm6 Or Recordstate = DBModifyForm6 Then
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.CancelDt)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, SaveDDPOIssue.DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.DDDate)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.DDAmount)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.PaymentDt)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.FavourOF)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveDDPOIssue.PaymentBank)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.PaymentBranch)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveDDPOIssue.CityName)
        End If
        If Recordstate = DBInsertForm6 Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.PaymentMode)
            If SaveDDPOIssue.Form6Date = Empty Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.Form6Date)
            End If
            
        End If
        If Recordstate = DBDeleteForm6 Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.ISdeleted)
        End If
        
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveDDPOIssue.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveDDPOIssue.InsertModifyDate)

        If Recordstate = DBInsertForm6 Then
            .Parameters.Append .CreateParameter("param1", adNumeric, adParamInputOutput)
        End If
        
        .Execute
        
        If Recordstate = DBInsertForm6 Then
            If IsNull(objCommand("param1").Value) = False Then
                m_TrnKey = CLng(objCommand("param1").Value)
            End If
        End If
        
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
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.DDAmount)
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
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.SecurityNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.DD_PO)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveDDPOIssue.TransType)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveDDPOIssue.DDPONo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveDDPOIssue.DDDate)
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
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveDDPOIssue.PaymentBank)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveDDPOIssue.PaymentBranch)
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
