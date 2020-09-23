VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************
'Class Name         :- CJV.Cls
'Pupose             :- Class created for Journal Voucher Entry
'Referred dll       :-
'Date of Creation   :- 1st April'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************


'Variables declared for saving on Deletion,Insertion,Modifcation
Private Type JVData
    TrnKey              As Long
    Trn_Type            As String
    VchNo               As Single
    VchDate             As Date
    GLAcNo              As String
    PLAcNo              As String
    Dr_Cr               As String
    Amount              As Double
    Narration           As String
    AdviceNo            As String
    CashKind            As String
    IS_Agri             As String
    AdviceDate          As Variant
    ISdeleted           As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
End Type

'Enumeration to define different stages of insert/modify and delete
Public Enum DBStateJV
    DBInsertD = 1
    DBModifyD = 2
    DBDeleteD = 3
    DBInsertC = 4
    DBModifyC = 5
    DBDeleteC = 6
    DBDeleteVch = 7
End Enum

Private m_TrnCurValue       As Long    'Variable to save current transaction key
Private SaveJV              As JVData


Public Property Let isAgriLoan(ByVal v_Value As String)
    SaveJV.IS_Agri = v_Value
End Property

Public Property Let CashKind(ByVal v_Value As String)
    SaveJV.CashKind = v_Value
End Property
Public Property Let TrnKey(ByVal v_Value As Long)
    SaveJV.TrnKey = v_Value
End Property
Public Property Let TrnType(ByVal v_Value As String)
    SaveJV.Trn_Type = v_Value
End Property
Public Property Let VchNo(ByVal v_Value As Single)
    SaveJV.VchNo = v_Value
End Property
Public Property Let VchDate(ByVal v_Value As Date)
    SaveJV.VchDate = v_Value
End Property
Public Property Let GLAcNo(ByVal v_Value As String)
    SaveJV.GLAcNo = v_Value
End Property
Public Property Let PLAcNo(ByVal v_Value As String)
    SaveJV.PLAcNo = v_Value
End Property
Public Property Let Dr_Cr(ByVal v_Value As String)
    SaveJV.Dr_Cr = v_Value
End Property
Public Property Let VchAmount(ByVal v_Value As Double)
    SaveJV.Amount = v_Value
End Property
Public Property Let Narration(ByVal v_Value As String)
    SaveJV.Narration = v_Value
End Property
Public Property Let AdviceNo(ByVal v_Value As String)
    SaveJV.AdviceNo = v_Value
End Property
Public Property Let AdvDate(ByVal v_Value As Date)
    SaveJV.AdviceDate = v_Value
End Property
Public Property Let ISdeleted(ByVal v_Value As String)
    SaveJV.ISdeleted = v_Value
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveJV.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveJV.UserName = v_Value
End Property
Public Property Let InsertModDate(ByVal v_Value As Date)
    SaveJV.InsertModifyDate = v_Value
End Property
Public Property Get TrnCurrentValue() As Long
     TrnCurrentValue = m_TrnCurValue
End Property


Public Function GetGLAccountHeadList(ByRef objDatasource As ADODB.Connection) _
                                     As ADODB.Recordset
'Purpose        :- This function will get list of all GL Account Heads to be filled in
'                  relevant combo box

    Dim objCommand As New ADODB.Command
        
    With objCommand
        .CommandText = "kccbank.PACK_JVoucher_Data.GetGLAccounts()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GetGLAccountHeadList = .Execute
    End With
    Set objCommand = Nothing

End Function



Public Sub JVoucherData(ByRef objDatasource As ADODB.Connection, _
                        ByVal Recordstate As DBStateJV)
''To pass on values at the time of Deletion/Insertion/Modification

Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertD Then
            .CommandText = "kccbank.PACK_JVoucher_Data.InsertJVDebit"
        ElseIf Recordstate = DBModifyD Then
            .CommandText = "kccbank.PACK_JVoucher_Data.UpdateJVDebit"
        ElseIf Recordstate = DBDeleteD Then
            .CommandText = "kccbank.PACK_JVoucher_Data.DeleteJVDebit"
        ElseIf Recordstate = DBInsertC Then
            .CommandText = "kccbank.PACK_JVoucher_Data.InsertJVCredit"
        ElseIf Recordstate = DBModifyC Then
            .CommandText = "kccbank.PACK_JVoucher_Data.UpdateJVCredit"
        ElseIf Recordstate = DBDeleteC Then
            .CommandText = "kccbank.PACK_JVoucher_Data.DeleteJVCredit"
        ElseIf Recordstate = DBDeleteVch Then
            .CommandText = "kccbank.PACK_JVoucher_Data.DeleteJV"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If Recordstate = DBDeleteC Or Recordstate = DBDeleteD Or Recordstate = DBModifyC Or Recordstate = DBModifyD Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveJV.TrnKey)
        End If
        
        If Recordstate = DBInsertC Or Recordstate = DBInsertD Or Recordstate = DBModifyC Or Recordstate = DBModifyD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveJV.Trn_Type)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveJV.VchNo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveJV.VchDate)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveJV.GLAcNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveJV.PLAcNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveJV.Dr_Cr)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveJV.Amount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 80, SaveJV.Narration)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveJV.AdviceNo)
            If SaveJV.AdviceDate = Empty Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveJV.AdviceDate)
            End If
        End If
        If Recordstate = DBInsertD Or Recordstate = DBModifyD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveJV.CashKind)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveJV.IS_Agri)
        End If
        If Recordstate = DBInsertC Or Recordstate = DBModifyC Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveJV.IS_Agri)
        End If
        If Recordstate = DBInsertC Or Recordstate = DBInsertD Or Recordstate = DBModifyC Or Recordstate = DBModifyD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveJV.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveJV.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveJV.InsertModifyDate)
        End If
                
        If Recordstate = DBDeleteVch Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveJV.VchNo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveJV.VchDate)
        End If
        If Recordstate = DBDeleteC Or Recordstate = DBDeleteD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveJV.ISdeleted)
        End If
        If Recordstate = DBDeleteC Or Recordstate = DBDeleteD Or Recordstate = DBDeleteVch Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveJV.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveJV.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveJV.InsertModifyDate)
        End If
        
        If Recordstate = DBInsertC Or Recordstate = DBInsertD Or Recordstate = DBModifyC Or Recordstate = DBModifyD Then
            .Parameters.Append .CreateParameter("param2", adNumeric, adParamOutput)
        End If
        
        .Execute
        
        If Recordstate = DBInsertC Or Recordstate = DBInsertD Then
            If IsNull(objCommand("param2").Value) = False Then
                m_TrnCurValue = CLng(objCommand("param2").Value)
            End If
        End If
    End With
    Set objCommand = Nothing
End Sub


