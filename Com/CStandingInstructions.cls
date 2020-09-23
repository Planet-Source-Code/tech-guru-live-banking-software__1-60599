VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStandingInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************
'Class Name         :- CStandingInstructions.Cls
'Pupose             :- Class created for Standing Instructions
'Referred dll       :-
'Date of Creation   :- 23nd March'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************


'Variables declared for saving on Deletion,Insertion,Modifcation
Private Type StandingData
    SI_Key              As Long
    Type_SI             As String
    SI_Date             As Date
    AC_Type             As String
    AC_No               As Long
    SIAmount            As Double
    Cheque_No           As Long
    Cheque_Date         As Variant
    Period_Month        As Long
    SIDay               As Long
    All_WorkingDay      As String
    Valid_Upto          As Long
    Bank_Charges        As Double
    TRF_Type            As String
    TRF_No              As Long
    FD_No               As String
    ISdeleted           As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
    StartDate           As Date
    Is_Reversed         As String
    ReversalDate        As Date
End Type

Public Enum DBStateS
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
    DBReverse = 4
End Enum

Private m_SIKeyCurValue As Long
Private SaveStandingData    As StandingData

Public Property Get SIKeyCurrentValue() As Long
     SIKeyCurrentValue = m_SIKeyCurValue
End Property

Public Property Let IsReverse(ByVal v_Value As String)
    SaveStandingData.Is_Reversed = v_Value
End Property

Public Property Let ReverseDate(ByVal v_Value As Date)
    SaveStandingData.ReversalDate = v_Value
End Property
Public Property Let IMDDate(ByVal v_Value As Date)
    SaveStandingData.InsertModifyDate = v_Value
End Property

Public Property Let StartDate(ByVal v_Value As Date)
    SaveStandingData.StartDate = v_Value
End Property

Public Property Let UserName(ByVal v_Value As String)
    SaveStandingData.UserName = v_Value
End Property

Public Property Let Terminal(ByVal v_Value As String)
    SaveStandingData.TerminalName = v_Value
End Property

Public Property Let FDNo(ByVal v_Value As String)
    SaveStandingData.FD_No = v_Value
End Property

Public Property Let TRFNo(ByVal v_Value As Long)
    SaveStandingData.TRF_No = v_Value
End Property

Public Property Let TRFType(ByVal v_Value As String)
    SaveStandingData.TRF_Type = v_Value
End Property

Public Property Let BankCharge(ByVal v_Value As Double)
    SaveStandingData.Bank_Charges = v_Value
End Property
    
Public Property Let ValidUpto(ByVal v_Value As Long)
    SaveStandingData.Valid_Upto = v_Value
End Property

Public Property Let AllWork(ByVal v_Value As String)
    SaveStandingData.All_WorkingDay = v_Value
End Property
    
Public Property Let SIDay(ByVal v_Value As Long)
    SaveStandingData.SIDay = v_Value
End Property

Public Property Let PeriodMonth(ByVal v_Value As Long)
    SaveStandingData.Period_Month = v_Value
End Property

Public Property Let ChequeDt(ByVal v_Value As Variant)
    SaveStandingData.Cheque_Date = v_Value
End Property

Public Property Let ChequeNo(ByVal v_Value As Long)
    SaveStandingData.Cheque_No = v_Value
End Property

Public Property Let ACNo(ByVal v_Value As Long)
    SaveStandingData.AC_No = v_Value
End Property

Public Property Let ACAmount(ByVal v_Value As Double)
    SaveStandingData.SIAmount = v_Value
End Property

Public Property Let ACType(ByVal v_Value As String)
    SaveStandingData.AC_Type = v_Value
End Property

Public Property Let SIDate(ByVal v_Value As Date)
    SaveStandingData.SI_Date = v_Value
End Property

Public Property Let SIKey(ByVal v_Value As Long)
    SaveStandingData.SI_Key = v_Value
End Property

Public Property Let TypeSI(ByVal v_Value As String)
    SaveStandingData.Type_SI = v_Value
End Property
Public Property Let ISDelete(ByVal v_Value As String)
    SaveStandingData.ISdeleted = v_Value
End Property
Public Sub StandingInstructions(ByRef objDatasource As ADODB.Connection, _
                                ByVal RecordState As DBStateS)
                       
''To pass on values at the time of Deletion/Insertion/Modification
    Dim objCommand As New ADODB.Command
    With objCommand
        If RecordState = DBInsert Then
            .CommandText = "Kccbank.PACK_StandingInstruct_Data.InsertStanding"
        ElseIf RecordState = DBModify Then
            .CommandText = "Kccbank.PACK_StandingInstruct_Data.UpdateStanding"
        ElseIf RecordState = DBDelete Then
            .CommandText = "Kccbank.PACK_StandingInstruct_Data.DeleteStanding"
        ElseIf RecordState = DBReverse Then
            .CommandText = "Kccbank.PACK_StandingInstruct_Data.UpdateReversal"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If RecordState = DBModify Or RecordState = DBDelete Or RecordState = DBReverse Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.SI_Key)
        End If
        
        If RecordState = DBReverse Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveStandingData.Is_Reversed)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveStandingData.ReversalDate)
        End If
        
        If RecordState = DBInsert Or RecordState = DBModify Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveStandingData.Type_SI)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveStandingData.SI_Date)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveStandingData.AC_Type)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.AC_No)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.SIAmount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.Cheque_No)
            If SaveStandingData.Cheque_Date = Empty Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , Null)
            Else
                .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveStandingData.Cheque_Date)
            End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.Period_Month)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.SIDay)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveStandingData.All_WorkingDay)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.Valid_Upto)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.Bank_Charges)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveStandingData.TRF_Type)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveStandingData.TRF_No)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveStandingData.FD_No)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveStandingData.StartDate)
        End If
        If RecordState = DBDelete Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveStandingData.ISdeleted)
        End If
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveStandingData.TerminalName)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveStandingData.UserName)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveStandingData.InsertModifyDate)

        If RecordState = DBInsert Or RecordState = DBModify Then
            .Parameters.Append .CreateParameter("param1", adNumeric, adParamOutput)
        End If
        .Execute
        If RecordState = DBInsert Then
            m_SIKeyCurValue = CLng(objCommand("param1").Value)
        End If
    End With
    Set objCommand = Nothing
End Sub


Public Function GetInstructionList(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
'Purpose        :- This function will get list of all instructions to be filled in
'                  relevant List view

    Dim objCommand As New ADODB.Command
            
    With objCommand
        .CommandText = "kccbank.Pack_StandingInstruct_Data.GetStandingAll()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GetInstructionList = .Execute
    End With
    Set objCommand = Nothing

End Function

Public Function GetInstructionListSpecific(ByRef objDatasource As ADODB.Connection, _
                                            ByVal strTypeSI As String) _
                                            As ADODB.Recordset
'Purpose        :- This function will get list of all instructions to be filled in
'                  relevant List view

    Dim objCommand As New ADODB.Command
            
    With objCommand
        .CommandText = "kccbank.Pack_StandingInstruct_Data.GetStanding()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, strTypeSI)
        Set GetInstructionListSpecific = .Execute
    End With
    Set objCommand = Nothing

End Function

