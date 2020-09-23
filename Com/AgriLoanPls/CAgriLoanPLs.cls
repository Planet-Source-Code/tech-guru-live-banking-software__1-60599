VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAgriLoanPLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************
'Class Name         :- KCCBAgriLoanPLs.Cls
'Pupose             :- Class created for creating PL Heads under AgriCultural Loan Heads
'Referred dll       :-
'Date of Creation   :- 16th April'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************


'Variables declared for saving on Deletion,Insertion,Modifcation

Private Type AgriLoanPLData
    AccType             As String
    AccNo               As Long
    Society             As String
    Address1            As String
    Address2            As String
    City                As String
    State               As String
    PinCode             As String
    PhoneNo             As String
    Fax                 As String
    FromMCL             As Long
    ToMCL               As Long
    KharifKind          As Single
    KharifCash          As Single
    RabiKind            As Single
    RabiCash            As Single
    MAxKind             As Single
    MaxCash             As Single
    ISDeleted           As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
End Type

Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum

Private SaveALPLs             As AgriLoanPLData
Public Property Let Fax(ByVal v_Value As String)
    SaveALPLs.Fax = v_Value
End Property
Public Property Let PhoneNo(ByVal v_Value As String)
    SaveALPLs.PhoneNo = v_Value
End Property
Public Property Let PinCode(ByVal v_Value As String)
    SaveALPLs.PinCode = v_Value
End Property
Public Property Let State(ByVal v_Value As String)
    SaveALPLs.State = v_Value
End Property
Public Property Let City(ByVal v_Value As String)
    SaveALPLs.City = v_Value
End Property
Public Property Let Address2(ByVal v_Value As String)
    SaveALPLs.Address2 = v_Value
End Property
Public Property Let Address1(ByVal v_Value As String)
    SaveALPLs.Address1 = v_Value
End Property
Public Property Let TypeOFAccount(ByVal v_Value As String)
    SaveALPLs.AccType = v_Value
End Property
Public Property Let SocietyName(ByVal v_Value As String)
    SaveALPLs.Society = v_Value
End Property
Public Property Let AccountNo(ByVal v_Value As Long)
    SaveALPLs.AccNo = v_Value
End Property
Public Property Let FromMCL(ByVal v_Value As Long)
    SaveALPLs.FromMCL = v_Value
End Property
Public Property Let ToMCL(ByVal v_Value As Long)
    SaveALPLs.ToMCL = v_Value
End Property

Public Property Let KharifKind(ByVal v_Value As Single)
    SaveALPLs.KharifKind = v_Value
End Property

Public Property Let KharifCash(ByVal v_Value As Single)
    SaveALPLs.KharifCash = v_Value
End Property

Public Property Let RabiKind(ByVal v_Value As Single)
    SaveALPLs.RabiKind = v_Value
End Property

Public Property Let RabiCash(ByVal v_Value As Single)
    SaveALPLs.RabiCash = v_Value
End Property

Public Property Let MAxKind(ByVal v_Value As Single)
    SaveALPLs.MAxKind = v_Value
End Property

Public Property Let MaxCash(ByVal v_Value As Single)
    SaveALPLs.MaxCash = v_Value
End Property
Public Property Let ISDeleted(ByVal v_Value As String)
    SaveALPLs.ISDeleted = v_Value
End Property

Public Property Let TerminalName(ByVal v_Value As String)
    SaveALPLs.TerminalName = v_Value
End Property

Public Property Let UserName(ByVal v_Value As String)
    SaveALPLs.UserName = v_Value
End Property

Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveALPLs.InsertModifyDate = v_Value
End Property


Public Sub AgriLoanPLData(ByRef objDatasource As ADODB.Connection, _
                          ByVal Recordstate As DBState)
''To pass on values at the time of Deletion/Insertion/Modification

Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsert Then
            .CommandText = "kccbank.PACK_AgriLoans_Data.InsertALoan"
        ElseIf Recordstate = DBModify Then
            .CommandText = "kccbank.PACK_AgriLoans_Data.UpdateALoan"
        ElseIf Recordstate = DBDelete Then
            .CommandText = "kccbank.PACK_AgriLoans_Data.DeleteALData"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If Recordstate = DBInsert Or Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveALPLs.AccType)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.AccNo)
        End If
        
        If Recordstate = DBInsert Or Recordstate = DBModify Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveALPLs.Society)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveALPLs.Address1)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveALPLs.Address2)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveALPLs.City)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveALPLs.State)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 10, SaveALPLs.PinCode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveALPLs.PhoneNo)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveALPLs.Fax)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.FromMCL)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.ToMCL)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.KharifCash)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.KharifKind)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.RabiCash)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.RabiKind)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.MaxCash)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveALPLs.MAxKind)
        End If
        
        
        If Recordstate = DBInsert Or Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveALPLs.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveALPLs.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveALPLs.InsertModifyDate)
        End If
        
        If Recordstate = DBInsert Then
            .Parameters.Append .CreateParameter("param2", adNumeric, adParamOutput)
        End If
        .Execute
    End With
    Set objCommand = Nothing
End Sub

Public Function GetAgriLoanGLs(ByRef objDatasource As ADODB.Connection, _
                              ByVal hWnd As Long) _
                              As ADODB.Recordset
'Purpose        :- This function will get list of all subgroups to be filled in
'                  relevant Combo box

    Dim objCommand As New ADODB.Command
    Dim objRecordset As ADODB.Recordset
        
    With objCommand
        .CommandText = "ho.PACK_AgriLoans_Data.GetAgriLoanGLHead()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set objRecordset = .Execute
        If objRecordset.RecordCount > 0 Then
            Set GetAgriLoanGLs = objRecordset
        Else
            MsgBox "No Agri. Loan Heads exists in database.", vbOKOnly + vbCritical, "KCCB Error"
        End If
    End With
    Set objCommand = Nothing
    Set objRecordset = Nothing
    
End Function

Public Function GetAgriLoanPLList(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
'Purpose        :- This function will get list of all Account Heads to be filled in
'                  relevant List view

    Dim objCommand As New ADODB.Command
        
    With objCommand
        .CommandText = "kccbank.PACK_AgriLoans_Data.GetALSocData()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GetAgriLoanPLList = .Execute
    End With
    Set objCommand = Nothing

End Function


