VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccHeadMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************
'Class Name         :- KCCBAccHeadMst.Cls
'Pupose             :- Class created for Account Head Master
'Referred dll       :-
'Date of Creation   :- 22nd March'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************


'Variables declared for saving on Deletion,Insertion,Modifcation
Private Type AccHeadData
    AccKey              As Long
    GroupKey            As Long
    GroupCode           As String
    ParentCode          As String
    ISDeleted           As String
    AccountCode         As String
    AccDescription      As String
    Subgroup            As String
    PL                  As String
    OpeningBalance      As Single
    AgriLoan            As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
End Type

Private Type AccHeadBalanceData
    AccountCode         As String
    OpeningBalance      As Single
    OpeningYear         As Long
    ISDeleted           As String
End Type

Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
    DBCheckUnique = 4
    DBCheckRecord = 5
    DBInsertBal = 6
    DBModifyBal = 7
    DBDeleteBal = 7
End Enum
'    MCL                 As String
'    KharifKind          As Single
'    KharifCash          As Single
'    RabiKind            As Single
'    RabiCash            As Single
'    MAxKind             As Single
'    MaxCash             As Single

Private m_blnUnique         As Boolean
Private m_AccountCurValue   As Long
Private SaveAcc             As AccHeadData
Private SaveAccBalance      As AccHeadBalanceData
Public Property Let AcCodeBalance(ByVal v_Value As String)
    SaveAccBalance.AccountCode = v_Value
End Property
Public Property Let ISDeletedBalance(ByVal v_Value As String)
    SaveAccBalance.ISDeleted = v_Value
End Property
Public Property Let OpenYearBalance(ByVal v_Value As Long)
    SaveAccBalance.OpeningYear = v_Value
End Property
Public Property Let OpenBalBalance(ByVal v_Value As Single)
    SaveAccBalance.OpeningBalance = v_Value
End Property

Public Property Let AgriLoan(ByVal v_Value As String)
    SaveAcc.AgriLoan = v_Value
End Property
'
'Public Property Let MCL(ByVal v_Value As String)
'    SaveAcc.MCL = v_Value
'End Property
'
'Public Property Let KharifKind(ByVal v_Value As Single)
'    SaveAcc.KharifKind = v_Value
'End Property
'
'Public Property Let KharifCash(ByVal v_Value As Single)
'    SaveAcc.KharifCash = v_Value
'End Property
'
'Public Property Let RabiKind(ByVal v_Value As Single)
'    SaveAcc.RabiKind = v_Value
'End Property
'
'Public Property Let RabiCash(ByVal v_Value As Single)
'    SaveAcc.RabiCash = v_Value
'End Property
'
'Public Property Let MAxKind(ByVal v_Value As Single)
'    SaveAcc.MAxKind = v_Value
'End Property
'
'Public Property Let MaxCash(ByVal v_Value As Single)
'    SaveAcc.MaxCash = v_Value
'End Property
Public Property Let AccKey(ByVal v_Value As Long)
    SaveAcc.AccKey = v_Value
End Property

Public Property Let GroupKey(ByVal v_Value As Long)
    SaveAcc.GroupKey = v_Value
End Property

Public Property Let ParentNo(ByVal v_Value As String)
    SaveAcc.ParentCode = v_Value
End Property

Public Property Let AccDescription(ByVal v_Value As String)
    SaveAcc.AccDescription = v_Value
End Property

Public Property Let GrpCode(ByVal v_Value As String)
    SaveAcc.GroupCode = v_Value
End Property

Public Property Let AccCode(ByVal v_Value As String)
    SaveAcc.AccountCode = v_Value
End Property
Public Property Let Subgroup(ByVal v_Value As String)
    SaveAcc.Subgroup = v_Value
End Property
Public Property Let ISDeleted(ByVal v_Value As String)
    SaveAcc.ISDeleted = v_Value
End Property

Public Property Let ISPL(ByVal v_Value As String)
    SaveAcc.PL = v_Value
End Property

Public Property Let OpeningBalance(ByVal v_Value As Single)
    SaveAcc.OpeningBalance = v_Value
End Property

Public Property Let TerminalName(ByVal v_Value As String)
    SaveAcc.TerminalName = v_Value
End Property

Public Property Let UserName(ByVal v_Value As String)
    SaveAcc.UserName = v_Value
End Property

Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveAcc.InsertModifyDate = v_Value
End Property

Public Property Get IsUnique() As Boolean
    IsUnique = m_blnUnique
End Property

Public Property Get AccountCurrentValue() As Long
     AccountCurrentValue = m_AccountCurValue
End Property


Public Sub AccHeadData(ByRef objDatasource As ADODB.Connection, _
                       ByVal Recordstate As DBState)
''To pass on values at the time of Deletion/Insertion/Modification

Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsert Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.InsertAccHead"
        ElseIf Recordstate = DBModify Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.UpdateAccHead"
        ElseIf Recordstate = DBDelete Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.DeleteAccHeadData"
        ElseIf Recordstate = DBCheckUnique Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.CheckDupliAccHead"
        ElseIf Recordstate = DBCheckRecord Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.CheckRecord"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        If Recordstate = DBCheckRecord Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.AccKey)
            .Parameters.Append .CreateParameter("param1", adNumeric, adParamOutput)
        End If
        
        If Recordstate = DBCheckUnique Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 60, SaveAcc.AccDescription)
            .Parameters.Append .CreateParameter("param1", adNumeric, adParamOutput)
        End If
        
        If Recordstate = DBModify Or Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.AccKey)
        End If
        
        If Recordstate = DBInsert Or Recordstate = DBModify Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.GroupKey)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveAcc.GroupCode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveAcc.ParentCode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 5, SaveAcc.AccountCode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 60, SaveAcc.AccDescription)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveAcc.Subgroup)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveAcc.PL)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.OpeningBalance)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveAcc.AgriLoan)
'            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveAcc.MCL)
'            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.KharifCash)
'            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.KharifKind)
'            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.RabiCash)
'            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.RabiKind)
'            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.MaxCash)
'            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAcc.MAxKind)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveAcc.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveAcc.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveAcc.InsertModifyDate)
        End If
        If Recordstate = DBInsert Then
            .Parameters.Append .CreateParameter("param2", adNumeric, adParamOutput)
        End If
        If Recordstate = DBDelete Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveAcc.ISDeleted)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveAcc.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveAcc.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveAcc.InsertModifyDate)
        End If
        .Execute
        If Recordstate = DBCheckRecord Or Recordstate = DBCheckUnique Then
            If CLng(objCommand("param1").Value) > 0 Then
                m_blnUnique = False
            Else
                m_blnUnique = True
            End If
        End If
        If Recordstate = DBInsert Then
            If IsNull(objCommand("param2").Value) = False Then
                m_AccountCurValue = CLng(objCommand("param2").Value)
            End If
        End If
      '  If Recordstate = DBModify Then MsgBox IIf(IsNull(objCommand("param2").Value), "Null", "Not Null")
            
            
        
    End With
    Set objCommand = Nothing
End Sub

Public Sub AccHeadBalanceData(ByRef objDatasource As ADODB.Connection, _
                              ByVal Recordstate As DBState)
''To pass on values at the time of Deletion/Insertion/Modification

Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertBal Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.InsertAccHeadBalance"
        ElseIf Recordstate = DBModifyBal Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.UpdateAccHeadBalance"
        ElseIf Recordstate = DBDeleteBal Then
            .CommandText = "kccbank.PACK_AcHeadBranch_Data.DeleteAccHeadBalance"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        If Recordstate = DBModifyBal Or Recordstate = DBInsertBal Or Recordstate = DBDeleteBal Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveAccBalance.AccountCode)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAccBalance.OpeningYear)
        End If
        
        If Recordstate = DBModifyBal Or Recordstate = DBInsertBal Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveAccBalance.OpeningBalance)
        End If
        
        .Execute
    End With
    Set objCommand = Nothing
End Sub

Public Function GetGroupList(ByRef objDatasource As ADODB.Connection, _
                              ByVal hWnd As Long) _
                              As ADODB.Recordset
'Purpose        :- This function will get list of all subgroups to be filled in
'                  relevant Combo box

    Dim objCommand As New ADODB.Command
    Dim objRecordset As ADODB.Recordset
    Dim objOLEServer As New KccbUtils.OLERemoteServer
        
    With objCommand
        .CommandText = "ho.PACK_ACHEADBranch_Data.GetGroupsAccHead()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set objRecordset = .Execute
        If objRecordset.RecordCount > 0 Then
            Set GetGroupList = objRecordset
        Else
            MsgBox "No Groups exists in database.", vbOKOnly + vbCritical, "KCCB Error"
            objOLEServer.ShutdownRemoteOLEServer (hWnd)
        End If
    End With
    Set objCommand = Nothing
    Set objRecordset = Nothing
    Set objOLEServer = Nothing
End Function

Public Function GetAccountHeadList(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
'Purpose        :- This function will get list of all Account Heads to be filled in
'                  relevant List view

    Dim objCommand As New ADODB.Command
'    Dim objOLEServer As New KccbUtils.OLERemoteServer
        
    With objCommand
        .CommandText = "kccbank.Pack_AcHeadBranch_Data.GetAccHeadData()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GetAccountHeadList = .Execute
    End With
    Set objCommand = Nothing
'    Set objOLEServer = Nothing

End Function
