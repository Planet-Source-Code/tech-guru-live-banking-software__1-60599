VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFDDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************
'Class Name         :- CFDDtl.Cls
'Pupose             :- Class created for FD Details Entry
'Referred dll       :-
'Date of Creation   :- 6st April'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************


'Variables declared for saving on Deletion,Insertion,Modifcation
Private Type FDDtlData
    FDNO                As String
    FDDate              As Date
    FDAmount            As Single
    PeriodCrt           As String
    Period              As Long
    ROI                 As Single
    EitherS             As String
    FDMode              As String
    ACType              As String
    ACNo                As Long
    MatureDate          As Date
    MatureAmt           As Single
    InterestAmt         As Single
    Old_New             As String
    FD_Lien             As String
    FD_Tds              As String
    fd_status           As Long
    ISdeleted           As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
    InttPayable         As Double
End Type

'Enumeration to define different stages of insert/modify and delete
Public Enum DBStateFD
    DBInsertFD = 1
    DBModifyFD = 2
    DBDeleteFD = 3
End Enum

Private SaveFD              As FDDtlData
Public Property Let Intt_Payable(ByVal v_Value As Double)
    SaveFD.InttPayable = v_Value
End Property


Public Property Let FD_Tds(ByVal v_Value As String)
    SaveFD.FD_Tds = v_Value
End Property

Public Property Let fd_status(ByVal v_Value As Long)
    SaveFD.fd_status = v_Value
End Property

Public Property Let FD_Lien(ByVal v_Value As String)
    SaveFD.FD_Lien = v_Value
End Property

Public Property Let UserName(ByVal v_Value As String)
    SaveFD.UserName = v_Value
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveFD.TerminalName = v_Value
End Property
Public Property Let ISDelete(ByVal v_Value As String)
    SaveFD.ISdeleted = v_Value
End Property
Public Property Let Old_New(ByVal v_Value As String)
    SaveFD.Old_New = v_Value
End Property
Public Property Let InterestAmt(ByVal v_Value As Single)
    SaveFD.InterestAmt = v_Value
End Property
Public Property Let MatureAmt(ByVal v_Value As Single)
    SaveFD.MatureAmt = v_Value
End Property
Public Property Let MatureDt(ByVal v_Value As Date)
    SaveFD.MatureDate = v_Value
End Property
Public Property Let ACNo(ByVal v_Value As Long)
    SaveFD.ACNo = v_Value
End Property
Public Property Let ACType(ByVal v_Value As String)
    SaveFD.ACType = v_Value
End Property
Public Property Let FDMode(ByVal v_Value As String)
    SaveFD.FDMode = v_Value
End Property
Public Property Let Either(ByVal v_Value As String)
    SaveFD.EitherS = v_Value
End Property
Public Property Let ROI(ByVal v_Value As Single)
    SaveFD.ROI = v_Value
End Property
Public Property Let Period(ByVal v_Value As Long)
    SaveFD.Period = v_Value
End Property
Public Property Let PCrt(ByVal v_Value As String)
    SaveFD.PeriodCrt = v_Value
End Property
Public Property Let FDNO(ByVal v_Value As String)
    SaveFD.FDNO = v_Value
End Property
Public Property Let FDDate(ByVal v_Value As Date)
    SaveFD.FDDate = v_Value
End Property
Public Property Let FDAmount(ByVal v_Value As Single)
    SaveFD.FDAmount = v_Value
End Property
Public Property Let InsertModDate(ByVal v_Value As Date)
    SaveFD.InsertModifyDate = v_Value
End Property
'Public Property Get TrnCurrentValue() As Long
'     TrnCurrentValue = m_TrnCurValue
'End Property

Public Sub FDData(ByRef objDatasource As ADODB.Connection, _
                  ByVal Recordstate As DBStateFD)
''To pass on values at the time of Deletion/Insertion/Modification

Dim objCommand As New ADODB.Command
    With objCommand
        If Recordstate = DBInsertFD Then
            .CommandText = "kccbank.PACK_FDDetails_Data.InsertFDDetails"
        ElseIf Recordstate = DBModifyFD Then
            .CommandText = "kccbank.PACK_FDDetails_Data.UpdateFDDetails"
        ElseIf Recordstate = DBDeleteFD Then
            .CommandText = "kccbank.PACK_FDDetails_Data.DeleteFDDetails"
        End If
        
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        
        
        If Recordstate = DBInsertFD Or Recordstate = DBModifyFD Or Recordstate = DBDeleteFD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveFD.FDNO)
        End If
        If Recordstate = DBInsertFD Or Recordstate = DBModifyFD Then
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveFD.FDDate)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.FDAmount)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.PeriodCrt)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.Period)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.ROI)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.EitherS)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.FDMode)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveFD.ACType)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.ACNo)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveFD.MatureDate)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.MatureAmt)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.InterestAmt)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.Old_New)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.FD_Lien)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.FD_Tds)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveFD.fd_status)
        End If
                
        If Recordstate = DBDeleteFD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveFD.ISdeleted)
        End If
        If Recordstate = DBInsertFD Or Recordstate = DBModifyFD Or Recordstate = DBDeleteFD Then
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveFD.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveFD.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, , SaveFD.InsertModifyDate)
        End If
        
        If Recordstate = DBInsertFD Or Recordstate = DBModifyFD Then
            .Parameters.Append .CreateParameter("param2", adNumeric, adParamOutput)
        End If
        
        .Execute
        
    End With
    Set objCommand = Nothing
End Sub


Public Function GetFDMasterDetails(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
'Purpose        :- This function will get Master Details corresponding to given FD No
'                  and populate them in required fields

    Dim objCommand As New ADODB.Command
        
    With objCommand
        .CommandText = "kccbank.Pack_FDDetails_Data.GetFDMasterDetails()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveFD.FDNO)
        Set GetFDMasterDetails = .Execute
    End With
    Set objCommand = Nothing
End Function


Public Function GetFDDetails(ByRef objDatasource As ADODB.Connection) _
                               As ADODB.Recordset
'Purpose        :- This function will get Details corresponding to given FD No
'                  and populate them in required fields

    Dim objCommand As New ADODB.Command
        
    With objCommand
        .CommandText = "kccbank.Pack_FDDetails_Data.GetFDDetails()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveFD.FDNO)
        Set GetFDDetails = .Execute
    End With
    Set objCommand = Nothing
End Function
