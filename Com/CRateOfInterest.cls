VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRateOfInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
'       PROJECT    :KCCRATEOFINTEREST
'       MODULE     :ESTABLISHMENT
'       PRUPOSE    :A genric solution for the stockhand drawing power
'
'       COM        :C RATE OF INTEREST
'       FORM
'       REFERRED   :RATE OF INTEREST
'       FORM CODE  :KCCB EST 0004
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :SANDEEP BASSI
'       MADE DATE  :19-03-2002
'       REVISIONS  :
'       COPY RIGHT @ 0-2002 SURYA INFONET LTD.
'------------------------------------------------------------------

Private Type ROIData
    ROICode         As Integer
    ROIDate         As Date
    AccCode         As String
    ROI_Public_Cr   As Double
    ROI_Public_Dr   As Double
    ROI_Staff_Cr    As Double
    ROI_Staff_Dr    As Double
    TerminalName    As String
    UserName        As String
    InsertDate      As Date
    ModifyDate      As Date
End Type

Public Enum DBState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum

Private DBData  As ROIData
Private m_blnCheck As Boolean

Public Property Let ROICode(ByVal v_Value As Integer)
    DBData.ROICode = v_Value
End Property

Public Property Let ROIDate(ByVal v_Value As Date)
    DBData.ROIDate = v_Value
End Property

Public Property Let AccCode(ByVal v_Value As String)
    DBData.AccCode = v_Value
End Property

Public Property Let ROI_Public_Cr(ByVal v_Value As Double)
    DBData.ROI_Public_Cr = v_Value
End Property
Public Property Let ROI_Public_Dr(ByVal v_Value As Double)
    DBData.ROI_Public_Dr = v_Value
End Property
Public Property Let ROI_Staff_Cr(ByVal v_Value As Double)
    DBData.ROI_Staff_Cr = v_Value
End Property

Public Property Let ROI_Staff_Dr(ByVal v_Value As Double)
    DBData.ROI_Staff_Dr = v_Value
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

Public Property Get IsUniqueCheck() As Boolean
    IsUniqueCheck = m_blnCheck
End Property

'To retrieve list of Existing RAte of interests from database
Public Function GetROIList(ByRef objDatasource As ADODB.Connection) As ADODB.Recordset
    Dim objCmd As ADODB.Command
    Set objCmd = New ADODB.Command
        
        With objCmd
            .CommandText = "Pack_Rate_Of_ineterst.GetROIListExists()"
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
                .CommandTimeout = 0
            Set GetROIList = .Execute
        End With
End Function

'To check weather Rate of interest for particular code and date exists or not
Public Function CheckROI(ByRef objDatasource As ADODB.Connection)
    Dim objCmd As ADODB.Command
    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Set objCmd = New ADODB.Command
    
        With objCmd
            .ActiveConnection = objDatasource
            Set Param1 = .CreateParameter("Param1", adNumeric, adParamInput, , DBData.ROICode)
                .Parameters.Append Param1
            Set Param2 = .CreateParameter("Param2", adDate, adParamInput, 8, CDate(Format(DBData.ROIDate, "dd-mm-yyyy")))
                .Parameters.Append Param2
            Set Param3 = .CreateParameter("Param3", adNumeric, adParamOutput)
                .Parameters.Append Param3
                
            .CommandText = "{call Pack_Rate_Of_ineterst.CheckROI(?,?,?)}"
            
            .Execute
            
                If Param3.Value = 0 Then
                    m_blnCheck = True
                Else
                    m_blnCheck = False
                End If
        
        End With
        
    Set Param3 = Nothing
    Set objCmd = Nothing

End Function

'To Insert/Modify/Delete data in database
Public Sub SaveData(ByRef objDatasource As ADODB.Connection, ByVal Recordstate As DBState)
    Dim objCmd As ADODB.Command
    Set objCmd = New ADODB.Command
    
        With objCmd
            If Recordstate = DBInsert Then
                .CommandText = "Pack_Rate_Of_ineterst.InsertData"
            ElseIf Recordstate = DBModify Then
                .CommandText = "Pack_Rate_Of_ineterst.ModifyData"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "Pack_Rate_Of_ineterst.DeleteData"
            End If
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = objDatasource
                .CommandTimeout = 0
                
                If Recordstate = DBInsert Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROICode)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.ROIDate, "dd-mm-yyyy")))
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, DBData.AccCode)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Public_Cr)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Public_Dr)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Staff_Cr)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Staff_Dr)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                ElseIf Recordstate = DBModify Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROICode)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.ROIDate, "dd-mm-yyyy")))
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, DBData.AccCode)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Public_Cr)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Public_Dr)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Staff_Cr)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.ROI_Staff_Dr)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.ModifyDate, "dd-mm-yyyy")))
                ElseIf Recordstate = DBDelete Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, DBData.ROICode)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.ROIDate, "dd-mm-yyyy")))
                End If
                .Execute
        End With
        If Recordstate = DBInsert Then
            MsgBox "Record Inserted Successfully", vbInformation, "Record Insertion"
        End If
          
        If Recordstate = DBModify Then
            MsgBox "Record Modified Successfully", vbInformation, "Record Modification"
        End If
        
        If Recordstate = DBDelete Then
            MsgBox "Record Deleted Successfully", vbInformation, "Record Deletion"
        End If
    Set objCmd = Nothing
End Sub
