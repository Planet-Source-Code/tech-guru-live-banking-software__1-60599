VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccountOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type CustomerData
    TypeOfAccount       As String
    AccountNo           As Long
    NameOfAcHolder      As String
    FatherName          As String
    Address             As String
    PhNo                As String
    AdultMinor          As String
    Dob                 As Date
    Guardian            As String
    DobGur              As Date
    RelationGur         As String
    Nominee             As String
    DOBNom              As Date
    AddressNom          As String
    RelationNom         As String
    SingleJoint         As String
    RateOfInterestC     As Double
    RateOfInterestD     As Double
    Occupation          As Long
    StaffPublic         As String
    IntroducedBy        As String
    CanIntroduce        As String
    ReasonIntro         As String
    Pan                 As String
    ClosedOperative     As String
    ReasonClosed        As String
    Remarks             As String
    LedgerNo            As Long
    Severally_Joint     As String
    jointName           As String
    JointKey            As Single
    MustOperate         As String
    CanOperate          As String
    JointData()         As String
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
End Type
Public Enum dbState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum
Private SaveCustomerDetail     As CustomerData
Public Property Let TypeOfAccount(ByVal v_Value As String)
    SaveCustomerDetail.TypeOfAccount = v_Value
End Property
Public Property Let AccountNo(ByVal v_Value As Long)
    SaveCustomerDetail.AccountNo = v_Value
End Property
Public Property Let NameOfAcHolder(ByVal v_Value As String)
    SaveCustomerDetail.NameOfAcHolder = v_Value
End Property
Public Property Let FatherName(ByVal v_Value As String)
    SaveCustomerDetail.FatherName = v_Value
End Property
Public Property Let Address(ByVal v_Value As String)
    SaveCustomerDetail.Address = v_Value
End Property
Public Property Let PhNo(ByVal v_Value As String)
    SaveCustomerDetail.PhNo = v_Value
End Property
Public Property Let AdultMinor(ByVal v_Value As String)
    SaveCustomerDetail.AdultMinor = v_Value
End Property
Public Property Let Dob(ByVal v_Value As Date)
    SaveCustomerDetail.Dob = v_Value
End Property
Public Property Let Guardian(ByVal v_Value As String)
    SaveCustomerDetail.Guardian = v_Value
End Property
Public Property Let DobGur(ByVal v_Value As Date)
    SaveCustomerDetail.DobGur = v_Value
End Property

Public Property Let RelationGur(ByVal v_Value As String)
    SaveCustomerDetail.RelationGur = v_Value
End Property
Public Property Let Nominee(ByVal v_Value As String)
    SaveCustomerDetail.Nominee = v_Value
End Property
Public Property Let DOBNom(ByVal v_Value As Date)
    SaveCustomerDetail.DOBNom = v_Value
End Property
Public Property Let AddressNom(ByVal v_Value As String)
    SaveCustomerDetail.AddressNom = v_Value
End Property
Public Property Let RelationNom(ByVal v_Value As String)
    SaveCustomerDetail.RelationNom = v_Value
End Property
Public Property Let SingleJoint(ByVal v_Value As String)
    SaveCustomerDetail.SingleJoint = v_Value
End Property
Public Property Let RateOfInterestC(ByVal v_Value As Double)
    SaveCustomerDetail.RateOfInterestC = v_Value
End Property
Public Property Let RateOfInterestD(ByVal v_Value As Double)
    SaveCustomerDetail.RateOfInterestD = v_Value
End Property
Public Property Let Occupation(ByVal v_Value As Long)
    SaveCustomerDetail.Occupation = v_Value
End Property
Public Property Let StaffPublic(ByVal v_Value As String)
    SaveCustomerDetail.StaffPublic = v_Value
End Property
Public Property Let IntroducedBy(ByVal v_Value As String)
    SaveCustomerDetail.IntroducedBy = v_Value
End Property
Public Property Let CanIntroduce(ByVal v_Value As String)
    SaveCustomerDetail.CanIntroduce = v_Value
End Property
Public Property Let ReasonIntro(ByVal v_Value As String)
    SaveCustomerDetail.ReasonIntro = v_Value
End Property
Public Property Let Pan(ByVal v_Value As String)
    SaveCustomerDetail.Pan = v_Value
End Property
Public Property Let ClosedOperative(ByVal v_Value As String)
    SaveCustomerDetail.ClosedOperative = v_Value
End Property
Public Property Let ReasonClosed(ByVal v_Value As String)
    SaveCustomerDetail.ReasonClosed = v_Value
End Property
Public Property Let Remarks(ByVal v_Value As String)
    SaveCustomerDetail.Remarks = v_Value
End Property
Public Property Let LedgerNo(ByVal v_Value As Integer)
    SaveCustomerDetail.LedgerNo = v_Value
End Property
Public Property Let Severally_Joint(ByVal v_Value As String)
    SaveCustomerDetail.Severally_Joint = v_Value
End Property
Public Property Let JointData(ByRef v_Value() As String)
    SaveCustomerDetail.JointData = v_Value
End Property

Public Property Let jointName(ByVal v_Value As String)
    SaveCustomerDetail.jointName = v_Value
End Property
Public Property Let JointKey(ByVal v_Value As Single)
    SaveCustomerDetail.JointKey = v_Value
End Property
Public Property Let MustOperate(ByVal v_Value As String)
    SaveCustomerDetail.MustOperate = v_Value
End Property
Public Property Let CanOperate(ByVal v_Value As String)
    SaveCustomerDetail.CanOperate = v_Value
End Property
Public Property Let TerminalName(ByVal v_Value As String)
    SaveCustomerDetail.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveCustomerDetail.UserName = v_Value
End Property
Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveCustomerDetail.InsertModifyDate = v_Value
End Property
Public Property Get TypeOfAccount() As String
    TypeOfAccount = SaveCustomerDetail.TypeOfAccount
End Property
Public Property Get AccountNo() As Long
    AccountNo = SaveCustomerDetail.AccountNo
End Property
Public Property Get NameOfAcHolder() As String
    NameOfAcHolder = SaveCustomerDetail.NameOfAcHolder
End Property
Public Property Get FatherName() As String
    FatherName = SaveCustomerDetail.FatherName
End Property
Public Property Get Address() As String
    Address = SaveCustomerDetail.Address
End Property
Public Property Get PhNo() As String
PhNo = SaveCustomerDetail.PhNo
End Property
Public Property Get AdultMinor() As String
    AdultMinor = SaveCustomerDetail.AdultMinor
End Property
Public Property Get Dob() As Date
    Dob = SaveCustomerDetail.Dob
End Property
Public Property Get Guardian() As String
    Guardian = SaveCustomerDetail.Guardian
End Property
Public Property Get DobGur() As Date
    DobGur = SaveCustomerDetail.DobGur
End Property

Public Property Get RelationGur() As String
    RelationGur = SaveCustomerDetail.RelationGur
End Property
Public Property Get Nominee() As String
    Nominee = SaveCustomerDetail.Nominee
End Property
Public Property Get DOBNom() As Date
    DOBNom = SaveCustomerDetail.DOBNom
End Property
Public Property Get AddressNom() As String
    AddressNom = SaveCustomerDetail.AddressNom
End Property
Public Property Get RelationNom() As String
    RelationNom = SaveCustomerDetail.RelationNom
End Property

Public Property Get SingleJoint() As String
    SingleJoint = SaveCustomerDetail.SingleJoint
End Property
Public Property Get RateOfInterestC() As Double
    RateOfInterestC = SaveCustomerDetail.RateOfInterestC
End Property
Public Property Get RateOfInterestD() As Double
    RateOfInterestD = SaveCustomerDetail.RateOfInterestD
End Property
Public Property Get Occupation() As Long
    Occupation = SaveCustomerDetail.Occupation
End Property
Public Property Get StaffPublic() As String
    StaffPublic = SaveCustomerDetail.StaffPublic
End Property
Public Property Get IntroducedBy() As String
    IntroducedBy = SaveCustomerDetail.IntroducedBy
End Property
Public Property Get CanIntroduce() As String
    CanIntroduce = SaveCustomerDetail.CanIntroduce
End Property
Public Property Get ReasonIntro() As String
    ReasonIntro = SaveCustomerDetail.ReasonIntro
End Property
Public Property Get Pan() As String
    Pan = SaveCustomerDetail.Pan
End Property
Public Property Get ClosedOperative() As String
    ClosedOperative = SaveCustomerDetail.ClosedOperative
End Property
Public Property Get ReasonClosed() As String
    ReasonClosed = SaveCustomerDetail.ReasonClosed
End Property
Public Property Get Remarks() As String
    Remarks = SaveCustomerDetail.Remarks
End Property
Public Property Get LedgerNo() As Integer
    LedgerNo = SaveCustomerDetail.LedgerNo
End Property
Public Property Get Severally_Joint() As String
    Severally_Joint = SaveCustomerDetail.Severally_Joint
End Property
Public Property Get JointData() As String()
    JointData = SaveCustomerDetail.JointData
End Property

Public Property Get jointName() As String
    jointName = SaveCustomerDetail.jointName
End Property
Public Property Get JointKey() As Single
    JointKey = SaveCustomerDetail.JointKey
End Property
Public Property Get MustOperate() As String
    MustOperate = SaveCustomerDetail.MustOperate
End Property
Public Property Get CanOperate() As String
    CanOperate = SaveCustomerDetail.CanOperate
End Property
Public Property Get TerminalName() As String
    TerminalName = SaveCustomerDetail.TerminalName
End Property
Public Property Get UserName() As String
    UserName = SaveCustomerDetail.UserName
End Property
Public Property Get InsertModifyDate() As Date
 InsertModifyDate = SaveCustomerDetail.InsertModifyDate
End Property
Public Sub INSERTModifyDATA(ByRef objDatasource As ADODB.Connection, ByVal doWork As Byte)
    Dim objCommand As ADODB.Command
                Set objCommand = New ADODB.Command
                With objCommand
                If doWork = 1 Then
                    .CommandText = "kccBank.PACK_Account_Opening.INSERTDATA"
                Else
                    .CommandText = "kccBank.PACK_Account_Opening.ModifyData"
                End If
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.NameOfAcHolder)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.FatherName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.Address)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveCustomerDetail.PhNo)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.AdultMinor)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.Dob)
                    If SaveCustomerDetail.Guardian = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.Guardian)
                    End If
                    If SaveCustomerDetail.DobGur = Empty Then
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.DobGur)
                    End If
                    If SaveCustomerDetail.RelationGur = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveCustomerDetail.RelationGur)
                    End If
                    If SaveCustomerDetail.Nominee = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.Nominee)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.Nominee)
                    End If
                    If SaveCustomerDetail.DOBNom = Empty Then
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.DOBNom)
                    End If
                    If SaveCustomerDetail.AddressNom = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.AddressNom)
                    End If
                    If SaveCustomerDetail.RelationNom = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveCustomerDetail.RelationNom)
                    End If
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.SingleJoint)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveCustomerDetail.RateOfInterestC)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveCustomerDetail.RateOfInterestD)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveCustomerDetail.Occupation)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.StaffPublic)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 9, SaveCustomerDetail.IntroducedBy)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.CanIntroduce)
                    If SaveCustomerDetail.ReasonIntro = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveCustomerDetail.ReasonIntro)
                    End If
                    If SaveCustomerDetail.Pan = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveCustomerDetail.Pan)
                    End If
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.ClosedOperative)
                    If SaveCustomerDetail.ReasonClosed = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveCustomerDetail.ReasonClosed)
                    End If
                    If SaveCustomerDetail.Remarks = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.Remarks)
                    End If
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveCustomerDetail.LedgerNo)
                    If SaveCustomerDetail.Severally_Joint = Empty Then
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.Severally_Joint)
                    End If
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
                    .Execute
                End With
                Set objCommand = Nothing
              End Sub
Public Sub INSERTModifyJointData(ByRef objDatasource As ADODB.Connection, ByVal doWork As Byte)
    Dim objCommand As ADODB.Command
    Dim i As Integer
            For i = 0 To UBound(SaveCustomerDetail.JointData) - 1 Step 4
            Set objCommand = New ADODB.Command
            With objCommand
                If SaveCustomerDetail.JointData(i + 3) = "i" Then
                    .CommandText = "kccBank.PACK_Account_Opening.INSERTJointDATA"
                Else
                    .CommandText = "kccBank.PACK_Account_Opening.ModifyJointData"
                End If
                .CommandType = adCmdStoredProc
                'Send connection
                .ActiveConnection = objDatasource
                'Set command time out
                .CommandTimeout = 0
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, (SaveCustomerDetail.JointData(i + 1)))
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, CInt(SaveCustomerDetail.JointData(i)))
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.JointData(i + 2))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
                .Execute
            End With
            Set objCommand = Nothing
        Next
End Sub
Public Function CheckAccount(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.CheckAccount"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
         .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
                CheckAccount = True
            Else
                CheckAccount = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function DeleteDATA(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteData"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        DeleteDATA = True
                    Else
                        DeleteDATA = False
                    End If
                End If
        End With
End Function
Public Function PopupLov(ByRef objDatasource As ADODB.Connection) As Boolean
    'Get the fields from t_emp_mst table Using LOV
    'Make the object of LOV
    Dim look As New LookUp
    Dim strProcedure As String
    Dim sql As String
    sql = "Select TYPE_OF_ACCOUNT, ACCOUNT_NO, NAME_OF_AC_HOLDER, ADDRESS, PHONE_NO from kccbank.T_Customer_detail where Delete_Flag=" & "'N'"
   
    'strProcedure = "{KccBank.PACK_Account_Opening.GETACCOUNTNOS()}"
    With look
        'Provide column header for the LOV
        .AddColumnHeaders "Type of Account", "Account No.", "Name Of A/C Holer", "Address", "Ph. No."
        'Provide actual column for the ListView According to the Stored Procedure"
        .AddDisplayFields "TYPE_OF_ACCOUNT", "ACCOUNT_NO", "NAME_OF_AC_HOLDER", "ADDRESS", "PHONE_NO"
        .Connection = objDatasource
        .SetColumnsWidth 1200, 700, 2000, 2500, 1000
        .ProcedureText = sql
        .TotalColumns = 5
        .PopulateList
        If .LOVState = False Then
            SaveCustomerDetail.TypeOfAccount = (.DisplayValueByName("TYPE_OF_ACCOUNT"))
            SaveCustomerDetail.AccountNo = CLng(.DisplayValueByName("ACCOUNT_NO"))
            PopupLov = True
            Exit Function
         
         End If
     PopupLov = False
    End With
End Function
Public Function GetACCOUNTDETAIL(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
                                
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Account_Opening.GetACCOUNTDETAIL"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            'Execute to fetch records
            Set GetACCOUNTDETAIL = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GETJOINTDATA(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Account_Opening.GETJOINTDATA"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            'Execute to fetch records
            Set GETJOINTDATA = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GetOccupation(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "{call KccBank.PACK_ACCOUNT_OPENING.GetOccupation()}"
            .CommandType = adCmdText
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            'Execute to fetch records
            Set GetOccupation = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GETAccNo(ByRef objDatasource As ADODB.Connection) _
                                As Long
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = " KccBank.PACK_ACCOUNT_OPENING.GETAccNo"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    GETAccNo = objCommand("Param1").Value
                End If
         End With
    Set objCommand = Nothing
End Function

Public Function DeleteJointSinleRecord(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteJointSinleRecord"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, CInt(SaveCustomerDetail.JointKey))
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        DeleteJointSinleRecord = True
                    Else
                        DeleteJointSinleRecord = False
                    End If
                End If
        End With
End Function
Public Function DeleteJointAllRecords(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteJointAllRecords"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    If objCommand("Param1").Value = 0 Then
                        DeleteJointAllRecords = True
                    Else
                        DeleteJointAllRecords = False
                    End If
                End If
        End With
End Function

