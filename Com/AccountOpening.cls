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
'Declare type for all variables
Private Type CustomerData
    Customer_key        As Long
    TypeOfAccount       As String
    AccountNo           As Long
    NameOfAcHolder      As String
    FatherName          As String
    Address             As String
    PhNo                As String
    AdultMinor          As String
    Dob                 As Date
    Guardian            As String
    DOBGur              As Date
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
    IntroducedBy        As Long
    TypeIntroduced      As String
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
    Gl_No               As String
    Opening_Date        As Date
    Tds                 As String
    Status_code         As Long
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
    Approved_By1        As Byte
    Approved_By2        As Byte
    Approved_By3        As Byte
    Item                As String
    Amount_Sanctioned   As Double
    Period              As Integer
    Installment         As Double
    
'    joiningDate         As Date
'    RetirementDate      As Date
'    GrossSalary         As Double
'    NetSalary           As Double
'    RegularConfirmed    As String
'    officeAddress       As String
    CDLData()             As String
'
'    NameSur1            As String
'    NameSur2            As String
'    AddSur1             As String
'    AddSur2             As String
'    AcTypeSur1          As String
'    AcTypeSur2          As String
'    AcNoSur1            As Long
'    AcNoSur2            As Long
'    NameWit1            As String
'    NameWit2            As String
'    AddWit1             As String
'    AddWit2             As String
'    AcTypeWit1          As String
'    AcTypeWit2          As String
'    AcNoWit1            As Long
'    AcNoWit2            As Long
    
End Type
Public Enum dbState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum
Private SaveCustomerDetail     As CustomerData
'Set Properties for all variables

Public Property Let Approved_By1(ByVal v_Value As String)
    SaveCustomerDetail.Approved_By1 = v_Value
End Property
Public Property Let Approved_By2(ByVal v_Value As String)
    SaveCustomerDetail.Approved_By2 = v_Value
End Property
Public Property Let Approved_By3(ByVal v_Value As String)
    SaveCustomerDetail.Approved_By3 = v_Value
End Property
Public Property Get Approved_By1() As String
    Approved_By1 = SaveCustomerDetail.Approved_By1
End Property
Public Property Get Approved_By2() As String
    Approved_By2 = SaveCustomerDetail.Approved_By2
End Property
Public Property Get Approved_By3() As String
    Approved_By3 = SaveCustomerDetail.Approved_By3
End Property
Public Property Let Customer_key(ByVal v_Value As Long)
    SaveCustomerDetail.Customer_key = v_Value
End Property
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
Public Property Let DOBGur(ByVal v_Value As Date)
    SaveCustomerDetail.DOBGur = v_Value
End Property
Public Property Let RelationGur(ByVal v_Value As String)
    SaveCustomerDetail.RelationGur = v_Value
End Property
Public Property Let Nominee(ByVal v_Value As String)
    SaveCustomerDetail.Nominee = v_Value
End Property
Public Property Let DOBNom(ByVal v_Value As Date)
    SaveCustomerDetail.DOBNom = Format(v_Value, "dd-mm-yyyy")
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
Public Property Let IntroducedBy(ByVal v_Value As Long)
    SaveCustomerDetail.IntroducedBy = v_Value
End Property
Public Property Let TypeIntroduced(ByVal v_Value As String)
    SaveCustomerDetail.TypeIntroduced = v_Value
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
Public Property Let Gl_No(ByVal v_Value As String)
    SaveCustomerDetail.Gl_No = v_Value
End Property
Public Property Let Opening_Date(ByVal v_Value As Date)
    SaveCustomerDetail.Opening_Date = v_Value
End Property
Public Property Let Tds(ByVal v_Value As String)
    SaveCustomerDetail.Tds = v_Value
End Property
Public Property Let Status_code(ByVal v_Value As String)
    SaveCustomerDetail.Status_code = v_Value
End Property
Public Property Let Item(ByVal v_Value As String)
    SaveCustomerDetail.Item = v_Value
End Property
Public Property Let Amount_Sanctioned(ByVal v_Value As Double)
    SaveCustomerDetail.Amount_Sanctioned = v_Value
End Property
Public Property Let Period(ByVal v_Value As Integer)
    SaveCustomerDetail.Period = v_Value
End Property
Public Property Let Installment(ByVal v_Value As Double)
    SaveCustomerDetail.Installment = v_Value
End Property
Public Property Let CDLData(ByRef v_Value() As String)
    SaveCustomerDetail.CDLData = v_Value
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
Public Property Get Customer_key() As Long
    Customer_key = SaveCustomerDetail.Customer_key
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
Public Property Get DOBGur() As Date
    DOBGur = SaveCustomerDetail.DOBGur
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
Public Property Get IntroducedBy() As Long
    IntroducedBy = SaveCustomerDetail.IntroducedBy
End Property
Public Property Get TypeIntroduced() As String
    TypeIntroduced = SaveCustomerDetail.TypeIntroduced
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
Public Property Get Gl_No() As String
    Gl_No = SaveCustomerDetail.Gl_No
End Property
Public Property Get Opening_Date() As Date
    Opening_Date = SaveCustomerDetail.Opening_Date
End Property
Public Property Get Tds() As String
    Tds = SaveCustomerDetail.Tds
End Property
Public Property Get Status_code() As String
    Status_code = SaveCustomerDetail.Status_code
End Property
Public Property Get Item() As String
    Item = SaveCustomerDetail.Item
End Property
Public Property Get Amount_Sanctioned() As Double
    Amount_Sanctioned = SaveCustomerDetail.Amount_Sanctioned
End Property
Public Property Get Period() As Integer
    Period = SaveCustomerDetail.Period
End Property
Public Property Get Installment() As Double
    Installment = SaveCustomerDetail.Installment
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
Public Function GetApproval(ByRef objDatasource As ADODB.Connection) _
                                As Byte
        'Get Balances
        Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "kccBank.PACK_Account_Opening.GetApproval"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
            .Parameters.Append .CreateParameter("Param1", adChar, adParamOutput, 30)
            .Parameters.Append .CreateParameter("Param2", adChar, adParamOutput, 30)
            .Parameters.Append .CreateParameter("Param3", adChar, adParamOutput, 30)
            .Parameters.Append .CreateParameter("Param4", adNumeric, adParamOutput)
            'Execute to fetch records
            .Execute
            End With
            If Not IsNull(objCommand("Param4").Value) Then
                GetApproval = objCommand("Param4").Value
                If objCommand("param1").Value = "N" Then
                    SaveCustomerDetail.Approved_By1 = 1
                Else
                        SaveCustomerDetail.Approved_By1 = 0
                End If
                If objCommand("param2").Value = "N" Then
                    SaveCustomerDetail.Approved_By2 = 1
                Else
                        SaveCustomerDetail.Approved_By2 = 0
                End If
                If objCommand("param3").Value = "N" Then
                    SaveCustomerDetail.Approved_By3 = 1
                Else
                    SaveCustomerDetail.Approved_By3 = 0
                End If
'            Else
'                GetApproval = -1
            End If
            Set objCommand = Nothing
End Function
Public Function GETStatusCode(ByRef objDatasource As ADODB.Connection) _
                                   As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "PACK_TDS_Master.GETStatusCode()"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = objDatasource
        .CommandTimeout = 0
        Set GETStatusCode = .Execute
    End With
    Set objCommand = Nothing
End Function
Public Sub INSERTModifyDATA(ByRef objDatasource As ADODB.Connection, ByVal Dowork As Byte)
    'Save insert and modify data
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
    If Dowork = 1 Then
        'if DoWork =1 then call InsertData Procedure
        .CommandText = "kccBank.PACK_Account_Opening.INSERTDATA"
    Else
        'Else call ModifyData Procedure
        .CommandText = "kccBank.PACK_Account_Opening.ModifyData"
    End If
    .CommandType = adCmdStoredProc
    'Send connection
    .ActiveConnection = objDatasource
    'Set command time out
    .CommandTimeout = 0
    'Set All I/O Parameters
    If Dowork = 2 Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
    End If
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.NameOfAcHolder)
    If SaveCustomerDetail.FatherName = Empty Then
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
     Else
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.FatherName)
     End If
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.Address)
    If SaveCustomerDetail.PhNo = Empty Then
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, Null)
    Else
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveCustomerDetail.PhNo)
    End If
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.AdultMinor)
    If SaveCustomerDetail.Dob = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
    Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.Dob)
    End If
    If SaveCustomerDetail.Guardian = Empty Then
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
    Else
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.Guardian)
    End If
    If SaveCustomerDetail.DOBGur = Empty Then
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
    Else
        .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.DOBGur)
    End If
    If SaveCustomerDetail.RelationGur = Empty Then
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
    Else
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveCustomerDetail.RelationGur)
    End If
    If SaveCustomerDetail.Nominee = Empty Then
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
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
    If SaveCustomerDetail.IntroducedBy = Empty Then
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, Null)
    Else
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.IntroducedBy)
    End If
    If SaveCustomerDetail.TypeIntroduced = Empty Then
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, Null)
    Else
        .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeIntroduced)
    End If
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
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveCustomerDetail.Gl_No)
    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.Opening_Date)
    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.Tds)
    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Status_code)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
    If Dowork = 1 Then
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
    End If
    .Execute
    If Dowork = 1 Then
        If Not IsNull(objCommand("Param1").Value) Then
            SaveCustomerDetail.Customer_key = CLng(objCommand("Param1").Value)
        End If
    End If
End With
Set objCommand = Nothing
              End Sub
Public Sub INSERTModifyCDLoanData(ByRef objDatasource As ADODB.Connection, ByVal Dowork As Byte)
    Dim objCommand As ADODB.Command
        Set objCommand = New ADODB.Command
        With objCommand
        If Dowork = 1 Then
                .CommandText = "kccBank.PACK_Account_Opening.INSERTDATACDL"
        ElseIf Dowork = 2 Then
                .CommandText = "kccBank.PACK_Account_Opening.ModifyDataCDL"
        End If
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.Item)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, SaveCustomerDetail.Amount_Sanctioned)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 3, SaveCustomerDetail.Period)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, SaveCustomerDetail.Installment)
            If SaveCustomerDetail.CDLData(0) = Empty Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
            Else
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(SaveCustomerDetail.CDLData(0)))
            End If
            If SaveCustomerDetail.CDLData(1) = Empty Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
            Else
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(SaveCustomerDetail.CDLData(1)))
            End If
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.CDLData(2))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.CDLData(3))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, CDbl(SaveCustomerDetail.CDLData(4)))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 10, CDbl(SaveCustomerDetail.CDLData(5)))
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.CDLData(6))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.CDLData(7))
            'If SaveCustomerDetail.CDLData(8) = Empty Then
            '    .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, Null)
            'Else
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.CDLData(8))
            'End If
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.CDLData(9)))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.CDLData(10))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.CDLData(11))
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.CDLData(12))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.CDLData(13)))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.CDLData(14))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.CDLData(15))
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.CDLData(16))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.CDLData(17)))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.CDLData(18))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveCustomerDetail.CDLData(19))
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.CDLData(20))
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.CDLData(21)))
            
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
            .Execute
       End With
       Set objCommand = Nothing
End Sub
Public Sub INSERTModifyJointData(ByRef objDatasource As ADODB.Connection, ByVal Dowork As Byte)
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
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
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
        Set .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
         .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
                'If duplicate record return true
                CheckAccount = True
            Else
                'else return false
                CheckAccount = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function DeleteData(ByRef objDatasource As ADODB.Connection, ByVal Dowork As Byte) _
                                As Boolean
    Dim objCommand As New ADODB.Command
    With objCommand
        If Dowork = 1 Then
            .CommandText = "KccBank.PACK_Account_Opening.DeleteDataPerm"
        Else
            .CommandText = "KccBank.PACK_Account_Opening.DeleteData"
        End If
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        If Dowork = 1 Then
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        Else
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveCustomerDetail.AccountNo))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        End If
        'Execute to fetch records
         .Execute
         If Not IsNull(objCommand("Param1").Value) Then
            If objCommand("Param1").Value = 0 Then
                DeleteData = True
            Else
                DeleteData = False
            End If
         End If
    End With
End Function
Public Function PopupLov(ByRef objDatasource As ADODB.Connection) As Boolean
    'Get the fields from t_emp_mst table Using LOV
    'Make the object of LOV
    Dim look As New Lov.LookUp
    Dim strProcedure As String
    Dim sql1, sql As String
    sql1 = "Select CUSTOMER_KEY,TYPE_OF_ACCOUNT, ACCOUNT_NO, NAME_OF_AC_HOLDER, ADDRESS," _
            & "PHONE_NO from kccbank.T_Customer_detail where " _
            & "Delete_Flag=" & "'N'" & " And type_of_account = " & "'" & SaveCustomerDetail.TypeOfAccount & "'" _
            & " order by TYPE_OF_ACCOUNT, ACCOUNT_NO"
    sql = "Select CUSTOMER_KEY,TYPE_OF_ACCOUNT, ACCOUNT_NO, NAME_OF_AC_HOLDER, ADDRESS, PHONE_NO from kccbank.T_Customer_detail where Delete_Flag=" & "'N'"
    With look
        'Provide column header for the LOV
        .AddColumnHeaders "Customer Key", "Type of Account", "Account No.", "Name Of A/C Holer", "Address", "Ph. No."
        'Provide actual column for the ListView According to the Stored Procedure"
        .AddDisplayFields "CUSTOMER_KEY", "TYPE_OF_ACCOUNT", "ACCOUNT_NO", "NAME_OF_AC_HOLDER", "ADDRESS", "PHONE_NO"
        .Connection = objDatasource
        .SetColumnsWidth 0, 1200, 700, 2000, 2500, 1000
        .ProcedureText = sql1
        .TotalColumns = 6
        .PopulateList
        If .LOVState = False Then
            SaveCustomerDetail.Customer_key = .DisplayValueByName("CUSTOMER_KEY")
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
Public Function GetCDLoanData(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.GetCDLoanData"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
        'Execute to fetch records
        Set GetCDLoanData = .Execute
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
Public Function GetAccountType(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "{call KccBank.PACK_ACCOUNT_OPENING.GetAccountType()}"
            .CommandType = adCmdText
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            'Execute to fetch records
            Set GetAccountType = .Execute
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
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
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
    'Delete Single Joint Record
    Dim objCommand As New ADODB.Command
        With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteJointSinleRecord"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.Customer_key)
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
    'Delete all Joint Records
    Dim objCommand As New ADODB.Command
    With objCommand
        .CommandText = "KccBank.PACK_Account_Opening.DeleteJointAllRecords"
        .CommandType = adCmdStoredProc
        'Send connection
        Set .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , CInt(40)) 'CLng(SaveCustomerDetail.Customer_key))
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch record
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
