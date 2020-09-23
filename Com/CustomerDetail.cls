VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type CustomerData
    TypeOfAccount       As String
    AccountNo           As Long
    NameOfAcHolder      As String
    Address             As String
    PhNo                As String
    AdultMinor          As String
    Dob                 As Data
    Guardian            As String
    DOBGur              As Data
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
    jointName           As String
    JointKey            As Single
    MustOperate         As String
    CanOperate          As String
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
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveCustomerDetail.AccountNo)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.NameOfAcHolder)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveCustomerDetail.Address)
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.PhNo)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveCustomerDetail.AdultMinor)
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, SaveCustomerDetail.Dob)
                    If SaveCustomerDetail.Guardian = Empty Then
                        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveCustomerDetail.Guardian)
                    End If
                    If SaveCustomerDetail.DOBGur = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.DOBGur)
                    End If
                    If SaveCustomerDetail.RelationGur = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.RelationGur)
                    End If
                    If SaveCustomerDetail.Nominee = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.Nominee)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.Nominee)
                    End If
                    If SaveCustomerDetail.AddressNom = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.AddressNom)
                    End If
                    If SaveCustomerDetail.RelationNom = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.RelationNom)
                    End If
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.SingleJoint)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.RateOfInterestC)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.RateOfInterestD)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.Occupation)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.StaffPublic)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.IntroducedBy)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.CanIntroduce)
                    If SaveCustomerDetail.ReasonIntro = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.ReasonIntro)
                    End If
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.Pan)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.ClosedOperative)
                    If SaveCustomerDetail.ReasonClosed = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.ReasonClosed)
                    End If
                    If SaveCustomerDetail.Remarks = Empty Then
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, Null)
                    Else
                        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.Remarks)
                    End If
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.LedgerNo)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
                    .Execute
                End With
                Set objCommand = Nothing
              End Sub
Public Sub INSERTModifyJointData(ByRef objDatasource As ADODB.Connection, ByVal doWork As Byte)
    Dim objCommand As ADODB.Command
            Set objCommand = New ADODB.Command
            With objCommand
                If doWork = 1 Then
                    .CommandText = "kccBank.PACK_Account_Opening.INSERTJointDATA"
                Else
                    .CommandText = "kccBank.PACK_Account_Opening.ModifyJointData"
                End If
                .CommandType = adCmdStoredProc
                'Send connection
                .ActiveConnection = objDatasource
                'Set command time out
                .CommandTimeout = 0
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveCustomerDetail.AccountNo)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveCustomerDetail.jointName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.JointKey)
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveCustomerDetail.MustOperate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 8, SaveCustomerDetail.CanOperate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveCustomerDetail.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveCustomerDetail.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveCustomerDetail.InsertModifyDate)
                .Execute
            End With
            Set objCommand = Nothing
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
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
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
    strProcedure = "{KccBank.PACK_Account_Opening.GETACCOUNTNOS()}"
    With look
        'Provide column header for the LOV
        .AddColumnHeaders "Type of Account", "Account No.", "Name Of A/C Holer", "Address", "Ph. No."
        'Provide actual column for the ListView According to the Stored Procedure"
        .AddDisplayFields "TYPE_OF_ACCOUNT", "ACCOUNT_NO", "NAME_OF_AC_HOLDER", "ADDRESS", "PHONE_NO"
        .Connection = objDatasource
        .SetColumnsWidth 700, 1200, 2000, 2500, 1000
        .ProcedureText = strProcedure
        .TotalColumns = 5
        .PopulateList
        If .LOVState = False Then
            SaveCustomerDetail.TypeOfAccount = CLng(.DisplayValueByName("TYPE_OF_ACCOUNT"))
            SaveCustomerDetail.AccountNo = .DisplayValueByName("ACCOUNT_NO")
            PopupLov = True
            Exit Function
        End If
     PopupLov = False
    End With
End Function
Public Function GetALLACCOUNTS(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Account_Opening.GetALLACCOUNTS"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
            'Execute to fetch records
            Set GetALLACCOUNTS = .Execute
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
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveCustomerDetail.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveCustomerDetail.AccountNo)
            'Execute to fetch records
            Set GETJOINTDATA = .Execute
         End With
    Set objCommand = Nothing
End Function
