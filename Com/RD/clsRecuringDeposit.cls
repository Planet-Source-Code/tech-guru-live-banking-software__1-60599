VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecuringDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Declare Type for All Variables
Private Type RecurinDeposit
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
    Joint1              As String
    Joint2              As String
    Joint3              As String
    Severally_Joint     As String
    Period              As Integer
    Amount              As Double
    RateOfInterest      As Double
    DateAccOpening      As Date
    MonthPeriod         As Single
    InterestAmount      As Double
    TerminalName        As String
    UserName            As String
    InsertModifyDate    As Date
    Maturity_Amount     As Double
    TDS_Deduct          As String
    TDSStatus           As Long
    RDLien              As String
End Type
Public Enum dbState
    DBInsert = 1
    DBModify = 2
    DBDelete = 3
End Enum
Private SaveRD     As RecurinDeposit
'Declare All Properties
Public Property Let TDSDeduct(ByVal v_Value As String)
    SaveRD.TDS_Deduct = v_Value
End Property
Public Property Let TDSStatus(ByVal v_Value As Long)
    SaveRD.TDSStatus = v_Value
End Property
Public Property Let RDLien(ByVal v_Value As String)
    SaveRD.RDLien = v_Value
End Property

Public Property Let TypeOfAccount(ByVal v_Value As String)
    SaveRD.TypeOfAccount = v_Value
End Property
Public Property Let AccountNo(ByVal v_Value As Long)
    SaveRD.AccountNo = v_Value
End Property
Public Property Let NameOfAcHolder(ByVal v_Value As String)
    SaveRD.NameOfAcHolder = v_Value
End Property
Public Property Let FatherName(ByVal v_Value As String)
    SaveRD.FatherName = v_Value
End Property
Public Property Let Address(ByVal v_Value As String)
    SaveRD.Address = v_Value
End Property
Public Property Let PhNo(ByVal v_Value As String)
    SaveRD.PhNo = v_Value
End Property
Public Property Let AdultMinor(ByVal v_Value As String)
    SaveRD.AdultMinor = v_Value
End Property
Public Property Let Dob(ByVal v_Value As Date)
    SaveRD.Dob = v_Value
End Property
Public Property Let Guardian(ByVal v_Value As String)
    SaveRD.Guardian = v_Value
End Property
Public Property Let DOBGur(ByVal v_Value As Date)
    SaveRD.DOBGur = v_Value
End Property
Public Property Let RelationGur(ByVal v_Value As String)
    SaveRD.RelationGur = v_Value
End Property
Public Property Let Nominee(ByVal v_Value As String)
    SaveRD.Nominee = v_Value
End Property
Public Property Let DOBNom(ByVal v_Value As Date)
    SaveRD.DOBNom = Format(v_Value, "dd-mm-yyyy")
End Property
Public Property Let AddressNom(ByVal v_Value As String)
    SaveRD.AddressNom = v_Value
End Property
Public Property Let RelationNom(ByVal v_Value As String)
    SaveRD.RelationNom = v_Value
End Property
Public Property Let SingleJoint(ByVal v_Value As String)
    SaveRD.SingleJoint = v_Value
End Property
Public Property Let Joint1(ByVal v_Value As String)
    SaveRD.Joint1 = v_Value
End Property
Public Property Let Joint2(ByVal v_Value As String)
    SaveRD.Joint2 = v_Value
End Property
Public Property Let Joint3(ByVal v_Value As String)
    SaveRD.Joint3 = v_Value
End Property
Public Property Let Period(ByVal v_Value As Single)
    SaveRD.Period = v_Value
End Property
Public Property Let Amount(ByVal v_Value As Double)
    SaveRD.Amount = v_Value
End Property
Public Property Let Maturity_Amount(ByVal v_Value As Double)
    SaveRD.Maturity_Amount = v_Value
End Property

Public Property Let RateOfInterest(ByVal v_Value As Double)
    SaveRD.RateOfInterest = v_Value
End Property
Public Property Let DateAccOpening(ByVal v_Value As Date)
    SaveRD.DateAccOpening = v_Value
End Property
Public Property Let MonthPeriod(ByVal v_Value As Single)
    SaveRD.MonthPeriod = v_Value
End Property
Public Property Let InterestAmount(ByVal v_Value As Double)
    SaveRD.InterestAmount = v_Value
End Property
Public Property Let Severally_Joint(ByVal v_Value As String)
    SaveRD.Severally_Joint = v_Value
End Property

Public Property Let TerminalName(ByVal v_Value As String)
    SaveRD.TerminalName = v_Value
End Property
Public Property Let UserName(ByVal v_Value As String)
    SaveRD.UserName = v_Value
End Property
Public Property Let InsertModifyDate(ByVal v_Value As Date)
    SaveRD.InsertModifyDate = v_Value
End Property
Public Property Get TypeOfAccount() As String
    TypeOfAccount = SaveRD.TypeOfAccount
End Property
Public Property Get AccountNo() As Long
    AccountNo = SaveRD.AccountNo
End Property
Public Property Get NameOfAcHolder() As String
    NameOfAcHolder = SaveRD.NameOfAcHolder
End Property
Public Property Get FatherName() As String
    FatherName = SaveRD.FatherName
End Property
Public Property Get Address() As String
    Address = SaveRD.Address
End Property
Public Property Get PhNo() As String
PhNo = SaveRD.PhNo
End Property
Public Property Get AdultMinor() As String
    AdultMinor = SaveRD.AdultMinor
End Property
Public Property Get Dob() As Date
    Dob = SaveRD.Dob
End Property
Public Property Get Guardian() As String
    Guardian = SaveRD.Guardian
End Property
Public Property Get DOBGur() As Date
    DOBGur = SaveRD.DOBGur
End Property
Public Property Get RelationGur() As String
    RelationGur = SaveRD.RelationGur
End Property
Public Property Get Nominee() As String
    Nominee = SaveRD.Nominee
End Property
Public Property Get DOBNom() As Date
    DOBNom = SaveRD.DOBNom
End Property
Public Property Get AddressNom() As String
    AddressNom = SaveRD.AddressNom
End Property
Public Property Get RelationNom() As String
    RelationNom = SaveRD.RelationNom
End Property
Public Property Get SingleJoint() As String
    SingleJoint = SaveRD.SingleJoint
End Property
Public Property Get Joint1() As String
    Joint1 = SaveRD.Joint1
End Property
Public Property Get Joint2() As String
    Joint2 = SaveRD.Joint2
End Property
Public Property Get Joint3() As String
    Joint3 = SaveRD.Joint3
End Property
Public Property Get RateOfInterest() As Double
    RateOfInterest = SaveRD.RateOfInterest
End Property
Public Property Get Severally_Joint() As String
    Severally_Joint = SaveRD.Severally_Joint
End Property
Public Property Get Period() As Single
    Period = SaveRD.Period
End Property
Public Property Get Amount() As Double
    Amount = SaveRD.Amount
End Property
Public Property Get Maturity_Amount() As Double
    Maturity_Amount = SaveRD.Maturity_Amount
End Property

Public Property Get DateAccOpening() As Date
    DateAccOpening = SaveRD.DateAccOpening
End Property
Public Property Get MonthPeriod() As Single
    MonthPeriod = SaveRD.MonthPeriod
End Property
Public Property Get InterestAmount() As Double
    InterestAmount = SaveRD.InterestAmount
End Property
Public Property Get TerminalName() As String
    TerminalName = SaveRD.TerminalName
End Property
Public Property Get UserName() As String
    UserName = SaveRD.UserName
End Property
Public Property Get InsertModifyDate() As Date
 InsertModifyDate = SaveRD.InsertModifyDate
End Property
Public Function INSERTModifyDATA(ByRef objDatasource As ADODB.Connection, ByVal DoWork As Byte) As Boolean
    Dim objCommand As ADODB.Command
     
                Set objCommand = New ADODB.Command
                With objCommand
                If DoWork = 1 Then
                'Insert data in Dowork=1
                    .CommandText = "kccBank.PACK_Recuring_Deposit.INSERTDATA"
                ElseIf DoWork = 2 Then
                    'Modify data in Dowork=2
                    .CommandText = "kccBank.PACK_Recuring_Deposit.ModifyData"
                End If
                    .CommandType = adCmdStoredProc
                    'Send connection
                    .ActiveConnection = objDatasource
                    'Set command time out
                    .CommandTimeout = 0
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveRD.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveRD.AccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.NameOfAcHolder)
                If SaveRD.FatherName = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                 Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.FatherName)
                 End If
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveRD.Address)
                If SaveRD.PhNo = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, SaveRD.PhNo)
                End If
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveRD.AdultMinor)
                If SaveRD.Dob = Empty Then
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                Else
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveRD.Dob)
                End If
                If SaveRD.Guardian = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.Guardian)
                End If
                If SaveRD.DOBGur = Empty Then
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                Else
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveRD.DOBGur)
                End If
                If SaveRD.RelationGur = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveRD.RelationGur)
                End If
                If SaveRD.Nominee = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.Nominee)
                End If
                If SaveRD.DOBNom = Empty Then
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, Null)
                Else
                    .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveRD.DOBNom)
                End If
                If SaveRD.AddressNom = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200, SaveRD.AddressNom)
                End If
                If SaveRD.RelationNom = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 30, SaveRD.RelationNom)
                End If
                .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveRD.SingleJoint)
                If SaveRD.Joint1 = Empty Then
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                Else
                    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.Joint1)
                End If
                If SaveRD.Joint2 = Empty Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                Else
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.Joint2)
                End If
                If SaveRD.Joint3 = Empty Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, Null)
                Else
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.Joint3)
                End If
                If SaveRD.Severally_Joint = Empty Then
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, Null)
                Else
                    .Parameters.Append .CreateParameter(, adChar, adParamInput, 1, SaveRD.Severally_Joint)
                End If
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 3, SaveRD.Period)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 16, SaveRD.Amount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 2, SaveRD.RateOfInterest)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveRD.DateAccOpening)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveRD.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveRD.InsertModifyDate)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 15, SaveRD.Maturity_Amount)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveRD.TDS_Deduct)
                If SaveRD.TDSStatus = Empty Then
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , Null)
                Else
                    .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , SaveRD.TDSStatus)
                End If
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, SaveRD.RDLien)
                If DoWork = 2 Then
                  .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput, 1)
                End If
                    .Execute
                  If DoWork = 2 Then
                    If objCommand("Param1").Value = 0 Then
                    'Return true if Modify Successfully
                          INSERTModifyDATA = True
                    Else
                          INSERTModifyDATA = False
                    End If
                   Else
                   'Return true on insertion
                   INSERTModifyDATA = True
                   End If
                End With
                Set objCommand = Nothing
              End Function
Public Function CheckAccount(ByRef objDatasource As ADODB.Connection) As Boolean
    Dim objCommand As New ADODB.Command
    With objCommand
    'Check Account No. For Duplicacy
        .CommandText = "KccBank.PACK_Recuring_Deposit.CheckAccount"
        .CommandType = adCmdStoredProc
        'Send connection
        .ActiveConnection = objDatasource
        'Set command time out
        .CommandTimeout = 0
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, SaveRD.TypeOfAccount)
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, SaveRD.AccountNo)
        .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
        'Execute to fetch records
         .Execute
        If Not IsNull(objCommand("Param1").Value) Then
            If CLng(objCommand("Param1").Value) > 0 Then
            'if Already exists return True
                CheckAccount = True
            Else
                CheckAccount = False
            End If
        End If
    End With
    Set objCommand = Nothing
End Function
Public Function DeleteData(ByRef objDatasource As ADODB.Connection) _
                                As Boolean
    Dim objCommand As New ADODB.Command
        With objCommand
        'delete the Displayed Record
        .CommandText = "KccBank.PACK_Recuring_Deposit.DeleteData"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveRD.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveRD.AccountNo))
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, SaveRD.TerminalName)
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, SaveRD.UserName)
            .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, SaveRD.InsertModifyDate)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
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
    Dim look As New LookUp
    Dim strProcedure As String
    Dim sql As String
    sql = "Select TYPE_OF_ACCOUNT, ACCOUNT_NO, NAME_OF_AC_HOLDER, ADDRESS, PHONE_NO from kccbank.T_RD_Mst where Delete_Flag=" & "'N'" & "Order by ACCOUNT_NO, NAME_OF_AC_HOLDER"
   
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
            SaveRD.TypeOfAccount = (.DisplayValueByName("TYPE_OF_ACCOUNT"))
            SaveRD.AccountNo = CLng(.DisplayValueByName("ACCOUNT_NO"))
            PopupLov = True
            Exit Function
         
         End If
     PopupLov = False
    End With
End Function
Public Function GetACCOUNTDETAIL(ByRef objDatasource As ADODB.Connection) _
                                As ADODB.Recordset
'Get all fields of a Sinle record
    Dim objCommand As New ADODB.Command
        With objCommand
            .CommandText = "KccBank.PACK_Recuring_Deposit.GetACCOUNTDETAIL"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveRD.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 7, CLng(SaveRD.AccountNo))
            'Execute to fetch records
            Set GetACCOUNTDETAIL = .Execute
         End With
    Set objCommand = Nothing
End Function
Public Function GETAccNo(ByRef objDatasource As ADODB.Connection) _
                                As Long
    Dim objCommand As New ADODB.Command
        With objCommand
        'Get Max Account No. From database
            .CommandText = " KccBank.PACK_Recuring_Deposit.GETAccNo"
            .CommandType = adCmdStoredProc
            'Send connection
            .ActiveConnection = objDatasource
            'Set command time out
            .CommandTimeout = 0
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 2, SaveRD.TypeOfAccount)
            .Parameters.Append .CreateParameter("Param1", adNumeric, adParamOutput)
            'Execute to fetch records
             .Execute
                If Not IsNull(objCommand("Param1").Value) Then
                    GETAccNo = objCommand("Param1").Value
                End If
         End With
    Set objCommand = Nothing
End Function
