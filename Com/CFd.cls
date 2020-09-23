VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       CLASS      :CFD MASTER
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :HARISH CHAUHAN
'       MADE DATE  :01-04-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------
Private Type GroupData
       
        TypeOfAccount              As String
        AccountNo                  As Long
        Name                       As String
        Address                    As String
        Dob                        As Variant
        FdNo                       As Double
        AccountDate                As Date
        Jn1                        As String
        Jn2                        As String
        Jn3                        As String
        Jn4                        As String
        Jn5                        As String
        Nominee                    As String
        NomineeAddress             As String
        AdultMinor                 As String
        Guardian                   As String
        GuardianDOB                As Variant
        GuardianAddress            As String
        GuardianRelation           As String
        TerminalName               As String
        UserName                   As String
        InsertDate                 As Date
        ModifyDate                 As Date
        
    End Type
        Public Enum DBState
        DBInsert = 1
        DBModify = 2
        DBDelete = 3
    End Enum
    
    Private DBData                  As GroupData
    Public Function GetFdList(ByRef objDataSource As ADODB.Connection) As ADODB.Recordset
        Dim objCommand As New ADODB.Command
        'Dim objRecordset As ADODB.Recordset

        With objCommand
            .CommandText = "PACK_FD_MST.GETFDDATA()"
            .CommandType = adCmdStoredProc
            If DBData.TypeOfAccount = "MMC" Then
                DBData.TypeOfAccount = "MM"
            End If
            .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
            .Parameters.Append .CreateParameter(, adVarNumeric, adParamInput, , DBData.AccountNo)

                Set .ActiveConnection = objDataSource
                    .CommandTimeout = 0
'                Set objRecordset = .Execute
                Set GetFdList = .Execute
        End With
         '       Set objRecordset = Nothing
          '      Set objOLEServer = Nothing
                Set objCommand = Nothing
        Exit Function
End Function

    Public Property Let TypeOfAccount(ByVal v_Value As String)
        DBData.TypeOfAccount = v_Value
    End Property

    Public Property Let AccountNo(ByVal v_Value As Long)
       DBData.AccountNo = v_Value
    End Property
     
    Public Property Let Name(ByVal v_Value As String)
       DBData.Name = v_Value
    End Property
     
    Public Property Let Address(ByVal v_Value As String)
       DBData.Address = v_Value
    End Property
     
    Public Property Let Dob(ByVal v_Value As Variant)
       DBData.Dob = v_Value
    End Property
     
    Public Property Let FdNo(ByVal v_Value As Double)
       DBData.FdNo = v_Value
    End Property
    
    Public Property Let AccountDate(ByVal v_Value As Long)
       DBData.AccountDate = v_Value
    End Property
    
    Public Property Let Jn1(ByVal v_Value As String)
       DBData.Jn1 = v_Value
     End Property
        
    Public Property Let Jn2(ByVal v_Value As String)
       DBData.Jn2 = v_Value
     End Property
    Public Property Let Jn3(ByVal v_Value As String)
       DBData.Jn3 = v_Value
     End Property
    Public Property Let Jn4(ByVal v_Value As String)
       DBData.Jn4 = v_Value
     End Property
    Public Property Let Jn5(ByVal v_Value As String)
       DBData.Jn5 = v_Value
     End Property
    
    Public Property Let Nominee(ByVal v_Value As String)
       DBData.Nominee = v_Value
     End Property
    
    Public Property Let NomineeAddress(ByVal v_Value As String)
       DBData.NomineeAddress = v_Value
     End Property
    
    Public Property Let AdultMinor(ByVal v_Value As String)
       DBData.AdultMinor = v_Value
     End Property
     
     Public Property Get AdultMinor() As String
        AdultMinor = DBData.AdultMinor
    End Property
     
     Public Property Let Guardian(ByVal v_Value As String)
       DBData.Guardian = v_Value
     End Property
    
     Public Property Let GuardianDOB(ByVal v_Value As Variant)
       DBData.GuardianDOB = v_Value
     End Property
       
       Public Property Let GuardianAddress(ByVal v_Value As String)
       DBData.GuardianAddress = v_Value
      End Property
    
     Public Property Let GuardianRelation(ByVal v_Value As String)
       DBData.GuardianRelation = v_Value
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
        
Public Sub SaveData(ByRef objDataSource As ADODB.Connection, _
                        ByVal Recordstate As DBState)
    'This Procedure Inserts and Modify Data depending upon the Mode
    Dim objCommand As New ADODB.Command
    
    Dim m_str   As String
    With objCommand
            If Recordstate = DBInsert Then
                .CommandText = "PACK_FD_MST.INSERT_FDDATA"
            ElseIf Recordstate = DBModify Then
                .CommandText = "PACK_FD_MST.MODIFY_FDDATA"
            ElseIf Recordstate = DBDelete Then
                .CommandText = "PACK_FD_MST.DELETE_FDDATA"
            End If
            .CommandType = adCmdStoredProc
          
          
            If m_DtCount > 0 Then
                MsgBox "Please enter date in format DD-MM-YYYY", vbInformation, "Bank Error"
                
            End If
            Set .ActiveConnection = objDataSource
            .CommandTimeout = 0
            If Recordstate = DBInsert Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.FdNo)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.AccountDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 60, DBData.Name)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 150, DBData.Address)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.Dob)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn1)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn2)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn3)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn4)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn5)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Nominee)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.NomineeAddress)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.AdultMinor)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.Guardian)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.GuardianAddress)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.GuardianDOB)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.GuardianRelation)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
           End If
           If Recordstate = DBModify Then
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.AccountDate)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 60, DBData.Name)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 150, DBData.Address)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.Dob)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn1)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn2)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn3)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn4)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Jn5)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.Nominee)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.NomineeAddress)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 1, DBData.AdultMinor)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.Guardian)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.GuardianAddress)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, DBData.GuardianDOB)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.GuardianRelation)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, DBData.TerminalName)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 100, DBData.UserName)
                .Parameters.Append .CreateParameter(, adDate, adParamInput, 8, CDate(Format(DBData.InsertDate, "dd-mm-yyyy")))
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, 6, DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.FdNo)
          End If
          If Recordstate = DBDelete Then
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 2, DBData.TypeOfAccount)
                .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , DBData.AccountNo)
                .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 20, DBData.FdNo)
          End If
               
                                                                                                                                                                                                                                                       
                                                                                                                                                                                                                                                       .Execute
    End With
            Set objCommand = Nothing
End Sub



