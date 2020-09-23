VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccountHeadMaster 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve Account Heads"
      Height          =   345
      Left            =   5280
      TabIndex        =   24
      Top             =   5535
      Width           =   2070
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   345
      Left            =   3015
      TabIndex        =   4
      ToolTipText     =   "Click to close this screen"
      Top             =   5535
      Width           =   885
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8820
      TabIndex        =   17
      ToolTipText     =   "Click to cancel mode"
      Top             =   4935
      Width           =   885
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   7890
      TabIndex        =   16
      ToolTipText     =   "Click to save added/modified units"
      Top             =   4935
      Width           =   885
   End
   Begin VB.Frame framain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   4005
      TabIndex        =   19
      Top             =   540
      Width           =   5700
      Begin Branch.NumberControl numOpeningBal 
         Height          =   255
         Left            =   1545
         TabIndex        =   33
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame FraAgriLoan 
         Caption         =   "ST Agri Loan Details"
         Height          =   1770
         Left            =   45
         TabIndex        =   25
         Top             =   2445
         Width           =   5625
         Begin Branch.NumberControl numMaxKind 
            Height          =   255
            Left            =   4080
            TabIndex        =   39
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Branch.NumberControl numMaxCash 
            Height          =   255
            Left            =   4080
            TabIndex        =   38
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Branch.NumberControl numRabiKind 
            Height          =   255
            Left            =   4080
            TabIndex        =   37
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Branch.NumberControl numRabiCash 
            Height          =   255
            Left            =   1440
            TabIndex        =   36
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Branch.NumberControl numKharifKind 
            Height          =   255
            Left            =   1440
            TabIndex        =   35
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Branch.NumberControl numKharifCash 
            Height          =   255
            Left            =   1440
            TabIndex        =   34
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtMCL 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1395
            MaxLength       =   20
            TabIndex        =   15
            Top             =   270
            Width           =   2850
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Kind"
            Height          =   195
            Index           =   3
            Left            =   2955
            TabIndex        =   32
            Top             =   1425
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Cash"
            Height          =   195
            Index           =   3
            Left            =   2910
            TabIndex        =   31
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rabi Kind"
            Height          =   195
            Index           =   2
            Left            =   3285
            TabIndex        =   30
            Top             =   705
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rabi Cash"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   29
            Top             =   1425
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kharif Kind"
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   28
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kharif Cash"
            Height          =   195
            Index           =   0
            Left            =   465
            TabIndex        =   27
            Top             =   705
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MCL Period"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   26
            Top             =   345
            Width           =   825
         End
      End
      Begin VB.CheckBox ChkAgriLoan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Is it a Agri-Loan ?"
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   1335
         Width           =   1575
      End
      Begin VB.CheckBox ChkPL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Is it a PL ?"
         Height          =   285
         Left            =   660
         TabIndex        =   10
         Top             =   1035
         Width           =   1095
      End
      Begin VB.TextBox txtAcCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1530
         MaxLength       =   5
         TabIndex        =   6
         Top             =   0
         Width           =   1305
      End
      Begin VB.ComboBox cmbSubGroup 
         Height          =   315
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1995
         Width           =   2835
      End
      Begin VB.CheckBox chkSubGroup 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Is it a Sub-Group ?"
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   750
         Width           =   1680
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1530
         MaxLength       =   60
         TabIndex        =   8
         ToolTipText     =   "Enter Description"
         Top             =   360
         Width           =   4170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00D5C0AE&
         BackStyle       =   0  'Transparent
         Caption         =   "&Account Code"
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   90
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group/SubGroup"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   2085
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label lblSubConveyance 
         AutoSize        =   -1  'True
         BackColor       =   &H00D5C0AE&
         BackStyle       =   0  'Transparent
         Caption         =   "&Description"
         Height          =   195
         Left            =   645
         TabIndex        =   7
         Top             =   450
         Width           =   795
      End
   End
   Begin VB.CommandButton CmdInsert 
      Caption         =   "&Insert"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Click to enter into insert mode"
      Top             =   5535
      Width           =   885
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "&Modify"
      Height          =   345
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Click to enter into modify mode"
      Top             =   5535
      Width           =   885
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   2055
      TabIndex        =   3
      ToolTipText     =   "Click to delete P&D Item"
      Top             =   5535
      Width           =   870
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "&Help"
      Height          =   345
      Left            =   8775
      TabIndex        =   18
      ToolTipText     =   "Click to get help"
      Top             =   5535
      Width           =   885
   End
   Begin MSComctlLib.ListView lsvAccHead 
      Height          =   4980
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "List of Account Heads"
      Top             =   420
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   8784
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ac_Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group_Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SubGroup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "opening_balance"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "GroupCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "IS PL"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ParentKey"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "isAgriLoan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "MCL"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "KharifCash"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "KharifKind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "RabiCash"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "RabiKind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "MaxCash"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "MaxKind"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblBottomRight 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8D7D1&
      Height          =   240
      Left            =   6435
      TabIndex        =   23
      Top             =   5985
      Width           =   3180
   End
   Begin VB.Label lblBottomLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KCCB - Head Office, Thanesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8D7D1&
      Height          =   240
      Left            =   75
      TabIndex        =   22
      Top             =   5985
      Width           =   3135
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D8D7D1&
      Height          =   240
      Left            =   105
      TabIndex        =   21
      Top             =   45
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   6315
      TabIndex        =   20
      Top             =   60
      Width           =   3360
   End
   Begin VB.Shape shpBottom 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   15
      Top             =   5985
      Width           =   9765
   End
   Begin VB.Shape shpTop 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   45
      Width           =   9765
   End
End
Attribute VB_Name = "frmAccountHeadMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************
'Form Name          :- frmAccountHeadMaster.frm
'Pupose             :- Form created for generating new Subgroups and Account Heads
'Referred dll       :- KCCBAccHeadMst.cls
'Date of Creation   :- 22nd March'2002
'Developed By       :- Kiran Kanwar
'Revisions          :-
'Copyright(c)2002-2003 SURYA INFONET LTD.
'**********************************************

Option Explicit

Private Const DEFAULT_DELIMITER As String = "^"  'Required to fill multiple fields in Combo box
Dim bytButton       As Byte                      'To define Add/Modify stage
Dim m_ObjRes        As New LoadRes               'To call Message Dll
Dim m_lngAccountKey As Long                      'To save Account Key
Dim SelectedIndex   As Long                      'To save index of selected item in Listview
Dim SubGroupVal     As Byte                      'To save that modified Account head was actually subgroup or account head
Dim m_strGroupNo    As String                    'To save Group Number
Dim m_strParentno   As String                    'to save Parent No.
Dim i               As Long
Dim m_strAgriLoan   As String                    'to save whether group is Agri Loan or Not

Private Sub Buttons_Module(strFlag As String)
'Purpose :- This function helps in enabling/disabling buttons
'           at different stages including Add/Modify/Save/Cancel

    If strFlag = "Add" Then
        cmdOk.Enabled = True
        cmdcancel.Enabled = True
        cmdModify.Enabled = False
        cmdInsert.Enabled = False
        cmdDelete.Enabled = False
        cmdExit.Enabled = False
        CmdHelp.Enabled = False
        framain.Enabled = True
        lsvAccHead.Enabled = False
    ElseIf strFlag = "Save" Then
        cmdOk.Enabled = False
        cmdcancel.Enabled = False
        cmdInsert.Enabled = True
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
        cmdExit.Enabled = True
        CmdHelp.Enabled = True
        framain.Enabled = False
        lsvAccHead.Enabled = True
        lsvAccHead.SetFocus
    ElseIf strFlag = "Cancel" Then
    End If
End Sub

Private Sub ChkAgriLoan_Click()
    If ChkAgriLoan.Value = 1 Then
        FraAgriLoan.Enabled = True
    ElseIf ChkAgriLoan.Value = 0 Then
        FraAgriLoan.Enabled = False
    End If
End Sub

Private Sub ChkPL_Click()
    If ChkPL.Value = 1 Then
        numOpeningBal.Locked = True
        numOpeningBal.BackColor = &HE0E0E0
        numOpeningBal.TabStop = False
        numOpeningBal.Text = "0.00"
    Else
        numOpeningBal.Locked = False
        numOpeningBal.BackColor = &H80000005
        numOpeningBal.TabStop = True
    End If
End Sub

Private Sub chkSubGroup_Click()
' If Sub group check box is checked - Opening balance field should be disabled
' If Sub group check box is not checked - Opening balance field should not be disabled
    
    If chkSubGroup.Value = 1 Then
        numOpeningBal.Locked = True
        numOpeningBal.BackColor = &HE0E0E0
        numOpeningBal.TabStop = False
        numOpeningBal.Text = "0.00"
    Else
        numOpeningBal.Locked = False
        numOpeningBal.BackColor = &H80000005
        numOpeningBal.TabStop = True
    End If
End Sub

Private Function GetComboGroupCode(ByVal v_ComboText As String, _
                             ByVal Delimiter As String) _
                             As String
'Purpose :- This function helps retrieving Groupcode from selected combobox text
                             
    Dim charAtPosition As Byte
    If Len(v_ComboText) > 0 Then
        charAtPosition = InStr(1, v_ComboText, Delimiter)
        GetComboGroupCode = Mid(v_ComboText, charAtPosition, Len(v_ComboText))
        GetComboGroupCode = Trim(Mid(GetComboGroupCode, 2, Len(GetComboGroupCode)))
    End If
End Function

Private Sub cmbSubGroup_Change()
'    GetAgriLoanORNot
End Sub
Private Sub GetAgriLoanORNot()

    Dim objCommand As New ADODB.Command
    Dim objRecordset As New ADODB.Recordset
    
  
    With objCommand
        Set .ActiveConnection = g_objDataSource.GetDataSource
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(, adNumeric, adParamInput, , GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER))
        .CommandText = "Kccbank.PACK_ACHEADBranch_Data.GetAgriLoanOrLoan()"
        Set objRecordset = .Execute
    End With
    If objRecordset.RecordCount > 0 Then
        If objRecordset.Fields("is_agriloan") = "Y" Then
            FraAgriLoan.Enabled = True
        Else
            FraAgriLoan.Enabled = False
        End If
    Else
        FraAgriLoan.Enabled = False
    End If
End Sub

Private Sub cmbSubGroup_Click()
'    If cmbSubGroup.Text <> "" Then GetAgriLoanORNot
End Sub

Private Sub cmbSubGroup_LostFocus()
    If cmbSubGroup.Text <> "" Then
        GetAgriLoanORNot
        If FraAgriLoan.Enabled = True Then txtMCL.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    
    Buttons_Module "Save"
    If bytButton = INSERT Then
        If lsvAccHead.ListItems.Count > 0 Then
            lsvAccHead_ItemClick lsvAccHead.ListItems(1)
        Else
            ClearControls
        End If
    ElseIf bytButton = MODIFY Then
        If lsvAccHead.ListItems.Count > 0 Then lsvAccHead_ItemClick lsvAccHead.ListItems(1)
    End If
    bytButton = 0
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DelErr:
Dim selIndex As Long
Dim blnSuccess As Boolean
Dim i As Long

If lsvAccHead.ListItems.Count > 0 Then
    If CheckRecord = False Then
        Screen.MousePointer = vbDefault
        GoTo DelErr:
        Exit Sub
    End If
    If MsgBox(m_ObjRes.LoadStringFromDLL(1006), vbQuestion + vbYesNo, "Please Confirm") = vbYes Then
        
        Dim objDelete As New KCCBAccHeadMst.CAccHeadMst
        With objDelete
            .AccKey = lsvAccHead.SelectedItem.SubItems(2)
            .ISDeleted = "Y"
            .TerminalName = strComputerName
            .UserName = strUserName
            .InsertModifyDate = dteTodaysDate
            .AccHeadData g_objDataSource.GetDataSource, DBDelete
        End With
        Set objDelete = Nothing
        
        blnSuccess = GetSubGroups()
        
        selIndex = lsvAccHead.SelectedItem.Index
        lsvAccHead.ListItems.Remove selIndex
        If lsvAccHead.ListItems.Count >= 1 Then
            lsvAccHead_ItemClick lsvAccHead.ListItems(1)
            lsvAccHead.ListItems(1).Selected = True
            lsvAccHead.SetFocus
        ElseIf lsvAccHead.ListItems.Count = 0 Then
            ClearControls
        End If
    End If
End If
Exit Sub
DelErr:
   MsgBox "Some transactions exists against selected Account Head/Subgroup.", vbInformation, "Please Check"
   lsvAccHead.SetFocus
   Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    ClearControls
    bytButton = INSERT
    Buttons_Module "Add"
    chkSubGroup_Click
    txtAcCode.SetFocus
End Sub

Private Sub ClearControls()
'To clear values of all entry controls  when 'Insert' button is clicked
    txtAcCode.Text = ""
    txtDescription.Text = ""
    chkSubGroup.Value = 0
    cmbSubGroup.ListIndex = -1
    ChkPL.Value = 0
    numOpeningBal.Text = "0.00"
    ChkAgriLoan.Value = 0
    FraAgriLoan.Enabled = False
    txtMCL.Text = ""
    numKharifCash.Text = "0.00"
    numKharifKind.Text = "0.00"
    numRabiCash.Text = "0.00"
    numRabiKind.Text = "0.00"
    numMaxCash.Text = "0.00"
    numMaxKind.Text = "0.00"
End Sub

Private Sub cmdModify_Click()
    If lsvAccHead.ListItems.Count > 0 Then
        lsvAccHead_ItemClick lsvAccHead.ListItems(lsvAccHead.SelectedItem.Index)
        chkSubGroup_Click
        bytButton = MODIFY
        Buttons_Module "Add"
        SubGroupVal = chkSubGroup.Value
        If ChkAgriLoan.Value = 1 Then
            FraAgriLoan.Enabled = True
        ElseIf ChkAgriLoan.Value = 0 Then
            FraAgriLoan.Enabled = False
        End If
        Call GetAgriLoanORNot
        txtAcCode.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err:
Dim blnSuccess As Boolean
    If DoValidation = False Then
        Screen.MousePointer = vbDefault
        On Error GoTo 0
        Exit Sub
    End If

    If bytButton = INSERT Then
        If CheckUniqueRecord = False Then
            Screen.MousePointer = vbDefault
            On Error GoTo 0
            Exit Sub
        End If
        If chkSubGroup.Value = 1 Then
            If CheckNoSubGroups = False Then
                Screen.MousePointer = vbDefault
                On Error GoTo 0
                Exit Sub
            End If
        End If
        Call GetMaxGroupNo
        Call InsertAccHead
        Call UpdateList
        cmbSubGroup.Clear
        blnSuccess = GetSubGroups()
        If Len(Trim(lsvAccHead.SelectedItem.SubItems(3))) > 0 Then
            For i = 0 To cmbSubGroup.ListCount - 1
                cmbSubGroup.ListIndex = i

                If Trim(lsvAccHead.SelectedItem.SubItems(3)) = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) Then
'                    MsgBox cmbSubGroup.Text
                    Exit For
                End If
            Next
        End If
    ElseIf bytButton = MODIFY Then
        If UCase(txtDescription.Text) <> UCase(lsvAccHead.SelectedItem.SubItems(1)) Then
            If CheckUniqueRecord = False Then
                Screen.MousePointer = vbDefault
                On Error GoTo 0
                Exit Sub
            End If
        End If
        If SubGroupVal = 1 And chkSubGroup.Value = 0 Then
            If CheckRecord = False Then
                Screen.MousePointer = vbDefault
                GoTo DelErr:
                Exit Sub
            End If
        End If
        'MsgBox lsvAccHead.SelectedItem.SubItems(8)
        If chkSubGroup.Value = 1 Then
            If Mid(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), 1, InStr(1, GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) - 2) <> lsvAccHead.SelectedItem.SubItems(8) Then
                If CheckNoSubGroups = False Then
                    Screen.MousePointer = vbDefault
                    On Error GoTo 0
                    Exit Sub
                End If
            End If
        End If
'        Call GetMaxGroupNo
        Call InsertAccHead
        Call UpdateList
        cmbSubGroup.Clear
        blnSuccess = GetSubGroups()
        If Len(Trim(lsvAccHead.SelectedItem.SubItems(3))) > 0 Then
            For i = 0 To cmbSubGroup.ListCount - 1
                cmbSubGroup.ListIndex = i

                If Trim(lsvAccHead.SelectedItem.SubItems(3)) = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) Then
'                    MsgBox cmbSubGroup.Text
                    Exit For
                End If
            Next
        End If
    End If
    Buttons_Module "Save"
    bytButton = 0
Exit Sub
DelErr:
    MsgBox "Cannot modify given Sub-Group into Account Head as it already has some Account Heads under it.", vbInformation, "KCCB"
    chkSubGroup.SetFocus
    Exit Sub
Err:
    If Err.Number = -2147217900 Then
        MsgBox "Please check that Account Code should not be duplicated.", vbInformation, "KCCB"
        txtAcCode.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        MsgBox "Error Number is -" & Err.Number
        MsgBox "Error Description is -" & Err.Description
    End If
End Sub

Private Sub cmdRetrieve_Click()
    Dim TALook As New Lov.LookUp
    Dim strProc As String
    
    strProc = "{ Call kccbank.PACK_ACHeadBranch_data.GetDeletedAcHeads()}"
    With TALook
        .AddColumnHeaders "Account Code", "Description", "Group_Key", "Group_Number", "Parent_Number", "SubGroup", "ISPL", "Opening_Balance", "Acc_Key"
        .AddDisplayFields "Acc_Code", "Acc_Description", "Group_Key", "Group_Number", "Parent_Number", "SubGroup", "ISPL", "Opening_Balance", "Acc_Key"
        .SetColumnsWidth 1500, 3500, 0, 0, 0, 0, 0, 0, 0
        .Connection = g_objDataSource.GetDataSource
        .ProcedureText = strProc
        .TotalColumns = 9
        .PopulateList
        If .LOVState <> True Then
            ''Code to update database to change delete status back to Not Deleted
    
            Dim objCommand As New ADODB.Command
            Set objCommand.ActiveConnection = g_objDataSource.GetDataSource
            objCommand.CommandType = adCmdStoredProc
            objCommand.Parameters.Append objCommand.CreateParameter("param1", adNumeric, adParamInput, , .DisplayValueByName("Acc_Key"))
            objCommand.CommandText = "Kccbank.Pack_ACHeadBranch_DATA.UpdateISdeleted()"
            objCommand.Execute
            
            Dim lst As ListItem
            Set lst = lsvAccHead.ListItems.Add(, , .DisplayValueByName("Acc_Code"))
            lst.SubItems(1) = .DisplayValueByName("Acc_Description")
            lst.SubItems(2) = .DisplayValueByName("Acc_Key")
            lst.SubItems(3) = .DisplayValueByName("Group_Key")
            lst.SubItems(4) = .DisplayValueByName("SubGroup")
            lst.SubItems(5) = .DisplayValueByName("Opening_Balance")
            lst.SubItems(6) = .DisplayValueByName("Group_Number")
            lst.SubItems(7) = .DisplayValueByName("ISPL")
            lst.SubItems(8) = .DisplayValueByName("Parent_Number")
            lst.SubItems(9) = .DisplayValueByName("IS_AgriLoan")
            lst.SubItems(10) = .DisplayValueByName("MCL_Period")
            lst.SubItems(11) = .DisplayValueByName("Kharif_Cash")
            lst.SubItems(12) = .DisplayValueByName("Kharif_Kind")
            lst.SubItems(13) = .DisplayValueByName("Rabi_Cash")
            lst.SubItems(14) = .DisplayValueByName("Rabi_Kind")
            lst.SubItems(15) = .DisplayValueByName("Max_Cash")
            lst.SubItems(16) = .DisplayValueByName("Max_Kind")
        Else
        End If
    End With
End Sub

Private Sub Form_Load()

    Dim blnSuccess As Boolean
    blnSuccess = GetSubGroups()
    Call FillAccountHeads
    framain.Enabled = False
    m_ObjRes.DllName = App.Path & "\DLLs\Resources\kccbres.dll"
    Label1.Caption = Format(strTodaysDate, "dddddd dd MMM yyyy")
    lblBottomRight.Caption = Trim(strUserName)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
        If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessageLong(Me.hWnd, &HA1, 2, 0&)
        End If
End Sub

Private Sub lsvAccHead_DblClick()
    If lsvAccHead.ListItems.Count > 0 Then cmdModify_Click
End Sub

Private Sub lsvAccHead_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim i As Long
    If lsvAccHead.ListItems.Count > 0 Then
        txtAcCode.Text = lsvAccHead.SelectedItem.Text
        txtDescription.Text = lsvAccHead.SelectedItem.SubItems(1)
        chkSubGroup.Value = IIf(lsvAccHead.SelectedItem.SubItems(4) = "Y", 1, 0)
        ChkPL.Value = IIf(lsvAccHead.SelectedItem.SubItems(7) = "Y", 1, 0)
        numOpeningBal.Text = Format(lsvAccHead.SelectedItem.SubItems(5), "#0.00")
        m_strParentno = lsvAccHead.SelectedItem.SubItems(8)
        m_strGroupNo = lsvAccHead.SelectedItem.SubItems(6)
        If Len(Trim(lsvAccHead.SelectedItem.SubItems(3))) > 0 Then
            For i = 0 To cmbSubGroup.ListCount - 1
                cmbSubGroup.ListIndex = i
'                MsgBox cmbSubGroup.Text
                If Trim(lsvAccHead.SelectedItem.SubItems(3)) = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) Then
'                    MsgBox cmbSubGroup.Text
                    Exit For
                End If
            Next
        End If
        ChkAgriLoan.Value = IIf(lsvAccHead.SelectedItem.SubItems(9) = "Y", 1, 0)
        txtMCL.Text = lsvAccHead.SelectedItem.SubItems(10)
        numKharifCash.Text = Format(Val(lsvAccHead.SelectedItem.SubItems(11)), "#0.00")
        numKharifKind.Text = Format(Val(lsvAccHead.SelectedItem.SubItems(12)), "#0.00")
        numRabiCash.Text = Format(Val(lsvAccHead.SelectedItem.SubItems(13)), "#0.00")
        numRabiKind.Text = Format(Val(lsvAccHead.SelectedItem.SubItems(14)), "#0.00")
        numMaxCash.Text = Format(Val(lsvAccHead.SelectedItem.SubItems(15)), "#0.00")
        numMaxKind.Text = Format(Val(lsvAccHead.SelectedItem.SubItems(16)), "#0.00")
    End If
End Sub

Private Function GetComboKey(ByVal v_ComboText As String, _
                             ByVal Delimiter As String) _
                             As String
''Procedure to get key(which is usually second column in combo value) of selected item
    Dim charAtPosition As Byte
    If Len(v_ComboText) > 0 Then
        charAtPosition = InStr(1, v_ComboText, Delimiter)
        GetComboKey = Mid(v_ComboText, charAtPosition, Len(v_ComboText))
        GetComboKey = Trim(Replace(GetComboKey, Delimiter, ""))
    End If
End Function
'
'Private Sub FillSubGroups()
'    Dim objAccountHead As New KCCBAccHeadMst.CAccHeadMst
'    Dim objRecordset As ADODB.Recordset
'
'    Set objRecordset = objAccountHead.GetGroupList(g_objDataSource.GetDataSource, Me.hWnd)
'    cmbSubGroup.Clear
'    Call FillCmb(objRecordset)
'
'    Set objAccountHead = Nothing
'
'End Sub
'
'Private Sub FillCmb(ByRef objRecordset As ADODB.Recordset)
'
'    If objRecordset.RecordCount > 0 Then
'        While Not objRecordset.EOF
'            FillComboBox cmbSubGroup, objRecordset.Fields("Group_Key"), objRecordset.Fields("abc")
'            objRecordset.MoveNext
'        Wend
'        Set objRecordset = Nothing
'        Exit Sub
'    End If
'End Sub
'
'Private Sub FillComboBox(ByRef Combo_Name As ComboBox, _
'                         ByVal IndexField As String, _
'                         ByVal Value As String)
'
'    With Combo_Name
'        .AddItem Trim(Value) & Space(65) & " ^ " & IndexField
'    End With
'
'End Sub
Private Sub FillAccountHeads()
''Procedure to get all accounts and their details
    Dim objAccountHead As New KCCBAccHeadMst.CAccHeadMst
    Dim objRecordset As ADODB.Recordset
    
    Set objRecordset = objAccountHead.GetAccountHeadList(g_objDataSource.GetDataSource)
    lsvAccHead.ListItems.Clear
    If objRecordset.RecordCount > 0 Then
        Call FillList(objRecordset)
    End If
    
    Set objAccountHead = Nothing
End Sub

Private Sub FillList(ByRef objRecordset As ADODB.Recordset)
    ''Procedure to populate subgroups/account heads in list
    Dim objListItem As ListItem
    
    If objRecordset.RecordCount > 0 Then
        While Not objRecordset.EOF
            Set objListItem = lsvAccHead.ListItems.Add
            
            With objListItem
                .Text = objRecordset.Fields("Acc_Code")
                .SubItems(1) = objRecordset.Fields("Acc_Description")
                .SubItems(2) = objRecordset.Fields("Acc_Key")
                .SubItems(3) = objRecordset.Fields("Group_Key")
                .SubItems(4) = objRecordset.Fields("Subgroup")
                .SubItems(5) = objRecordset.Fields("Opening_Balance")
                .SubItems(6) = objRecordset.Fields("Group_Number")
                .SubItems(7) = objRecordset.Fields("ISPL")
                .SubItems(8) = objRecordset.Fields("Parent_Number")
                .SubItems(9) = IIf(IsNull(objRecordset.Fields("IS_AgriLoan")), "N", "Y")
                .SubItems(10) = IIf(IsNull(objRecordset.Fields("MCL_Period")), "", objRecordset.Fields("MCL_Period"))
                .SubItems(11) = IIf(IsNull(objRecordset.Fields("Kharif_Cash")), 0, objRecordset.Fields("Kharif_Cash"))
                .SubItems(12) = IIf(IsNull(objRecordset.Fields("Kharif_Kind")), 0, objRecordset.Fields("Kharif_Kind"))
                .SubItems(13) = IIf(IsNull(objRecordset.Fields("Rabi_Cash")), 0, objRecordset.Fields("Rabi_Cash"))
                .SubItems(14) = IIf(IsNull(objRecordset.Fields("Rabi_Kind")), 0, objRecordset.Fields("Rabi_Kind"))
                .SubItems(15) = IIf(IsNull(objRecordset.Fields("Max_Cash")), 0, objRecordset.Fields("Max_Cash"))
                .SubItems(16) = IIf(IsNull(objRecordset.Fields("Max_Kind")), 0, objRecordset.Fields("Max_Kind"))
            End With
            Set objListItem = Nothing
            objRecordset.MoveNext
        Wend
        lsvAccHead_ItemClick lsvAccHead.ListItems(1)
        Set objRecordset = Nothing
        Exit Sub
    End If
End Sub

Private Function DoValidation() As Boolean
'Purpose    :-  This function before saving data checks that entry in all the fields
'               is proper and relevant
    With m_ObjRes
        If txtAcCode.Text = "" Then
            .AddDynamicParameter = "Account's Code"
            MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1003)), vbInformation, "KCCB"
            txtAcCode.SetFocus
            Exit Function
        ElseIf txtDescription.Text = "" Then
            .AddDynamicParameter = "Account's Description"
            MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1003)), vbInformation, "KCCB"
            txtDescription.SetFocus
            Exit Function
        ElseIf chkSubGroup.Value = 0 And ChkPL.Value = 0 And Val(numOpeningBal.Text) = 0 Then
            .AddDynamicParameter = "Opening Balance"
            MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1003)), vbInformation, "KCCB"
            numOpeningBal.SetFocus
            Exit Function
        ElseIf cmbSubGroup.Text = "" Then
            .AddDynamicParameter = "Group/Sub-Group"
            MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1003)), vbInformation, "KCCB"
            cmbSubGroup.SetFocus
            Exit Function
        ElseIf UCase(txtDescription.Text) = UCase(GetComboValue(cmbSubGroup.Text, DEFAULT_DELIMITER)) Then
            .AddDynamicParameter = "Group/Sub-Group"
            MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
            cmbSubGroup.SetFocus
            Exit Function
        ElseIf m_strAgriLoan = "Y" Then
            If Trim(txtMCL.Text) = "" Then
                .AddDynamicParameter = "MCL Period"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                txtMCL.SetFocus
                Exit Function
            ElseIf Val(numKharifCash.Text) = "" Then
                .AddDynamicParameter = "Kharif Cash"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                numKharifCash.SetFocus
                Exit Function
            ElseIf Val(numKharifKind.Text) = "" Then
                .AddDynamicParameter = "Kharif Kind"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                numKharifKind.SetFocus
                Exit Function
            ElseIf Val(numRabiCash.Text) = "" Then
                .AddDynamicParameter = "Rabi Cash"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                numRabiCash.SetFocus
                Exit Function
            ElseIf Val(numRabiKind.Text) = "" Then
                .AddDynamicParameter = "Rabi Kind"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                numRabiKind.SetFocus
                Exit Function
            ElseIf Val(numMaxCash.Text) = "" Then
                .AddDynamicParameter = "Maximum Cash"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                numMaxCash.SetFocus
                Exit Function
            ElseIf Val(numMaxKind.Text) = "" Then
                .AddDynamicParameter = "Maximum Kind"
                MsgBox .ReplaceStringByParam(.LoadStringFromDLL(1005)), vbInformation, "KCCB"
                numMaxKind.SetFocus
                Exit Function
            End If
        End If
    End With
    DoValidation = True
End Function

Private Function CheckUniqueRecord() As Boolean
    ''Procedure to check uniqueness of inserted/modified Account description
    Dim objCheck As New KCCBAccHeadMst.CAccHeadMst
    objCheck.AccDescription = UCase(txtDescription.Text)
    objCheck.AccHeadData g_objDataSource.GetDataSource, DBCheckUnique
    
    If objCheck.IsUnique Then
    ElseIf Not objCheck.IsUnique Then
        With m_ObjRes
            .AddDynamicParameter = "Account Description"
            MsgBox .ReplaceStringByParam(m_ObjRes.LoadStringFromDLL(1009))
        End With
        txtDescription.SetFocus
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Set objCheck = Nothing
    CheckUniqueRecord = True
End Function

Private Function CheckRecord() As Boolean
''Procedure to check that selected sub group is already having some children or not
    Dim objCheck As New KCCBAccHeadMst.CAccHeadMst
    objCheck.AccKey = lsvAccHead.SelectedItem.SubItems(2)
    objCheck.AccHeadData g_objDataSource.GetDataSource, DBCheckRecord
    
    If Not objCheck.IsUnique Then
        CheckRecord = False
        Exit Function
    End If
    
    Set objCheck = Nothing
    CheckRecord = True
End Function

Private Sub InsertAccHead()
''Procedure to insert/update data
    Dim objInsertAcc As New KCCBAccHeadMst.CAccHeadMst
    Dim i As Long
    Dim OriginalGroup As String
    
    If bytButton = INSERT Then
        With objInsertAcc
            .GrpCode = Trim(m_strGroupNo)
            .ParentNo = m_strParentno
            .AccCode = Trim(txtAcCode.Text)
            .AccDescription = Trim(txtDescription.Text)
            .Subgroup = IIf(chkSubGroup.Value = 1, "Y", "N")
            .ISPL = IIf(ChkPL.Value = 1, "Y", "N")
            .OpeningBalance = Val(numOpeningBal.Text)
            .AgriLoan = IIf(ChkAgriLoan.Value = 1, "Y", "N")
            .MCL = txtMCL.Text
            .KharifCash = Val(numKharifCash.Text)
            .KharifKind = Val(numKharifKind.Text)
            .RabiCash = Val(numRabiCash.Text)
            .RabiKind = Val(numRabiKind.Text)
            .MaxCash = Val(numMaxCash.Text)
            .MAxKind = Val(numMaxKind.Text)
            .GroupKey = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER)
            .TerminalName = strComputerName
            .UserName = strUserName
            .InsertModifyDate = dteTodaysDate
            .AccHeadData g_objDataSource.GetDataSource, DBInsert
            m_lngAccountKey = .AccountCurrentValue
        End With
        If chkSubGroup.Value = 1 Then
            cmbSubGroup.AddItem Trim(txtDescription.Text) & Space(65) & " ^ " & m_lngAccountKey
        End If
    ElseIf bytButton = MODIFY Then
        With objInsertAcc
            .GrpCode = Trim(m_strGroupNo)
            .ParentNo = m_strParentno
            .AccCode = Trim(txtAcCode.Text)
            .AccKey = lsvAccHead.SelectedItem.SubItems(2)
            .AccDescription = Trim(txtDescription.Text)
            .Subgroup = IIf(chkSubGroup.Value = 1, "Y", "N")
            .ISPL = IIf(ChkPL.Value = 1, "Y", "N")
            .OpeningBalance = Val(numOpeningBal.Text)
            .AgriLoan = IIf(ChkAgriLoan.Value = 1, "Y", "N")
            .MCL = txtMCL.Text
            .KharifCash = Val(numKharifCash.Text)
            .KharifKind = Val(numKharifKind.Text)
            .RabiCash = Val(numRabiCash.Text)
            .RabiKind = Val(numRabiKind.Text)
            .MaxCash = Val(numMaxCash.Text)
            .MAxKind = Val(numMaxKind.Text)
            .GroupKey = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER)
            .TerminalName = strComputerName
            .UserName = strUserName
            .InsertModifyDate = dteTodaysDate
            .AccHeadData g_objDataSource.GetDataSource, DBModify
        End With
'        If SubGroupVal = 0 And chkSubGroup.Value = 1 Then
'            cmbSubGroup.AddItem Trim(txtDescription.Text) & Space(65) & " ^ " & lsvAccHead.SelectedItem.SubItems(1)
'        Else
'            OriginalGroup = GetComboValue(cmbSubGroup.Text, DEFAULT_DELIMITER)
'
'            If Len(Trim(lsvAccHead.SelectedItem.SubItems(1))) > 0 Then
'                For i = 0 To cmbSubGroup.ListCount - 1
'                    cmbSubGroup.ListIndex = i
'                    If Trim(lsvAccHead.SelectedItem.SubItems(1)) = GetComboKey(cmbSubGroup.Text, "^") Then
'                        cmbSubGroup.RemoveItem (i)
'                        Exit For
'                    End If
'                Next
'            End If
'            cmbSubGroup.AddItem Trim(txtDescription.Text) & Space(65) & " ^ " & lsvAccHead.SelectedItem.SubItems(1)
'
'            If Len(Trim(OriginalGroup)) > 0 Then
'                For i = 0 To cmbSubGroup.ListCount - 1
'                    cmbSubGroup.ListIndex = i
'                    If UCase(Trim(OriginalGroup)) = UCase(GetComboValue(cmbSubGroup.Text, DEFAULT_DELIMITER)) Then
'                        Exit For
'                    End If
'                Next
'            End If
'        End If
    End If
    Set objInsertAcc = Nothing
End Sub
Private Sub UpdateList()
''Procedure to update list after insertion/modification of new record

    If bytButton = INSERT Then
        Dim objListItem As ListItem
        
        Set objListItem = lsvAccHead.ListItems.Add(, , txtAcCode.Text)
        With objListItem
            .SubItems(1) = txtDescription.Text
            .SubItems(2) = m_lngAccountKey
            .SubItems(3) = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER)
            .SubItems(4) = IIf(chkSubGroup.Value = 1, "Y", "N")
            .SubItems(5) = Val(numOpeningBal.Text)
            .SubItems(6) = m_strGroupNo
            .SubItems(7) = IIf(ChkPL.Value = 1, "Y", "N")
            .SubItems(8) = m_strParentno
            .SubItems(9) = IIf(ChkAgriLoan.Value = 1, "Y", "N")
            .SubItems(10) = txtMCL.Text
            .SubItems(11) = Format(Val(numKharifCash.Text))
            .SubItems(12) = Format(Val(numKharifKind.Text))
            .SubItems(13) = Format(Val(numRabiCash.Text))
            .SubItems(14) = Format(Val(numRabiKind.Text))
            .SubItems(15) = Format(Val(numMaxCash.Text))
            .SubItems(16) = Format(Val(numMaxCash.Text))
            SelectedIndex = objListItem.Index
            lsvAccHead_ItemClick lsvAccHead.ListItems(SelectedIndex)
        End With
    ElseIf bytButton = MODIFY Then
        With lsvAccHead.SelectedItem
            .Text = txtAcCode.Text
            .SubItems(1) = txtDescription.Text
            .SubItems(3) = GetComboKey(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER)
            .SubItems(4) = IIf(chkSubGroup.Value = 1, "Y", "N")
            .SubItems(5) = Val(numOpeningBal.Text)
            .SubItems(6) = m_strGroupNo
            .SubItems(7) = IIf(ChkPL.Value = 1, "Y", "N")
            .SubItems(8) = m_strParentno
            .SubItems(9) = IIf(ChkAgriLoan.Value = 1, "Y", "N")
            .SubItems(10) = txtMCL.Text
            .SubItems(11) = Format(Val(numKharifCash.Text))
            .SubItems(12) = Format(Val(numKharifKind.Text))
            .SubItems(13) = Format(Val(numRabiCash.Text))
            .SubItems(14) = Format(Val(numRabiKind.Text))
            .SubItems(15) = Format(Val(numMaxCash.Text))
            .SubItems(16) = Format(Val(numMaxCash.Text))
'            lsvAccHead_ItemClick lsvAccHead.ListItems(SelectedIndex)
        End With
    End If
End Sub
Private Function GetComboValue(ByVal v_ComboText As String, _
                               ByVal Delimiter As String) _
                               As String
''To get just visible text of combo box from multiple invisible columns of combo
    Dim charAtPosition As Byte
    If Len(v_ComboText) > 0 Then
        charAtPosition = InStr(1, v_ComboText, Delimiter)
        GetComboValue = Trim(VBA.Left(v_ComboText, charAtPosition - 1))
    End If
End Function

Private Sub lsvAccHead_KeyPress(KeyAscii As Integer)
    If lsvAccHead.ListItems.Count > 0 Then
        If KeyAscii = 13 Then cmdModify_Click
    End If
End Sub

Private Function GetSubGroups() As Boolean
    
    'This Function Will Retrieve Branches from the Database into the cboBranchPacs
    'Combo Box
    
    Dim objRecordset As ADODB.Recordset
    Dim strStoredProcedureText As String
    Dim blnIsListIndexSet As Boolean
    
    strStoredProcedureText = "{ Call PACK_ACHEADBranch_Data.GetGroupsAccHead()}"
    
    Set objRecordset = GetDataFromStoredProcedure(strStoredProcedureText)
                
    With objRecordset
        If .RecordCount > 0 Then
            While Not .EOF
                If Not IsNull(.Fields("ABC")) Then
                    cmbSubGroup.AddItem Trim(.Fields("ABC")) & Space(65) & " ^ " & .Fields("Gcode") & " ^ " & .Fields("Gkey")
'                    MsgBox Trim(.Fields("ABC")) & Space(65) & " ^ " & .Fields("Gcode") & " ^ " & .Fields("Gkey")
                End If
                .MoveNext
            Wend
            GetSubGroups = True
            Exit Function
        Else
            MsgBox "Cannot enter any Account Head as there are no Groups specified in Master", vbInformation, "KCCB"
            GetSubGroups = False
            Exit Function
        End If
    End With
    GetSubGroups = False
    Exit Function
    
End Function

Private Sub GetMaxGroupNo()
    ''Code to genarate New Group number for new account head
    
    Dim objRecordset As ADODB.Recordset
    Dim objCommand As New ADODB.Command
            
    With objCommand
        Set .ActiveConnection = g_objDataSource.GetDataSource
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("param1", adVarChar, adParamInput, 20, Mid(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), 1, InStr(1, GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) - 2))
        .CommandText = "Kccbank.Pack_ACHeadBranch_DATA.GetMaxGroupNoMst()"
        Set objRecordset = .Execute
    End With

    With objRecordset
        If .Fields("MaxNo") = 0 Then
            m_strGroupNo = Mid(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), 1, InStr(1, GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) - 2) & "01"
        Else
            m_strGroupNo = Mid(.Fields("MaxNo"), 1, Len(.Fields("MaxNo")) - 1) & Val(Mid(.Fields("MaxNo"), Len(.Fields("MaxNo")), 1)) + 1
        End If
    End With
    m_strParentno = Mid(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), 1, InStr(1, GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) - 2)
    Exit Sub
End Sub

Private Function GetDataFromStoredProcedure(ByVal strStoredProcedure As String) _
                                            As ADODB.Recordset
'To execute Procedure
    Dim objCommand As ADODB.Command
    Dim intCounter As Integer
    
    Set objCommand = New ADODB.Command
    
    With objCommand
        .CommandText = strStoredProcedure
        .CommandType = adCmdText
        
        Set .ActiveConnection = g_objDataSource.GetDataSource
        .CommandTimeout = 0
        
        Set GetDataFromStoredProcedure = .Execute
        Exit Function
    End With
    
End Function

Private Function CheckNoSubGroups() As Boolean
    ''Code to Check No. of subgroups attached under one parent
    ''which can be maximum of 8
    
    Dim objCommand As New ADODB.Command
            
    With objCommand
        Set .ActiveConnection = g_objDataSource.GetDataSource
        .CommandType = adCmdStoredProc
'        MsgBox cmbSubGroup.Text
'        MsgBox Mid(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), 1, InStr(1, GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) - 2)
        .Parameters.Append .CreateParameter("param1", adVarChar, adParamInput, 20, Mid(GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), 1, InStr(1, GetComboGroupCode(cmbSubGroup.Text, DEFAULT_DELIMITER), DEFAULT_DELIMITER) - 2))
        .Parameters.Append .CreateParameter("param2", adNumeric, adParamOutput)
        .CommandText = "Kccbank.Pack_ACHeadBranch_DATA.GetCountSubGroupsMst()"
        .Execute
    End With

    If IsNull(objCommand.Parameters("param2")) Then
    ElseIf objCommand.Parameters("param2") >= 8 Then
        MsgBox "Cannot add more subgroups under selected Group/Subgroup.", vbInformation, "KCCB"
        CheckNoSubGroups = False
        cmbSubGroup.SetFocus
        Exit Function
    End If

    CheckNoSubGroups = True
    Exit Function

End Function

Private Sub txtAcCode_KeyPress(KeyAscii As Integer)
    'To write account code always in Capitals
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
