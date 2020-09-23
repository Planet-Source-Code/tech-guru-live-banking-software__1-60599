VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChequeBookIssue 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7125
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTotalNoCheques 
      Height          =   615
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   7575
      Begin VB.Label lblTotalNoofCheques 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Total No of Cheques to be used"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraAccount 
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   7575
      Begin VB.TextBox txtAccountNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboAccountType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Account No"
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Type of Account"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   345
      Left            =   7080
      TabIndex        =   16
      Top             =   6360
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   6240
      TabIndex        =   15
      Top             =   6360
      Width           =   780
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   345
      Left            =   5400
      TabIndex        =   14
      Top             =   6360
      Width           =   780
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&E&xit"
      Height          =   345
      Left            =   2760
      TabIndex        =   13
      Top             =   6360
      Width           =   780
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   1920
      TabIndex        =   12
      Top             =   6360
      Width           =   780
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   345
      Left            =   1080
      TabIndex        =   11
      Top             =   6360
      Width           =   780
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   345
      Left            =   240
      TabIndex        =   10
      Top             =   6360
      Width           =   780
   End
   Begin VB.Frame fraChequeBook 
      Caption         =   "ChequeBooks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   7575
      Begin MSComctlLib.ListView lsvChequeBooks 
         Height          =   1875
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "SelectRecord tomodify/delete "
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date of Issue"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Loose"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cheque No From"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheque No To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Leaves"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cheque No Used"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Balance Cheques"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Charge Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ChequeBook Key"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame fradetails 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   7575
      Begin Establishment.NumberControl NumLeaves 
         Height          =   315
         Left            =   5400
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         MaxLength       =   4
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
      Begin Establishment.NumberControl NumChargeAmount 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Decimals        =   2
      End
      Begin Establishment.NumberControl NumChequeTo 
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         MaxLength       =   11
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
      Begin Establishment.NumberControl NumChequeFrom 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         MaxLength       =   11
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
      Begin VB.ComboBox cboLoose 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin MSMask.MaskEdBox medIssueDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "No of Leaves"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Charge Amount"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cheque No To"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cheque No From"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Date of Issue"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Loose"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5520
      TabIndex        =   37
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Balance"
      Height          =   255
      Left            =   4560
      TabIndex        =   36
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      TabIndex        =   35
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Name of A/c Holder"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KCCB -Head Office Thanesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   33
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label lblbottomright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "user anme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6480
      TabIndex        =   32
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChequeBook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   1350
   End
   Begin VB.Label lbltopright 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6360
      TabIndex        =   30
      Top             =   0
      Width           =   1320
   End
   Begin VB.Shape shpBottom 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   -120
      Top             =   6840
      Width           =   8715
   End
   Begin VB.Shape shpTop 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   8565
   End
End
Attribute VB_Name = "frmChequeBookIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_ObjRes            As New ResLoad.LoadRes
Dim m_blnInsertMode     As Boolean
Dim m_blnModifyMode     As Boolean
Dim m_blnDeleteMode     As Boolean
Dim LngGroupkey         As Long
Dim strprevCode         As String
Dim blnCodeExists       As Boolean
Dim blnValidation       As Boolean
Dim objDB As New KCCBChequeBookIssue.CChequeBookIssue

Private Sub cmdDelete_Click()
   Dim objDB As New KCCBChequeBookIssue.CChequeBookIssue
     Dim objChequebookList As New KCCBChequeBookIssue.CChequeBookIssue
     Dim objChequebook As New KCCBChequeBookIssue.CChequeBookIssue
     Dim objRecordset As ADODB.Recordset
     If lsvChequeBooks.ListItems.Count = 0 Then
            Exit Sub
     Else
     End If
     If Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(4).Text) <> _
       Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(6).Text) Then
            MsgBox "This Record cannot be Deleted,Cheques have been Issued", vbInformation + vbOKOnly, "Bank Error"
            Exit Sub
    End If
         m_blnDeleteMode = True
         Dim IntResponse As Integer
         IntResponse = MsgBox("Are You Sure To Delete The Selected Record ?", vbYesNo + vbQuestion + vbDefaultButton2, "Record Deletion")
                    If IntResponse = 6 Then
                            With objDB
                                    .ChequebookKey = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(8).Text)
                                    .SaveData g_objDataSource.GetDataSource, DBDelete
                            End With
                                    Call ClearControls
                                    lsvChequeBooks.ListItems.Clear
                            With objChequebook
                                    .TypeOfAccount = Trim(cboAccountType.Text)
                                    .AccountNo = Trim(txtAccountNo.Text)
                            End With
                                    Set objRecordset = objChequebook.GetChequeBookIssueList(g_objDataSource.GetDataSource)
                                    Call FillList(objRecordset)
                                    fradetails.Enabled = False
                                    fraAccount.Enabled = True
                                    EnableButtons
                                    Call DisplayValues
                                    SetTotalUnusedCheques
                     Else
                                    Exit Sub
                     End If
                                 m_blnDeleteMode = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Activate()
cmdCancel_Click
End Sub

Private Sub Form_Load()
    If CheckResourceDLL Then
                m_ObjRes.DllName = App.Path & "\Dlls\Resources\kccbres.dll"
        Else
                Unload Me
            End
    End If
            lblbottomright.Caption = Trim(strUserName)
            lbltopright.Caption = strTodaysDate
            Call FillAccountCombo
            Call FillLooseCombo
            fradetails.Enabled = False
            lblTotalNoofCheques.Caption = ""
            Label9.Caption = ""
            Label11.Caption = ""
            cmdOk.Enabled = False
            cmdCancel.Enabled = False
               m_blnInsertMode = False
                m_blnModifyMode = False
                m_blnDeleteMode = False
'    Set objOccupation = Nothing
End Sub

Private Sub FillAccountCombo()
        cboAccountType.AddItem ("SB")
        cboAccountType.AddItem ("CA")
        cboAccountType.AddItem ("CC")
        cboAccountType.AddItem ("FD")
        cboAccountType.AddItem ("RD")
End Sub

Private Sub FillLooseCombo()
        cboLoose.AddItem ("Y")
        cboLoose.AddItem ("N")
End Sub

Private Sub cmdInsert_Click()
        fraAccount.Enabled = False
        fradetails.Enabled = True
        Call DisableButtons
        Call DetailsEnable
        Call ClearControls
        m_blnInsertMode = True
        medIssueDate.SetFocus
End Sub

Private Sub EnableButtons()
    cmdInsert.Enabled = True
    cmdModify.Enabled = True
    cmdDelete.Enabled = True
    cmdExit.Enabled = True
    cmdOk.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub DisableButtons()
    cmdInsert.Enabled = False
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    cmdOk.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub ClearControls()
    medIssueDate.Mask = ""
    medIssueDate.Text = ""
    medIssueDate.Mask = "##/##/####"
    cboLoose.ListIndex = -1
    NumChequeFrom.Text = ""
    NumChequeTo.Text = ""
    NumChargeAmount.Text = ""
    NumLeaves.Text = ""
End Sub

Private Sub DetailsDisable()
    fradetails.Enabled = False
End Sub
Private Sub DetailsEnable()
    fradetails.Enabled = True
End Sub

Private Sub FillList(ByRef objRecordset As ADODB.Recordset)
   Dim objListItem As ListItem
        
        While Not objRecordset.EOF
            Set objListItem = lsvChequeBooks.ListItems.Add
            With objListItem
                .Text = objRecordset("Issue_Date")
                .SubItems(1) = objRecordset("Loose_Cheque")
                .SubItems(2) = objRecordset("Cheque_No_From")
                .SubItems(3) = objRecordset("Cheque_No_To")
                .SubItems(4) = (CLng(.SubItems(3)) - CLng(.SubItems(2))) + 1
                     If Not IsNull(objRecordset.Fields("NO_OF_CHEQUES_USED")) Then
                            .SubItems(5) = Trim(objRecordset.Fields("NO_OF_CHEQUES_USED"))
                    Else
                            .SubItems(5) = 0
                    End If
                .SubItems(6) = CLng(.SubItems(4)) - CLng(.SubItems(5))
                .SubItems(7) = objRecordset("Money_Charged")
                .SubItems(8) = objRecordset("Chequebook_Key")
            End With
            objRecordset.MoveNext
            Set objListItem = Nothing
        Wend
        Set objRecordset = Nothing
End Sub
Private Sub lsvChequeBooks_KeyDown(KeyCode As Integer, Shift As Integer)
            Dim bytIndex    As Byte
        If Not lsvChequeBooks.ListItems.Count = 0 Then
                If KeyCode = vbKeyUp Then
                                'Select the previous row of the visible list only,
                                'The same corresponding row will be selected when
                                'procedure to display values will be called.
                                'Now Check the visibile property of the list.
                        If Not lsvChequeBooks.SelectedItem.Index = 1 Then
                                lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index - 1).Selected = True
                                'Scroll the list view
                                lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).EnsureVisible '= True
                        End If
                                'Call procedure to display values
                                 Call DetailsEnable
                                 Call DisplayValues
                                 Call DetailsDisable
                                 KeyCode = 0
                ElseIf KeyCode = vbKeyDown Then

                        If Not lsvChequeBooks.SelectedItem.Index = lsvChequeBooks.ListItems.Count Then
                                lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index + 1).Selected = True
                                'Scroll the list view
                                lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).EnsureVisible
                        End If
                                'Call procedure to display values
                                Call DetailsEnable
                                Call DisplayValues
                                Call DetailsDisable
                                KeyCode = 0
                End If
        End If
            Screen.MousePointer = vbDefault
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
     If cboAccountType.Text <> "" And txtAccountNo.Text <> "" Then
            If KeyAscii = vbKeyReturn Then
                    Dim objChequebook As New KCCBChequeBookIssue.CChequeBookIssue
            Dim objRecordset As ADODB.Recordset
            
         With objChequebook
                   .TypeOfAccount = Trim(cboAccountType.Text)
                   .AccountNo = Trim(txtAccountNo.Text)
                    
         End With
                    Set objRecordset = objChequebook.GetChequeBookIssueList(g_objDataSource.GetDataSource)
                    

                    If objRecordset.RecordCount = 0 Then
                            MsgBox "No Data Found", vbInformation + vbOKOnly, "Finding Data"
                            Call ClearControls
                            lsvChequeBooks.ListItems.Clear
                            fradetails.Enabled = False
                            Call EnableButtons
                            lblTotalNoofCheques.Caption = ""
                            Label9.Caption = ""
                            Label11.Caption = ""
                            Exit Sub
                    Else
                    
                           lsvChequeBooks.ListItems.Clear
                           Call FillList(objRecordset)
                           fradetails.Enabled = False
                           fraAccount.Enabled = True
                           EnableButtons
                           Call DisplayValues
                           Call SetTotalUnusedCheques
                           Call SetNameAndBalance
                  End If
        End If
End If

End Sub

Private Sub txtAccountNo_LostFocus()
    If cboAccountType.Text <> "" And txtAccountNo.Text <> "" Then
            Dim objChequebook As New KCCBChequeBookIssue.CChequeBookIssue
            Dim objRecordset As ADODB.Recordset
            
         With objChequebook
                   .TypeOfAccount = Trim(cboAccountType.Text)
                   .AccountNo = Trim(txtAccountNo.Text)
                    
         End With
                    Set objRecordset = objChequebook.GetChequeBookIssueList(g_objDataSource.GetDataSource)
                    

                    If objRecordset.RecordCount = 0 Then
                            MsgBox "No Data Found", vbInformation + vbOKOnly, "Finding Data"
                            Call ClearControls
                            lsvChequeBooks.ListItems.Clear
                            fradetails.Enabled = False
                            Call EnableButtons
                            lblTotalNoofCheques.Caption = ""
                            Label9.Caption = ""
                            Label11.Caption = ""
                            Exit Sub
                    Else
                    
                           lsvChequeBooks.ListItems.Clear
                           Call FillList(objRecordset)
                           fradetails.Enabled = False
                           fraAccount.Enabled = True
                           EnableButtons
                           Call DisplayValues
                           Call SetTotalUnusedCheques
                           Call SetNameAndBalance
                  End If
        End If
End Sub

Private Sub cmdOK_Click()
      Dim objChequebook As New KCCBChequeBookIssue.CChequeBookIssue
     Dim objRecordset As ADODB.Recordset
                   Call Validate
        If blnValidation = False Then
                Exit Sub
        End If
        If m_blnInsertMode And blnValidation Then
           With objDB
                .TypeOfAccount = Trim(cboAccountType.Text)
                .AccountNo = Trim(txtAccountNo.Text)
                .ChequeNoFrom = Trim(NumChequeFrom.Text)
                .ChequeNoTo = Trim(NumChequeTo.Text)
                .LooseCheque = Trim(cboLoose.Text)
                .MoneyCharged = Trim(NumChargeAmount.Text)
                .IssueDate = CDate(medIssueDate.Text)
                .TerminalName = Trim(strComputerName)
                .UserName = Trim(strUserName)
                .InsertDate = CDate(dteTodaysDate)
                .SaveData g_objDataSource.GetDataSource, DBInsert
           End With
                Call ClearControls
                lsvChequeBooks.ListItems.Clear
            With objChequebook
                   .TypeOfAccount = Trim(cboAccountType.Text)
                   .AccountNo = Trim(txtAccountNo.Text)
            End With
                Set objRecordset = objChequebook.GetChequeBookIssueList(g_objDataSource.GetDataSource)
                Call FillList(objRecordset)
                fradetails.Enabled = False
                fraAccount.Enabled = True
                EnableButtons
                Call DisplayValues
                SetTotalUnusedCheques
       ElseIf m_blnModifyMode And blnValidation Then
          With objDB
                .ChequebookKey = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(8).Text)
                .TypeOfAccount = Trim(cboAccountType.Text)
                .AccountNo = Trim(txtAccountNo.Text)
                .ChequeNoFrom = Trim(NumChequeFrom.Text)
                .ChequeNoTo = Trim(NumChequeTo.Text)
                .LooseCheque = Trim(cboLoose.Text)
                .MoneyCharged = Trim(NumChargeAmount.Text)
                .IssueDate = CDate(medIssueDate.Text)
                .TerminalName = Trim(strComputerName)
                .UserName = Trim(strUserName)
                .InsertDate = CDate(dteTodaysDate)
                .SaveData g_objDataSource.GetDataSource, DBModify
         End With
                Call ClearControls
                lsvChequeBooks.ListItems.Clear
            With objChequebook
                   .TypeOfAccount = Trim(cboAccountType.Text)
                   .AccountNo = Trim(txtAccountNo.Text)
            End With
                Set objRecordset = objChequebook.GetChequeBookIssueList(g_objDataSource.GetDataSource)
                Call FillList(objRecordset)
                fradetails.Enabled = False
                fraAccount.Enabled = True
                EnableButtons
                Call DisplayValues
                SetTotalUnusedCheques
         End If
                    m_blnInsertMode = False
                    m_blnModifyMode = False
                    m_blnDeleteMode = False
                    Set objDB = Nothing
End Sub

Private Sub DisplayValues()
    If Not lsvChequeBooks.ListItems.Count = 0 Then
          NumChequeTo.Text = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(3).Text)
          NumChequeFrom.Text = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(2).Text)
          NumChargeAmount.Text = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(7).Text)
          NumLeaves.Text = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(4).Text)
          medIssueDate.Text = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).Text)
          cboLoose.Text = Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(1).Text)
    Else
        End If
End Sub

Private Sub Validate()
    blnValidation = True
    With m_ObjRes
       If Not IsDate(medIssueDate.FormattedText) Then
                .AddDynamicParameter = "Date of Issue"
                MsgBox .ReplaceStringByParam(m_ObjRes.LoadStringFromDLL(1011)), vbInformation, "Bank Error"
                medIssueDate.SetFocus
                blnValidation = False
                Exit Sub
       ElseIf Trim(cboLoose.Text) = Empty Then
                .AddDynamicParameter = "Loose"
                MsgBox .ReplaceStringByParam(m_ObjRes.LoadStringFromDLL(1003)), vbInformation, "Bank Error"
                cboLoose.SetFocus
                blnValidation = False
                Exit Sub
       ElseIf Trim(NumChequeFrom.Text) = Empty Then
                .AddDynamicParameter = "Cheque No From"
                MsgBox .ReplaceStringByParam(m_ObjRes.LoadStringFromDLL(1003)), vbInformation, "Bank Error"
                NumChequeFrom.SetFocus
                blnValidation = False
                Exit Sub
       ElseIf Trim(NumChequeTo.Text) = Empty Then
                .AddDynamicParameter = "Cheque No To"
                MsgBox .ReplaceStringByParam(m_ObjRes.LoadStringFromDLL(1003)), vbInformation, "Bank Error"
                NumChequeTo.SetFocus
                blnValidation = False
                Exit Sub

        ElseIf Trim(NumLeaves.Text) <> (CDbl(Trim(NumChequeTo.Text)) - CDbl(Trim(NumChequeFrom.Text))) And (Trim(NumChequeTo.Text) <> Trim(NumChequeFrom.Text)) Then
                MsgBox "No of Leaves is not Matching with your cheques series", vbInformation + vbOKOnly, "Bank Error"
                NumLeaves.SetFocus
                blnValidation = False
                Exit Sub

       Else
                blnValidation = True
       End If
   End With

        If Trim(NumChargeAmount.Text) = "" Then
                NumChargeAmount.Text = 0
        Else
                NumChargeAmount.Text = Trim(NumChargeAmount.Text)
        End If
End Sub
Private Sub lsvChequeBooks_Click()
    If lsvChequeBooks.ListItems.Count <> 0 Then
            DisplayValues
    Else
            Exit Sub
    End If
End Sub
Private Sub lsvChequeBooks_DblClick()
     If lsvChequeBooks.ListItems.Count <> 0 Then
            lsvChequeBooks_Click
    Else
            Exit Sub
    End If
End Sub

Private Sub cmdModify_Click()
     If lsvChequeBooks.ListItems.Count = 0 Then
            Exit Sub
     Else
     End If
    If Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(4).Text) <> _
       Trim(lsvChequeBooks.ListItems(lsvChequeBooks.SelectedItem.Index).ListSubItems(6).Text) Then
            MsgBox "This Record cannot be Modified,Cheques have been Issued", vbInformation + vbOKOnly, "Bank Error"
            Exit Sub
    End If
        fraAccount.Enabled = False
        fradetails.Enabled = True
        Call DisableButtons
        Call DetailsEnable
         m_blnModifyMode = True
        medIssueDate.SetFocus
End Sub

Private Sub SetTotalUnusedCheques()
  Dim i As Integer
  Dim dblbalance As Double
        If lsvChequeBooks.ListItems.Count <> 0 Then
            dblbalance = Trim(lsvChequeBooks.ListItems(1).ListSubItems(6).Text)
                For i = 2 To lsvChequeBooks.ListItems.Count
                    dblbalance = dblbalance + CDbl(Trim(lsvChequeBooks.ListItems(i).ListSubItems(6).Text))
                Next i
        Else
            Exit Sub
        End If
            lblTotalNoofCheques.Caption = dblbalance
End Sub

Private Sub SetNameAndBalance()
           Dim objChequebook As New KCCBChequeBookIssue.CChequeBookIssue
           Dim dblbalance As Double
           Dim objRecordset As ADODB.Recordset
            
         With objChequebook
                   .TypeOfAccount = Trim(cboAccountType.Text)
                   .AccountNo = Trim(txtAccountNo.Text)
         End With
                    Set objRecordset = objChequebook.GetNameAndBalance(g_objDataSource.GetDataSource)
                             If objRecordset.RecordCount <> 0 Then
                                    Label9.Caption = objRecordset.Fields("NAME_OF_AC_HOLDER")
                                    dblbalance = CDbl(objRecordset.Fields("Opening_Balance")) + (CDbl(objRecordset.Fields("Tot_dr")) - CDbl(objRecordset.Fields("Tot_Cr")))
                                    Label11.Caption = dblbalance
                             Else
                                    Exit Sub
                             End If
            Set objRecordset = Nothing
End Sub

Private Sub cmdCancel_Click()
                lsvChequeBooks.ListItems.Clear
                Call DetailsEnable
                Call DetailsDisable
                Call EnableButtons
                Call ClearControls
                cboAccountType.ListIndex = -1
                txtAccountNo.Text = ""
                fraAccount.Enabled = True
                lblTotalNoofCheques.Caption = ""
                Label9.Caption = ""
                Label11.Caption = ""
                
                m_blnInsertMode = False
                m_blnModifyMode = False
                m_blnDeleteMode = False
End Sub
