VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAccountOpening 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   945
   ClientTop       =   405
   ClientWidth     =   10290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab ssTabAccOpening 
      Height          =   6945
      Left            =   -30
      TabIndex        =   46
      Top             =   270
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   12250
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Account Opening"
      TabPicture(0)   =   "frmAccountOpening.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTop"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAllFields"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Joint Detail"
      TabPicture(1)   =   "frmAccountOpening.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraButtonJoint"
      Tab(1).Control(1)=   "fraFieldsJoint"
      Tab(1).Control(2)=   "fraLsvDetail"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraButtonJoint 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   -69960
         TabIndex        =   84
         Top             =   2730
         Width           =   4365
         Begin VB.CommandButton cmdAddToList 
            Caption         =   "<<"
            Height          =   345
            Left            =   150
            TabIndex        =   86
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdInsertJoint 
            Caption         =   "+"
            Height          =   345
            Left            =   810
            TabIndex        =   87
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdModifyJoint 
            Caption         =   "$"
            Height          =   345
            Left            =   1500
            TabIndex        =   88
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdDeletejoint 
            Caption         =   "X"
            Height          =   345
            Left            =   2160
            TabIndex        =   89
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdOKJoint 
            Caption         =   "&OK"
            Height          =   345
            Left            =   3540
            TabIndex        =   91
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdCancelJoint 
            Caption         =   "Cancel"
            Height          =   345
            Left            =   2850
            TabIndex        =   90
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraFieldsJoint 
         Height          =   2025
         Left            =   -69810
         TabIndex        =   79
         Top             =   570
         Width           =   4035
         Begin VB.TextBox txtJointName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1110
            MaxLength       =   100
            TabIndex        =   80
            Top             =   420
            Width           =   2655
         End
         Begin VB.Frame fraOperate 
            Caption         =   "Operations"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   960
            TabIndex        =   82
            Top             =   930
            Width           =   2955
            Begin VB.OptionButton optMust 
               Caption         =   "Must-Operate"
               Height          =   195
               Left            =   1500
               TabIndex        =   85
               Top             =   300
               Width           =   1275
            End
            Begin VB.OptionButton optCan 
               Caption         =   "Can-Operate"
               Height          =   225
               Left            =   120
               TabIndex        =   83
               Top             =   270
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Label lblJN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Joint Name"
            Height          =   195
            Left            =   180
            TabIndex        =   81
            Top             =   450
            Width           =   795
         End
      End
      Begin VB.Frame fraLsvDetail 
         Caption         =   "Joint Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   -74820
         TabIndex        =   77
         Top             =   540
         Width           =   4185
         Begin MSComctlLib.ListView lsvJointDetail 
            Height          =   2745
            Left            =   30
            TabIndex        =   78
            Top             =   210
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   4842
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
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Joint Key"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Joint Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Operate"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Staus"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame fraAllFields 
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   10935
         Begin Branch.NumberControl numLedgerNo 
            Height          =   375
            Left            =   7440
            TabIndex        =   94
            Top             =   4800
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
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
         Begin VB.TextBox txtRemarks 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   870
            MaxLength       =   200
            TabIndex        =   39
            Top             =   5370
            Width           =   6765
         End
         Begin VB.Frame fraClosed 
            Caption         =   "Closed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   73
            Top             =   4710
            Width           =   6195
            Begin VB.CheckBox chkClosed 
               Caption         =   "A/C Closed/Ceased/Suspended"
               Height          =   345
               Left            =   150
               TabIndex        =   37
               Top             =   210
               Width           =   2610
            End
            Begin VB.TextBox txtClosedReason 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3780
               MaxLength       =   30
               TabIndex        =   38
               Top             =   240
               Width           =   2325
            End
            Begin VB.Label lblReasonClosed 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reason"
               Height          =   195
               Left            =   3180
               TabIndex        =   74
               Top             =   300
               Width           =   555
            End
         End
         Begin VB.Frame fraAccHolder 
            Caption         =   "A/C Holder"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   30
            TabIndex        =   65
            Top             =   -30
            Width           =   10155
            Begin VB.TextBox txtPanNo 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7650
               MaxLength       =   30
               TabIndex        =   12
               Top             =   660
               Width           =   2415
            End
            Begin VB.ComboBox cboOccupation 
               Height          =   315
               Left            =   7650
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   180
               Width           =   2445
            End
            Begin VB.Frame fraAdultMinor 
               Height          =   495
               Left            =   480
               TabIndex        =   13
               Top             =   990
               Width           =   1700
               Begin VB.OptionButton optMinor 
                  Caption         =   "Minor"
                  Height          =   315
                  Left            =   150
                  TabIndex        =   15
                  Top             =   120
                  Width           =   735
               End
               Begin VB.OptionButton optAdult 
                  Caption         =   "Adult"
                  Height          =   315
                  Left            =   930
                  TabIndex        =   14
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VB.TextBox txtAddress 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4050
               MaxLength       =   200
               TabIndex        =   16
               Top             =   1140
               Width           =   6045
            End
            Begin VB.TextBox txtPhNo 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4050
               MaxLength       =   20
               TabIndex        =   11
               Top             =   660
               Width           =   1815
            End
            Begin VB.TextBox txtFName 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4050
               MaxLength       =   100
               TabIndex        =   8
               Top             =   180
               Width           =   2655
            End
            Begin VB.TextBox txtName 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   510
               MaxLength       =   100
               TabIndex        =   7
               Top             =   210
               Width           =   2475
            End
            Begin MSMask.MaskEdBox medDob 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Left            =   510
               TabIndex        =   10
               Top             =   630
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblPan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pan No."
               Height          =   195
               Left            =   6870
               TabIndex        =   72
               Top             =   750
               Width           =   585
            End
            Begin VB.Label lblOcc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Occupation"
               Height          =   195
               Left            =   6810
               TabIndex        =   71
               Top             =   210
               Width           =   825
            End
            Begin VB.Label lblAddress 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   195
               Left            =   3270
               TabIndex        =   70
               Top             =   1170
               Width           =   570
            End
            Begin VB.Label lblPhno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ph. No."
               Height          =   195
               Left            =   3270
               TabIndex        =   69
               Top             =   690
               Width           =   540
            End
            Begin VB.Label lblDOB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dob"
               Height          =   195
               Left            =   60
               TabIndex        =   68
               Top             =   660
               Width           =   300
            End
            Begin VB.Label lblFname 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "F' Name"
               Height          =   195
               Left            =   3270
               TabIndex        =   67
               Top             =   240
               Width           =   585
            End
            Begin VB.Label lblName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               Height          =   195
               Left            =   60
               TabIndex        =   66
               Top             =   270
               Width           =   420
            End
         End
         Begin VB.Frame fraNominee 
            Caption         =   "Nominee"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   0
            TabIndex        =   60
            Top             =   2190
            Width           =   10125
            Begin VB.TextBox txtAddressNom 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   660
               MaxLength       =   200
               TabIndex        =   23
               Top             =   630
               Width           =   5985
            End
            Begin VB.TextBox txtRelationNom 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7380
               MaxLength       =   30
               TabIndex        =   22
               Top             =   210
               Width           =   2655
            End
            Begin VB.TextBox txtNameNom 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   660
               MaxLength       =   100
               TabIndex        =   20
               Top             =   210
               Width           =   2655
            End
            Begin MSMask.MaskEdBox medDOBNom 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Left            =   3990
               TabIndex        =   21
               Top             =   210
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblAddressNom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   195
               Left            =   60
               TabIndex        =   64
               Top             =   660
               Width           =   570
            End
            Begin VB.Label lblrelationNom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Relation With A/C Holder"
               Height          =   195
               Left            =   5400
               TabIndex        =   63
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label lblDOBNom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dob"
               Height          =   195
               Left            =   3600
               TabIndex        =   62
               Top             =   270
               Width           =   300
            End
            Begin VB.Label lblNamenom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               Height          =   195
               Left            =   180
               TabIndex        =   61
               Top             =   270
               Width           =   420
            End
         End
         Begin VB.Frame fraAccount 
            Caption         =   "Operations - ROI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   0
            TabIndex        =   33
            Top             =   3300
            Width           =   10185
            Begin Branch.NumberControl numRateOfInterestD 
               Height          =   375
               Left            =   9480
               TabIndex        =   95
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   661
            End
            Begin Branch.NumberControl numRateOfInterestC 
               Height          =   375
               Left            =   7080
               TabIndex        =   93
               Top             =   240
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   661
            End
            Begin VB.Frame fraSevJoint 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1680
               TabIndex        =   27
               Top             =   180
               Width           =   1785
               Begin VB.OptionButton optJointOperate 
                  Caption         =   "Joint"
                  Height          =   315
                  Left            =   1050
                  TabIndex        =   29
                  Top             =   150
                  Width           =   645
               End
               Begin VB.OptionButton optSeverallyOperate 
                  Caption         =   "Survival"
                  Height          =   315
                  Left            =   30
                  TabIndex        =   28
                  Top             =   150
                  Width           =   975
               End
            End
            Begin VB.Frame fraSJ 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   60
               TabIndex        =   24
               Top             =   180
               Width           =   1575
               Begin VB.OptionButton optSingle 
                  Caption         =   "Single"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   25
                  Top             =   150
                  Width           =   735
               End
               Begin VB.OptionButton optJoint 
                  Caption         =   "Joint"
                  Height          =   315
                  Left            =   810
                  TabIndex        =   26
                  Top             =   150
                  Width           =   735
               End
            End
            Begin VB.Frame fraSP 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3510
               TabIndex        =   30
               Top             =   180
               Width           =   1725
               Begin VB.OptionButton optStaff 
                  Caption         =   "Staff"
                  Height          =   315
                  Left            =   90
                  TabIndex        =   31
                  Top             =   150
                  Width           =   735
               End
               Begin VB.OptionButton optPublic 
                  Caption         =   "Public"
                  Height          =   315
                  Left            =   900
                  TabIndex        =   32
                  Top             =   150
                  Width           =   735
               End
            End
            Begin VB.Label lblROfID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate Of Interest - DR"
               Height          =   195
               Left            =   7830
               TabIndex        =   59
               Top             =   330
               Width           =   1500
            End
            Begin VB.Label lblROfIC 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rate on Interest - CR"
               Height          =   195
               Left            =   5430
               TabIndex        =   58
               Top             =   360
               Width           =   1500
            End
         End
         Begin VB.Frame fraGuardian 
            Caption         =   "Guardian"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            TabIndex        =   54
            Top             =   1560
            Width           =   10125
            Begin VB.TextBox txtNameGur 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   630
               MaxLength       =   100
               TabIndex        =   17
               Top             =   210
               Width           =   2655
            End
            Begin VB.TextBox txtRelationGur 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7350
               MaxLength       =   30
               TabIndex        =   19
               Top             =   210
               Width           =   2655
            End
            Begin MSMask.MaskEdBox medDOBGur 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Left            =   3960
               TabIndex        =   18
               Top             =   210
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblNameGur 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               Height          =   195
               Left            =   180
               TabIndex        =   57
               Top             =   270
               Width           =   420
            End
            Begin VB.Label lblDOBGur 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dob"
               Height          =   195
               Left            =   3600
               TabIndex        =   56
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lblRelation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Relation With A/C Holder"
               Height          =   195
               Left            =   5370
               TabIndex        =   55
               Top             =   240
               Width           =   1800
            End
         End
         Begin VB.Frame fraIntroduction 
            Caption         =   "Introduction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   51
            Top             =   4080
            Width           =   10125
            Begin VB.TextBox txtIntroducedBy 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1230
               MaxLength       =   9
               TabIndex        =   34
               Top             =   210
               Width           =   2475
            End
            Begin VB.TextBox txtReasonIntro 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7380
               MaxLength       =   30
               TabIndex        =   36
               Top             =   210
               Width           =   2655
            End
            Begin VB.CheckBox chkCanIntoduce 
               Caption         =   "Can Introduce Others "
               Height          =   345
               Left            =   3930
               TabIndex        =   35
               Top             =   180
               Value           =   1  'Checked
               Width           =   1365
            End
            Begin VB.Label lblReasonintro 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reason, If can not Inroduce"
               Height          =   195
               Left            =   5340
               TabIndex        =   53
               Top             =   240
               Width           =   1995
            End
            Begin VB.Label lblIntoBy 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Introduced By"
               Height          =   195
               Left            =   120
               TabIndex        =   52
               Top             =   270
               Width           =   990
            End
         End
         Begin VB.Label lblLedgerNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LedgerNo"
            Height          =   195
            Left            =   6630
            TabIndex        =   76
            Top             =   4890
            Width           =   705
         End
         Begin VB.Label lblRemarks 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   195
            Left            =   150
            TabIndex        =   75
            Top             =   5430
            Width           =   630
         End
      End
      Begin VB.Frame fraTop 
         Height          =   675
         Left            =   90
         TabIndex        =   47
         Top             =   360
         Width           =   10155
         Begin Branch.NumberControl numAccountNo 
            Height          =   375
            Left            =   5760
            TabIndex        =   92
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
         End
         Begin VB.ComboBox cboAccountType 
            Height          =   315
            Left            =   2340
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label lblTAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type of A/C"
            Height          =   195
            Left            =   1320
            TabIndex        =   49
            Top             =   210
            Width           =   870
         End
         Begin VB.Label lblAccNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
            Height          =   195
            Left            =   4830
            TabIndex        =   48
            Top             =   210
            Width           =   900
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   9420
      TabIndex        =   45
      ToolTipText     =   "Click to get Help"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8610
      TabIndex        =   41
      ToolTipText     =   "Click to Cancel Record"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   7800
      TabIndex        =   40
      ToolTipText     =   "Click to Save Record"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   360
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "Click to Modify Selected Record"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   360
      Left            =   1650
      TabIndex        =   5
      ToolTipText     =   "Click to Delete Selected Record"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdLookUp 
      Caption         =   "&Look Up"
      Height          =   360
      Left            =   3780
      TabIndex        =   2
      ToolTipText     =   "Click to Select the Entered Record"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   360
      Left            =   2460
      TabIndex        =   6
      ToolTipText     =   "Click to Close the Form"
      Top             =   7320
      Width           =   780
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   360
      Left            =   30
      TabIndex        =   3
      ToolTipText     =   "Click to Insert New Record"
      Top             =   7320
      Width           =   780
   End
   Begin VB.Label lblUserName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jagtar Singh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   8880
      TabIndex        =   44
      Top             =   7740
      Width           =   1320
   End
   Begin VB.Label lblBranch 
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
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   0
      TabIndex        =   43
      Top             =   7740
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Top             =   7710
      Width           =   11415
   End
   Begin VB.Label lblFormName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Opening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   75
      TabIndex        =   42
      Top             =   30
      Width           =   1770
   End
   Begin VB.Label lblDateTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jagtar Singh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   8880
      TabIndex        =   0
      Top             =   30
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   -60
      Top             =   0
      Width           =   11505
   End
End
Attribute VB_Name = "frmAccountOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------
'       PROJECT    :KCCB
'       MODULE     :BRANCH
'       FORM       :ACCOUNT OPENING MASTER
'       OBJECTIVE  :FOR BRANCH, MASTER FORM
'       MADE BY    :Jagtar Singh
'       MADE DATE  :21-03-2002
'       MODIFY BY  :
'       MODIFY DATE :
'       REASON OF MODIFICATION :
'       COPY RIGHT @ 2001-2002 SURYA INFONET LTD.
'------------------------------------------------------------------

Private objDB As New clsAccountOpening
Private bytButton As Byte 'Main buttons state
Private bytJoint As Byte
Private JointState As Byte
Private bytjointButton As Byte 'Joint tab buttons state
Private bytFillAgain As Byte
Private strTypeofAccount As String
Private lngAccountNumber As Long
Private Sub FormState()
    'Initial state of form
    ssTabAccOpening.Tab = 0
    bytButton = 0
    bytjointButton = 0
    fraTop.Enabled = True
    fraAllFields.Enabled = False
    fraFieldsJoint.Enabled = False
    fraButtonJoint.Enabled = False
    cmdLookUp.Enabled = True
    cmdInsert.Enabled = True
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
    cmdExit.Enabled = True
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    cmdInsertJoint.Enabled = True
    cmdModifyJoint.Enabled = True
    cmdDeletejoint.Enabled = True
    cmdAddToList.Enabled = False
    cmdOKJoint.Enabled = True
    cmdCancelJoint.Enabled = True
    txtClosedReason.Enabled = False
End Sub
Private Sub GetSinleAccount()
    Dim objRecordset As ADODB.Recordset
    Dim objRecordset1 As ADODB.Recordset
 With objDB
        If bytFillAgain = 0 Then
            'if bytfillaain is 0 the send value again
            .TypeOfAccount = Trim(Left(cboAccountType.Text, 2))
            .AccountNo = CLng(numAccountNo.Text)
        End If
     'Get Single record
     Set objRecordset = .GetACCOUNTDETAIL(g_objDataSource.GetDataSource)
     If objRecordset.RecordCount = 0 Then
        MsgBox "Record Not Found"
        Exit Sub
     End If
     'If record found Display it
    Call DisplayDetail(objRecordset)
    lsvJointDetail.ListItems.Clear
    txtJointName.Text = Empty
    optCan.Value = False
    optMust.Value = False
    If optJoint.Value = True Then
        'if account is joint the get joint data
        Set objRecordset1 = .GETJOINTDATA(g_objDataSource.GetDataSource)
        'Fill values in listview
        Call FillListView(objRecordset1)
        'Display selected record in fields
        Call DisplayValues
    End If
 End With
    cmdModify.Enabled = True
    cmdDelete.Enabled = True
End Sub
Private Sub cboAccountType_KeyPress(KeyAscii As Integer)
    If cboAccountType.Text <> Empty And numAccountNo.Text <> Empty Then
        If KeyAscii = vbKeyReturn Then
            'If Both fields are filled and key is Enter then get record
            Call GetSinleAccount
            bytFillAgain = 0
        End If
    End If
End Sub
Private Sub chkCanIntoduce_Click()
    txtReasonIntro.Text = Empty
    If chkCanIntoduce.Value = Checked Then
        'if checked then set its enabled false and set text empty
        txtReasonIntro.Enabled = False
        txtReasonIntro.Text = Empty
    Else
        txtReasonIntro.Enabled = True
    End If
End Sub
Private Sub chkClosed_Click()
    txtClosedReason.Text = Empty
    If chkClosed.Value = Checked Then
        txtClosedReason.Enabled = True
    Else
        'if checked then set its enabled false and set text empty
        txtClosedReason.Enabled = False
        txtClosedReason.Text = Empty
    End If
End Sub
Private Sub cmdAddToList_Click()
'Insert values into listview
    Dim objListItem As ListItem
    If txtJointName.Text = Empty Then
        MsgBox "Enter Joint Name"
        txtJointName.SetFocus
        Exit Sub
    ElseIf optCan.Value = False And optMust.Value = False Then
        MsgBox "Select Minimum one option"
        optCan.SetFocus
        Exit Sub
    End If
    If bytjointButton = INSERT Then
        If lsvJointDetail.ListItems.Count = 0 Then
            Set objListItem = lsvJointDetail.ListItems.Add(, , 1)
        Else
            Set objListItem = lsvJointDetail.ListItems.Add(, , lsvJointDetail.ListItems(lsvJointDetail.ListItems.Count).Text + 1)
        End If
        objListItem.SubItems(1) = Trim(txtJointName.Text)
        If optCan.Value = True Then
            objListItem.SubItems(2) = "C"
        Else
            objListItem.SubItems(2) = "M"
        End If
        objListItem.SubItems(3) = "i"
        txtJointName.SetFocus
    ElseIf bytjointButton = MODIFY Then
    'if button state is modify then modify the seleced record
        lsvJointDetail.SelectedItem.SubItems(1) = Trim(txtJointName.Text)
        If optCan.Value = True Then
            lsvJointDetail.SelectedItem.SubItems(2) = "C"
        Else
            lsvJointDetail.SelectedItem.SubItems(2) = "D"
        End If
        If lsvJointDetail.SelectedItem.SubItems(3) = "d" Then
            lsvJointDetail.SelectedItem.SubItems(3) = "m"
        End If
     End If
    If bytjointButton = INSERT Then
        txtJointName.Text = Empty
        optCan.Value = True
    ElseIf bytjointButton = MODIFY Then
        fraLsvDetail.Enabled = True
        fraFieldsJoint.Enabled = False
        cmdInsertJoint.Enabled = True
        cmdModifyJoint.Enabled = True
        cmdDeletejoint.Enabled = True
    End If
End Sub
Private Sub cmdCancel_Click()
    FormState
    Call ClearControls(Me)
End Sub

Private Sub cmdCancelJoint_Click()
    fraFieldsJoint.Enabled = False
    fraLsvDetail.Enabled = True
    cmdInsertJoint.Enabled = True
    cmdModifyJoint.Enabled = True
    cmdDeletejoint.Enabled = True
    cmdAddToList.Enabled = False
    cmdOKJoint.Enabled = True
    cmdCancelJoint.Enabled = True
    If lsvJointDetail.ListItems.Count > 0 Then
        lsvJointDetail.ListItems(1).Selected = True
        Call DisplayValues
    End If
End Sub
Private Sub DisplayValues()
    If Not lsvJointDetail.ListItems.Count = 0 Then
        'Display values from listview to fields
         objDB.JointKey = Trim(lsvJointDetail.ListItems(lsvJointDetail.SelectedItem.Index).Text)
         txtJointName.Text = Trim(lsvJointDetail.ListItems(lsvJointDetail.SelectedItem.Index).ListSubItems(1).Text)
        If Trim(lsvJointDetail.ListItems(lsvJointDetail.SelectedItem.Index).ListSubItems(2).Text) = "M" Then
            optMust.Value = True
        Else
            optCan.Value = True
       End If
    End If
End Sub
Private Sub cmdDelete_Click()
    If MsgBox("Delete Selected Record", vbYesNo) = vbNo Then
        Exit Sub
    End If
    'if yes the delete the record
    If objDB.DeleteData(g_objDataSource.GetDataSource) = True Then
        MsgBox "Record Deleted Succesflly"
    End If
    'Clear all fields
    ClearControls Me
    'call initial state of form
    FormState
    bytButton = DELETE
End Sub
Private Sub cmdDeleteJoint_Click()
    'Delete the Joint Record
    With objDB
        .TypeOfAccount = Trim(Left(cboAccountType.Text, 2))
        .AccountNo = CLng(numAccountNo.Text)
    End With
    If lsvJointDetail.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Delete Selected Record", vbYesNo) = vbNo Then
        Exit Sub
    End If
    If lsvJointDetail.ListItems(lsvJointDetail.SelectedItem.Index).SubItems(3) <> "i" Then
        If objDB.DeleteJointSinleRecord(g_objDataSource.GetDataSource) = True Then
            MsgBox "Record Deleted Succesflly"
        Else
            MsgBox "Cannot Delete the Record"
            Exit Sub
        End If
    End If
    lsvJointDetail.ListItems.Remove lsvJointDetail.SelectedItem.Index
    txtJointName.Text = Empty
    optCan.Value = False
    optMust.Value = False
    If lsvJointDetail.ListItems.Count > 0 Then
        lsvJointDetail.ListItems(1).Selected = True
        Call DisplayValues
    End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdInsert_Click()
    If cboAccountType.Text = Empty Then
        MsgBox "Enter Type Of Account"
        Exit Sub
    End If
        With objDB
            .TypeOfAccount = Trim(Left(cboAccountType.Text, 2))
        End With
    Call ClearControls(Me)
    fraAllFields.Enabled = True
    fraButtonJoint.Enabled = True
    fraSevJoint.Enabled = False
    cmdLookUp.Enabled = False
    cmdInsert.Enabled = False
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    cmdOK.Enabled = True
    cmdModify.Enabled = False
    cmdCancel.Enabled = True
    bytButton = INSERT
    'get Next Account No.
    Call AccNo
    txtName.SetFocus
    cmdOK.Default = True
End Sub
Private Function DoValidations() As Boolean
    'Validate the Data
    If cboAccountType.ListIndex = -1 Or cboAccountType.Text = Empty Then
        MsgBox "Select Account Type"
        cboAccountType.SetFocus
        Exit Function
    ElseIf Val(numAccountNo.Text) <= 0 Or Trim(numAccountNo.Text) = Empty Then
        MsgBox "Enter Account No."
        numAccountNo.SetFocus
        Exit Function
    ElseIf Trim(txtName.Text) = Empty Then
        MsgBox "Enter Name of Account Holder"
        txtName.SetFocus
        Exit Function
'    ElseIf Trim(txtFName.Text) = Empty Then
'        MsgBox "Enter Father's Name"
'        txtFName.SetFocus
'        Exit Function
    ElseIf cboOccupation.ListIndex = -1 Or cboOccupation.Text = Empty Then
        MsgBox "Select Occupation of Account Holder"
        cboOccupation.SetFocus
        Exit Function
      End If
    'ElseIf Not IsDate(medDob.Text) Then
     '   MsgBox "Enter Date of birth of Account Holder"
      '  medDob.SetFocus
       ' Exit Function
        If Not Trim(medDob.Text) = "__/__/____" Then
        If Not IsDate(medDob.Text) Then
            MsgBox "Enter Date of Birth of Account Holder"
            Exit Function
        End If
    End If
'    ElseIf Trim(txtPhNo.Text) = Empty Then
'        MsgBox "Enter Ph. No. of Account Holder"
'        txtPhNo.SetFocus
'        Exit Function
    If Trim(txtAddress.Text) = Empty Then
        MsgBox "Enter Address of Account Holder"
        txtAddress.SetFocus
        Exit Function
     End If
     If Not Trim(medDOBGur.Text) = "__/__/____" Then
        If Not IsDate(medDOBGur.Text) Then
            MsgBox "Enter Date of Birth of Guardian"
            Exit Function
        End If
    End If
    If Not Trim(medDOBNom.Text) = "__/__/____" Then
        If Not IsDate(medDOBNom.Text) Then
            MsgBox "Enter Date of Birth on Nominee"
            Exit Function
        End If
    End If
    If optMinor.Value = True Then
        If txtNameGur.Text = Empty Then
            MsgBox "Enter Name of the Guardian"
            txtNameGur.SetFocus
            Exit Function
        ElseIf Not IsDate(medDOBGur.Text) Then
            MsgBox "Enter DOB of Guardian"
            medDOBGur.SetFocus
            Exit Function
         ElseIf txtRelationGur.Text = Empty Then
            MsgBox "Enter the Relation of the Guardian"
            txtRelationGur.SetFocus
            Exit Function
         End If
     End If
    If optJoint.Value = True Then
        If lsvJointDetail.ListItems.Count = 0 Then
            MsgBox "Joints are not Entered"
            Exit Function
        End If
     End If
    If Trim(numRateOfInterestC.Text) = Empty Then
        MsgBox "Enter Rate of Interest of Credit"
        numRateOfInterestC.SetFocus
        Exit Function
    ElseIf Trim(numRateOfInterestD.Text) = Empty Then
        MsgBox "Enter Rate of Interest of Debit"
        numRateOfInterestD.SetFocus
        Exit Function
    ElseIf Trim(txtIntroducedBy.Text) = Empty Then
        MsgBox "Select Account No. of Introduced by Person"
        txtIntroducedBy.SetFocus
        Exit Function
    ElseIf chkCanIntoduce.Value = Unchecked Then
        If Trim(txtReasonIntro.Text) = Empty Then
            MsgBox "Enter Reason of Cannot Introduce"
            If txtReasonIntro.Enabled = True Then
                txtReasonIntro.SetFocus
            End If
            Exit Function
        End If
     End If
    If chkClosed.Value = Checked Then
        If Trim(txtClosedReason.Text) = Empty Then
            MsgBox "Enter Reason of Closed"
            txtClosedReason.SetFocus
            Exit Function
        End If
     End If
    If Val(numLedgerNo.Text) <= 0 Or numLedgerNo.Text = Empty Then
        MsgBox "Enter Ledger No."
        numLedgerNo.SetFocus
        Exit Function
    ElseIf optJoint.Value = True Then
        If lsvJointDetail.ListItems.Count = 0 Then
            MsgBox "Joints are not Entered"
            Exit Function
        End If
    End If
    DoValidations = True
End Function
Private Sub ClearControls(frmName As Form)
        Dim Acno As Long
        Dim TypeOfAc As String
        Dim ctl As Control
        TypeOfAc = cboAccountType.Text
        Acno = numAccountNo.Text
        'Take one by one control and reset it
        For Each ctl In frmName.Controls
            If TypeOf ctl Is TextBox Then
              ctl.Text = Empty
            ElseIf TypeOf ctl Is NumberControl Then
                ctl.Text = "0"
            ElseIf TypeOf ctl Is ComboBox Then
                ctl.ListIndex = -1
            ElseIf TypeOf ctl Is CheckBox Then
                ctl.Value = False
            ElseIf TypeOf ctl Is ListView Then
                ctl.ListItems.Clear
            ElseIf TypeOf ctl Is MaskEdBox Then
                ctl.Mask = Empty
                ctl.Text = ""
                ctl.Mask = "##/##/####"
            End If
        Next
        If TypeOfAc = Empty Or IsNull(TypeOfAc) Then
            cboAccountType.ListIndex = -1
        Else
            cboAccountType.Text = TypeOfAc
        End If
     'Set Contant values
     chkCanIntoduce.Value = Checked
     numAccountNo.Text = Acno
     optAdult.Value = True
     optSingle.Value = True
     optPublic.Value = True
     optCan.Value = True
End Sub
Private Sub cmdInsertJoint_Click()
    fraLsvDetail.Enabled = False
    fraFieldsJoint.Enabled = True
    cmdInsertJoint.Enabled = False
    cmdModifyJoint.Enabled = False
    cmdDeletejoint.Enabled = False
    cmdAddToList.Enabled = True
    txtJointName.Text = Empty
    optCan.Value = True
    bytjointButton = INSERT
    txtJointName.SetFocus
End Sub
Private Sub cmdLookUp_Click()
    Dim objRecordset As ADODB.Recordset
    Dim objRecordset1 As ADODB.Recordset
    With objDB
        'Get all records in LOV
        If .PopupLov(g_objDataSource.GetDataSource) = True Then
            'if one record is selected the get its detail
            Set objRecordset = .GetACCOUNTDETAIL(g_objDataSource.GetDataSource)
         Else
            Exit Sub
        End If
        'Display Record
        Call DisplayDetail(objRecordset)
        lsvJointDetail.ListItems.Clear
        txtJointName.Text = Empty
        optCan.Value = False
        optMust.Value = False
        If optJoint.Value = True Then
            'if record if joint then get its joints and fill values in listview and fields
            Set objRecordset1 = .GETJOINTDATA(g_objDataSource.GetDataSource)
            Call FillListView(objRecordset1)
            Call DisplayValues
        End If
    End With
    cmdModify.Enabled = True
    cmdDelete.Enabled = True
End Sub
Private Sub FillListView(objRecordset As ADODB.Recordset)
    Dim objListItem As ListItem
    lsvJointDetail.ListItems.Clear
    With objRecordset
        While Not .EOF
            Set objListItem = lsvJointDetail.ListItems.Add(, , .Fields(1))
            objListItem.SubItems(1) = .Fields(0)
            objListItem.SubItems(2) = .Fields(2)
            objListItem.SubItems(3) = "d"
            .MoveNext
        Wend
    End With
    If lsvJointDetail.ListItems.Count > 0 Then
        lsvJointDetail.ListItems(1).Selected = True
    End If
End Sub
Private Sub DisplayDetail(objRecordset As ADODB.Recordset)
    JointState = 1
    'Clear all Fields
    Call ClearControls(Me)
    With objRecordset
    'display record in fields
        While Not .EOF
            cboAccountType.Text = .Fields(0) & Space(80) & "-" & .Fields(29)
            numAccountNo.Text = .Fields(1)
            txtName.Text = .Fields(2)
            If IsNull(.Fields(3)) Then
                txtFName.Text = Empty
            Else
                txtFName.Text = .Fields(3)
            End If
            txtAddress.Text = .Fields(4)
            If Not IsNull(.Fields(5)) Then
               txtPhNo.Text = .Fields(5)
            Else
                txtPhNo.Text = Empty
            End If
             If .Fields(6) = "M" Then
                optMinor.Value = True
              Else
                optAdult.Value = True
             End If
             If IsNull(.Fields(7)) Then
                medDob.Mask = Empty
                medDob.Text = Empty
                medDob.Mask = "##/##/####"
             ElseIf .Fields(7) = "__/__/____" Then
                medDob.Text = Empty
             Else
                medDob.Text = Format(.Fields(7), "dd/mm/YYYY")
             End If
             If Not IsNull(.Fields(8)) Then
                txtNameGur.Text = .Fields(8)
             Else
                txtNameGur.Text = Empty
             End If
             If Not IsNull(.Fields(9)) Then
               medDOBGur.Text = Format(.Fields(9), "dd/mm/yyyy")
             Else
                  medDOBGur.Mask = Empty
                  medDOBGur.Text = Empty
                  medDOBGur.Mask = "##/##/####"
             End If
             If Not IsNull(.Fields(10)) Then
               txtRelationGur.Text = .Fields(10)
             Else
                  txtRelationGur.Text = Empty
             End If
             If Not IsNull(.Fields(11)) Then
               txtNameNom.Text = .Fields(11)
             Else
                  txtNameNom.Text = Empty
             End If
            If Not IsNull(.Fields(12)) Then
               medDOBNom.Text = Format(.Fields(12), "dd/mm/yyyy")
            Else
                  medDOBNom.Mask = Empty
                  medDOBNom.Text = Empty
                  medDOBNom.Mask = "##/##/####"
            End If
            If Not IsNull(.Fields(13)) Then
               txtAddressNom.Text = .Fields(13)
            Else
                  txtAddressNom.Text = Empty
            End If
            If Not IsNull(.Fields(14)) Then
               txtRelationNom.Text = .Fields(14)
            Else
                  txtRelationNom.Text = Empty
            End If
            If .Fields(15) = "S" Then
               optSingle.Value = True
            Else
                optJoint.Value = True
            End If
                numRateOfInterestC.Text = .Fields(16)
                numRateOfInterestD.Text = .Fields(17)
                cboOccupation.Text = .Fields(18)
            If .Fields(19) = "S" Then
              optStaff.Value = True
            Else
                optPublic.Value = True
            End If
            txtIntroducedBy.Text = .Fields(20)
            If .Fields(21) = "Y" Then
                chkCanIntoduce.Value = Checked
            Else
                chkCanIntoduce.Value = Unchecked
            End If
            If Not IsNull(.Fields(22)) Then
                txtReasonIntro.Text = .Fields(22)
            Else
                txtReasonIntro.Text = Empty
            End If
            If Not IsNull(.Fields(23)) Then
                txtPanNo.Text = .Fields(23)
            Else
                txtPanNo.Text = Empty
            End If
            If .Fields(24) = "C" Then
                 chkClosed.Value = Checked
            Else
                 chkClosed.Value = Unchecked
            End If
            If Not IsNull(.Fields(25)) Then
                txtClosedReason.Text = .Fields(25)
            Else
                txtClosedReason.Text = Empty
            End If
            If Not IsNull(.Fields(26)) Then
                txtRemarks.Text = .Fields(26)
            Else
                 txtRemarks.Text = Empty
            End If
                numLedgerNo.Text = .Fields(27)
            If optJoint.Value = True Then
                If .Fields(28) = "J" Then
                    optJointOperate.Value = True
                 Else
                    optSeverallyOperate.Value = True
                 End If
            End If
           .MoveNext
        Wend
     End With
End Sub
Private Sub cmdModify_Click()
    'Set Controls click on Modify button
    fraAllFields.Enabled = True
    fraButtonJoint.Enabled = True
    fraSevJoint.Enabled = False
    fraTop.Enabled = False
    cmdLookUp.Enabled = False
    cmdInsert.Enabled = False
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdModify.Enabled = False
    cmdOK.Default = True
    bytButton = MODIFY
    strTypeofAccount = Trim(Left(cboAccountType.Text, 2))
    lngAccountNumber = numAccountNo.Text
End Sub
Private Sub cmdModifyJoint_Click()
    'Set Controls click on Modify of joint tab button
    If lsvJointDetail.ListItems.Count = 0 Then
        Exit Sub
    End If
     fraLsvDetail.Enabled = False
     fraFieldsJoint.Enabled = True
     cmdInsertJoint.Enabled = False
     cmdModifyJoint.Enabled = False
     cmdDeletejoint.Enabled = False
     cmdAddToList.Enabled = True
     bytjointButton = MODIFY
End Sub
Private Sub cmdOk_Click()
    Dim arrjoint() As String
    Dim i, j, k, lenarr As Integer
    'Validate the data
    If DoValidations = False Then
        Exit Sub
    End If
    With objDB
        .TypeOfAccount = Trim(Left(cboAccountType.Text, 2))
        .AccountNo = Trim(numAccountNo.Text)
        'Check Account No If it is duplicate the show mess and exit
        If bytButton = INSERT Then
            If .CheckAccount(g_objDataSource.GetDataSource) = True Then
                MsgBox "Account already Exists"
                numAccountNo.SetFocus
                Exit Sub
            End If
        End If
        'Send values to class's variables
        .NameOfAcHolder = Trim(txtName.Text)
        If Trim(txtFName.Text) = Empty Then
            .FatherName = Empty
         Else
            .FatherName = Trim(txtFName.Text)
         End If
        .Address = Trim(txtAddress.Text)
        If Trim(txtPhNo.Text) = Empty Then
            .PhNo = Empty
        Else
            .PhNo = Trim(txtPhNo.Text)
        End If
        If medDob.Text = "__/__/____" Then
            .Dob = Empty
        Else
            .Dob = medDob.FormattedText
        End If
        If optMinor.Value = True Then
           .AdultMinor = "M"
           .Guardian = Trim(txtNameGur.Text)
           .DOBGur = medDOBGur.FormattedText
           .RelationGur = Trim(txtRelationGur.Text)
        Else
            .AdultMinor = "A"
            .Guardian = Empty
            .DOBGur = Empty
            .RelationGur = Empty
         End If
        .Nominee = Trim(txtNameNom.Text)
        If Not IsDate(medDOBNom.FormattedText) Then
                .DOBNom = Empty
        Else
                .DOBNom = medDOBNom.Text
        End If
        .AddressNom = Trim(txtAddressNom.Text)
        .RelationNom = Trim(txtRelationNom)
        If optSingle.Value = True Then
             .SingleJoint = "S"
        Else
            .SingleJoint = "J"
        End If
        .RateOfInterestC = numRateOfInterestC.Text
        .RateOfInterestD = numRateOfInterestD.Text
        .Occupation = Val(cboOccupation.ItemData(cboOccupation.ListIndex))
         If optStaff.Value = True Then
            .StaffPublic = "S"
         Else
            .StaffPublic = "P"
         End If
        .IntroducedBy = txtIntroducedBy.Text
         If chkCanIntoduce.Value = Checked Then
            .CanIntroduce = "Y"
         Else
            .CanIntroduce = "N"
         End If
         .ReasonIntro = Trim(txtReasonIntro.Text)
         .Pan = Trim(txtPanNo.Text)
         If chkClosed.Value = Checked Then
            .ClosedOperative = "C"
         Else
            .ClosedOperative = "O"
         End If
         .ReasonClosed = Trim(txtClosedReason.Text)
         .Remarks = Trim(txtRemarks.Text)
         .LedgerNo = Val(numLedgerNo.Text)
         If optJoint.Value = True Then
            If optSeverallyOperate.Value = True Then
                .Severally_Joint = "S"
            Else
                .Severally_Joint = "J"
            End If
         Else
            .Severally_Joint = Empty
         End If
         'get Gl No from cboAccount Type
         .Gl_No = FetchParentKey(cboAccountType.Text)
         .TerminalName = strComputerName
         .UserName = strUserName
         .InsertModifyDate = dteTodaysDate
         lenarr = 0
         For i = 1 To lsvJointDetail.ListItems.Count
            If lsvJointDetail.ListItems(i).SubItems(3) <> "d" Then
            'Get the length of array lenarr
                lenarr = lenarr + 4
            End If
          Next
          ReDim arrjoint(lenarr)
          j = 1
          k = 0
          For i = 0 To (lsvJointDetail.ListItems.Count * 4) - 1 Step 4
             If lsvJointDetail.ListItems(j).SubItems(3) <> "d" Then
                'Fill all the jointdetails into array
                arrjoint(k) = lsvJointDetail.ListItems(j).Text
                arrjoint(k + 1) = lsvJointDetail.ListItems(j).SubItems(1)
                arrjoint(k + 2) = lsvJointDetail.ListItems(j).SubItems(2)
                arrjoint(k + 3) = lsvJointDetail.ListItems(j).SubItems(3)
                k = k + 4
            End If
            j = j + 1
           Next
           'Send the array to class
           .JointData = arrjoint
          If bytButton = INSERT Then
                'Insert record
              .INSERTModifyDATA g_objDataSource.GetDataSource, INSERT
              If optJoint.Value = True Then
                'insert joint detail
                .INSERTModifyJointData g_objDataSource.GetDataSource, INSERT
              End If
         ElseIf bytButton = MODIFY Then
             'Save Modified data
              .INSERTModifyDATA g_objDataSource.GetDataSource, MODIFY
            If optJoint.Value = True Then
                'Save Modified joint Data
                .INSERTModifyJointData g_objDataSource.GetDataSource, MODIFY
            End If
          End If
        If optSingle.Value = True Then
            'If record is Modified from joint to single then delete all joint details
            If .DeleteJointAllRecords(g_objDataSource.GetDataSource) = False Then
                MsgBox "Error In Transaction"
            End If
        End If
    End With
        Call FormState
        Call ClearControls(Me)
        bytFillAgain = 1
        'Get same record from database
        Call GetSinleAccount
        cmdInsert.SetFocus
End Sub
Private Sub cmdOkJoint_Click()
    fraFieldsJoint.Enabled = False
    fraLsvDetail.Enabled = True
    cmdInsertJoint.Enabled = True
    cmdModifyJoint.Enabled = True
    cmdDeletejoint.Enabled = True
    cmdAddToList.Enabled = False
    cmdOKJoint.Enabled = True
    cmdCancelJoint.Enabled = True
    ssTabAccOpening.Tab = 0
End Sub
Private Sub FillCombo()
'    With cboAccountType
'        .AddItem "SB"
'        .AddItem "CC"
'        .AddItem "CA"
'        .AddItem "OD"
'    End With
End Sub
Private Sub AccNo()
    'Get Max Account no from database
  numAccountNo.Text = (objDB.GETAccNo(g_objDataSource.GetDataSource)) + 1
End Sub
Private Sub GetOccupation()
     Dim objRecordset As ADODB.Recordset
     Set objRecordset = objDB.GetOccupation(g_objDataSource.GetDataSource)
     With objRecordset
        'If any record is fetched then fill it in combo
        While Not .EOF
            'Add category names in combo
            cboOccupation.AddItem Trim(.Fields("OCCUPATION_DESC"))
            cboOccupation.ItemData(cboOccupation.NewIndex) = Trim(.Fields("OCCUPATION_Key").Value)
            .MoveNext
         Wend
     End With
End Sub
Private Sub GetAccountType()
     Dim objRecordset As ADODB.Recordset
     Set objRecordset = objDB.GetAccountType(g_objDataSource.GetDataSource)
    With objRecordset
        'If any record is fetched then fill it in combo
        While Not .EOF
            'Add category names in combo
            cboAccountType.AddItem Left((Trim(.Fields(1))), 2) & Space(80) & "-" & .Fields(0)
            .MoveNext
         Wend
    End With
End Sub
Private Function FetchParentKey(STR As String) As String
    'Get the Key from the String
    Dim lngstartingpos As Long
    Dim strParentKey As String
    lngstartingpos = InStrRev(STR, "-")
    strParentKey = Mid(STR, lngstartingpos + 1, Len(STR) - lngstartingpos)
    FetchParentKey = Trim(strParentKey)
End Function
Private Function FetchAccType(STR As String) As String
    'Dim lngstartingpos As Long
    'Dim strParentKey As String
    'lngstartingpos = InStrRev(STR, "-")
    'strParentKey = Mid(STR, lngstartingpos + 1, Len(STR) - lngstartingpos)
    FetchAccType = Trim(Left(STR, 20))
End Function
Private Sub Form_Load()
    'Fill Occupation Combo Box
    Call GetOccupation
    'Set forms initial state
    Call FormState
    'Get Account type and its Gl No
    Call GetAccountType
    lblDateTime.Caption = strTodaysDate
    lblUserName.Caption = strUserName
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
        If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessageLong(Me.hwnd, &HA1, 2, 0&)
        End If
End Sub

Private Sub lsvJointDetail_Click()
    Call DisplayValues
End Sub
Private Sub numAccountNO_KeyPress(KeyAscii As Integer)
    If cboAccountType.Text <> Empty And numAccountNo.Text <> Empty Then
        If KeyAscii = vbKeyReturn Then
            'Get Single Record
            Call GetSinleAccount
            bytFillAgain = 0
        End If
    End If
End Sub
Private Sub numAccountNo_LostFocus()
    If cboAccountType.Text <> Empty And numAccountNo.Text <> Empty Then
        If KeyAscii = vbKeyReturn Then
            'Get Single Record
            Call GetSinleAccount
            bytFillAgain = 0
        End If
    End If
End Sub
Private Sub optAdult_Click()
    If optAdult.Value = True Then
        fraGuardian.Enabled = False
        txtNameGur.Text = Empty
        medDOBGur.Mask = Empty
        medDOBGur.Text = Empty
        medDOBGur.Mask = "##/##/####"
        txtRelationGur.Text = Empty
    Else
        fraGuardian.Enabled = True
    End If
End Sub
Private Sub optJoint_Click()
    If JointState = 1 Then
        JointState = 0
        Exit Sub
    End If
    If optJoint.Value = True Then
        fraSevJoint.Enabled = True
        ssTabAccOpening.Tab = 1
        optSeverallyOperate.Value = True
    End If
End Sub
Private Sub optMinor_Click()
    If optMinor.Value = True Then
        fraGuardian.Enabled = True
    Else
        fraGuardian.Enabled = False
    End If
End Sub
Private Sub optSingle_Click()
    fraSevJoint.Enabled = False
    optSeverallyOperate.Value = False
    optJointOperate.Value = False
End Sub
Private Sub ssTabAccOpening_Click(PreviousTab As Integer)
    'Set Enable of controls
    If bytjointButton = INSERT Or bytjointButton = MODIFY Then
        If optJoint.Enabled = True Then
            optJoint.SetFocus
        End If
        Exit Sub
    End If
    fraLsvDetail.Enabled = True
    If bytButton = INSERT Or bytButton = MODIFY Then
        If ssTabAccOpening.Tab = 1 Then
            If optJoint.Value = False Then
                fraFieldsJoint.Enabled = False
                fraButtonJoint.Enabled = False
                
            Else
                fraFieldsJoint.Enabled = True
                fraButtonJoint.Enabled = True
                If cmdInsertJoint.Enabled = True Then
                    cmdInsertJoint.SetFocus
                End If
            End If
         End If
     End If
   If ssTabAccOpening.Tab = 1 Then
        If bytButton = INSERT Or bytButton = MODIFY Then
            If bytjointButton = INSERT Or bytjointButton = MODIFY Then
                fraFieldsJoint.Enabled = True
            Else
                fraFieldsJoint.Enabled = False
            End If
         End If
    End If
End Sub
