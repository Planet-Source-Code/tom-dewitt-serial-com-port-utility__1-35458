VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Serial Communication Test Utility"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13785
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Tag             =   "15480"
   Begin VB.Frame fraPort 
      Caption         =   "Communication Port B Is Closed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Index           =   1
      Left            =   6960
      TabIndex        =   129
      Top             =   25
      Width           =   6735
      Begin VB.Frame fraSettings 
         Caption         =   "Port 5 Settings 57600, Even, 8, 2, String, XOn/XOff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   1
         Left            =   120
         TabIndex        =   225
         Top             =   360
         Width           =   6495
         Begin VB.Frame fraSend 
            Caption         =   "Data To Send"
            Height          =   1455
            Index           =   1
            Left            =   120
            TabIndex        =   226
            Top             =   240
            Visible         =   0   'False
            Width           =   6255
            Begin VB.TextBox txtSendTime 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   2520
               MaxLength       =   5
               TabIndex        =   257
               Text            =   "0"
               Top             =   1035
               Width           =   615
            End
            Begin VB.CommandButton cmdSettings 
               Caption         =   "Settings"
               Height          =   375
               Index           =   1
               Left            =   5400
               TabIndex        =   234
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtMultiple 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   1680
               MaxLength       =   6
               TabIndex        =   233
               Text            =   "0"
               Top             =   1035
               Width           =   735
            End
            Begin VB.CommandButton cmdMultiSend 
               Caption         =   "Start"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   4560
               TabIndex        =   232
               Top             =   960
               Width           =   735
            End
            Begin VB.ComboBox cmbMultiSend 
               Height          =   315
               Index           =   1
               ItemData        =   "frmMain.frx":0442
               Left            =   3240
               List            =   "frmMain.frx":044F
               TabIndex        =   231
               Text            =   "Once"
               Top             =   1005
               Width           =   1215
            End
            Begin VB.TextBox txtAscii 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   230
               Top             =   960
               Width           =   495
            End
            Begin VB.CommandButton cmdAscii 
               Caption         =   "Char Code"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   229
               Top             =   960
               Width           =   975
            End
            Begin VB.CommandButton cmdSend 
               Caption         =   "Send"
               Height          =   375
               Index           =   1
               Left            =   5400
               TabIndex        =   228
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtSend 
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   227
               Top             =   360
               Width           =   5175
            End
            Begin VB.Label lblSendInterval 
               AutoSize        =   -1  'True
               Caption         =   "Interval"
               Height          =   195
               Index           =   1
               Left            =   2565
               TabIndex        =   236
               Top             =   840
               Width           =   525
            End
            Begin VB.Label lblTimes 
               AutoSize        =   -1  'True
               Caption         =   "Iterations"
               Height          =   195
               Index           =   1
               Left            =   1725
               TabIndex        =   235
               Top             =   840
               Width           =   645
            End
         End
         Begin VB.ComboBox cmbHandShake 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMain.frx":046F
            Left            =   4920
            List            =   "frmMain.frx":047F
            TabIndex        =   249
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cmbBaudRate 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMain.frx":04A2
            Left            =   1200
            List            =   "frmMain.frx":04D3
            TabIndex        =   248
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cmbStopBits 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMain.frx":0539
            Left            =   4080
            List            =   "frmMain.frx":0543
            TabIndex        =   247
            Top             =   600
            Width           =   735
         End
         Begin VB.ComboBox cmbDataBits 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMain.frx":054D
            Left            =   3240
            List            =   "frmMain.frx":0560
            TabIndex        =   246
            Top             =   600
            Width           =   735
         End
         Begin VB.ComboBox cmbParity 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMain.frx":0573
            Left            =   2280
            List            =   "frmMain.frx":0586
            TabIndex        =   245
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cmbPort 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMain.frx":05A8
            Left            =   240
            List            =   "frmMain.frx":05AA
            TabIndex        =   244
            Top             =   600
            Width           =   855
         End
         Begin VB.Frame fraInputMode 
            Caption         =   "Input"
            Height          =   615
            Index           =   1
            Left            =   240
            TabIndex        =   241
            Top             =   960
            Width           =   1815
            Begin VB.OptionButton optString 
               Caption         =   "String"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   243
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optBinary 
               Caption         =   "Binary"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   242
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "Open"
            Height          =   495
            Index           =   1
            Left            =   3480
            TabIndex        =   240
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Enabled         =   0   'False
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   239
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton cmdShowSend 
            Caption         =   "Send Data"
            Enabled         =   0   'False
            Height          =   495
            Index           =   1
            Left            =   5400
            TabIndex        =   238
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cmbParityReplace 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   1
            Left            =   2160
            TabIndex        =   237
            Top             =   1192
            Width           =   1215
         End
         Begin VB.Label lblHandShake 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Hand Shaking"
            Height          =   195
            Index           =   1
            Left            =   5077
            TabIndex        =   256
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblBaudRate 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Baud Rate"
            Height          =   195
            Index           =   1
            Left            =   1305
            TabIndex        =   255
            Top             =   360
            Width           =   765
         End
         Begin VB.Label lblStopBits 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Stop Bits"
            Height          =   195
            Index           =   1
            Left            =   4125
            TabIndex        =   254
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblDataBits 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Data Bits"
            Height          =   195
            Index           =   1
            Left            =   3285
            TabIndex        =   253
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lblParity 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Parity"
            Height          =   195
            Index           =   1
            Left            =   2505
            TabIndex        =   252
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblPort 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Port"
            Height          =   195
            Index           =   1
            Left            =   450
            TabIndex        =   251
            Top             =   360
            Width           =   315
         End
         Begin VB.Label lblParityReplace 
            AutoSize        =   -1  'True
            Caption         =   "Parity Replace"
            Height          =   195
            Index           =   1
            Left            =   2250
            TabIndex        =   250
            Top             =   975
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Receive"
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   175
         Top             =   7800
         Width           =   975
      End
      Begin VB.ListBox lstData 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Index           =   1
         ItemData        =   "frmMain.frx":05AC
         Left            =   120
         List            =   "frmMain.frx":05AE
         TabIndex        =   174
         Top             =   2400
         Width           =   6495
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Output"
         Height          =   855
         Index           =   1
         Left            =   1320
         TabIndex        =   171
         Top             =   8520
         Width           =   975
         Begin VB.OptionButton optBinaryOut 
            Caption         =   "Binary"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   173
            Top             =   480
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optStringOut 
            Caption         =   "String"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   172
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   170
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Frame fraRThreshold 
         Caption         =   "Receive Threshold Count"
         Height          =   615
         Index           =   1
         Left            =   1320
         TabIndex        =   166
         Top             =   7080
         Width           =   2535
         Begin VB.CommandButton cmdApplyRThresh 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   169
            Top             =   225
            Width           =   735
         End
         Begin VB.TextBox txtRThreshold 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   168
            Top             =   225
            Width           =   615
         End
         Begin VB.CheckBox chkRThreshold 
            Caption         =   "Enable"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   167
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame fraLineControl 
         Caption         =   "Line Toggle"
         Height          =   1695
         Index           =   1
         Left            =   120
         TabIndex        =   162
         Top             =   7680
         Width           =   1095
         Begin VB.CommandButton cmdDTR 
            Caption         =   "DTR"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   165
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdRTS 
            Caption         =   "RTS"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   164
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdBreak 
            Caption         =   "BREAK"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   163
            Top             =   240
            Width           =   735
         End
         Begin VB.Shape shpDTR 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   840
            Top             =   1200
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape shpRTS 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   840
            Top             =   720
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape shpBreak 
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   840
            Top             =   240
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape shpDTROff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   840
            Top             =   1200
            Width           =   135
         End
         Begin VB.Shape shpRTSOff 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   840
            Top             =   720
            Width           =   135
         End
         Begin VB.Shape shpBreakOff 
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   840
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Frame fraHolding 
         Caption         =   "Holding"
         Height          =   975
         Index           =   1
         Left            =   1320
         TabIndex        =   158
         Top             =   9480
         Width           =   975
         Begin VB.Label lblCD 
            AutoSize        =   -1  'True
            Caption         =   "CD"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   161
            Top             =   720
            Width           =   210
         End
         Begin VB.Shape shpCD 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   765
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblDSR 
            AutoSize        =   -1  'True
            Caption         =   "DSR"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   160
            Top             =   480
            Width           =   315
         End
         Begin VB.Shape shpDSR 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblCTS 
            AutoSize        =   -1  'True
            Caption         =   "CTS"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   159
            Top             =   240
            Width           =   315
         End
         Begin VB.Shape shpCTS 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   285
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpCDOff 
            FillColor       =   &H00008080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   765
            Width           =   255
         End
         Begin VB.Shape shpDSROff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   525
            Width           =   255
         End
         Begin VB.Shape shpCTSOff 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   285
            Width           =   255
         End
      End
      Begin VB.Frame fraLineDetect 
         Caption         =   "Detection"
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   154
         Top             =   9480
         Width           =   1095
         Begin VB.Timer tmrLineDetect 
            Enabled         =   0   'False
            Index           =   1
            Interval        =   250
            Left            =   840
            Top             =   720
         End
         Begin VB.Shape shpBreakEvent 
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   285
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblBreak 
            AutoSize        =   -1  'True
            Caption         =   "Break"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   157
            Top             =   240
            Width           =   525
         End
         Begin VB.Shape shpRing 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblRing 
            AutoSize        =   -1  'True
            Caption         =   "Ring"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   156
            Top             =   480
            Width           =   420
         End
         Begin VB.Shape shpEOF 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   765
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblEOF 
            AutoSize        =   -1  'True
            Caption         =   "EOF"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   155
            Top             =   720
            Width           =   315
         End
         Begin VB.Shape shpBreakEventOff 
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   285
            Width           =   255
         End
         Begin VB.Shape shpRingOff 
            FillColor       =   &H00008080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   525
            Width           =   255
         End
         Begin VB.Shape shpEOFOff 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   765
            Width           =   255
         End
      End
      Begin VB.Frame fraPoll 
         Caption         =   "Poll Port"
         Height          =   2655
         Index           =   1
         Left            =   2400
         TabIndex        =   147
         Top             =   7800
         Width           =   1095
         Begin VB.TextBox txtBuffer 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   4
            TabIndex        =   151
            Text            =   "0"
            Top             =   2200
            Width           =   855
         End
         Begin VB.CommandButton cmdApplyPoll 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   150
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtPoll 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   4
            TabIndex        =   149
            Text            =   "0"
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkPoll 
            Caption         =   "Enable"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   148
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblReadAt 
            Alignment       =   1  'Right Justify
            Caption         =   "Read When Buffer Has A Count Of :    "
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   153
            Top             =   1560
            Width           =   875
         End
         Begin VB.Label lblInterval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Time Interval "
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   152
            Top             =   960
            Width           =   960
         End
      End
      Begin VB.Frame fraInputLen 
         Caption         =   "Read x Characters At A Time"
         Height          =   615
         Index           =   1
         Left            =   3960
         TabIndex        =   143
         Top             =   7080
         Width           =   2655
         Begin VB.TextBox txtCount 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   146
            Top             =   225
            Width           =   855
         End
         Begin VB.OptionButton optCount 
            Caption         =   "Count Of"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   145
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "ALL"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   144
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.PictureBox picScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   833
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":05B0
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   431
         TabIndex        =   142
         Top             =   6000
         Width           =   6495
      End
      Begin VB.Frame fraComError 
         Caption         =   "Port Errors"
         Height          =   2655
         Index           =   1
         Left            =   3600
         TabIndex        =   131
         Top             =   7800
         Width           =   1225
         Begin VB.Timer tmrErrors 
            Enabled         =   0   'False
            Index           =   1
            Interval        =   250
            Left            =   720
            Top             =   1440
         End
         Begin VB.Shape shpSThresh 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   2445
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblSThresh 
            AutoSize        =   -1  'True
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   141
            Top             =   2400
            Width           =   420
         End
         Begin VB.Shape shpFrame 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   285
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblFrame 
            AutoSize        =   -1  'True
            Caption         =   "Frame"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   140
            Top             =   240
            Width           =   525
         End
         Begin VB.Shape shpParity 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblParityError 
            AutoSize        =   -1  'True
            Caption         =   "Parity"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   139
            Top             =   480
            Width           =   630
         End
         Begin VB.Shape shpOverRun 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   765
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblOver 
            AutoSize        =   -1  'True
            Caption         =   "Loss"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   138
            Top             =   720
            Width           =   420
         End
         Begin VB.Shape shpCTSTO 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1005
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblCTSTO 
            AutoSize        =   -1  'True
            Caption         =   "CTS"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   137
            Top             =   960
            Width           =   315
         End
         Begin VB.Shape shpDSRTO 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1245
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblDSRTO 
            AutoSize        =   -1  'True
            Caption         =   "DSR"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   136
            Top             =   1200
            Width           =   315
         End
         Begin VB.Shape shpCDTO 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1485
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblCDTO 
            AutoSize        =   -1  'True
            Caption         =   "CD"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   135
            Top             =   1440
            Width           =   210
         End
         Begin VB.Shape shpRXOver 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1725
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblRxOver 
            AutoSize        =   -1  'True
            Caption         =   "RX"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   134
            Top             =   1680
            Width           =   210
         End
         Begin VB.Shape shpTXFull 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1965
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblTxFull 
            AutoSize        =   -1  'True
            Caption         =   "TX"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   133
            Top             =   1920
            Width           =   210
         End
         Begin VB.Shape shpDCB 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   2205
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblDCB 
            AutoSize        =   -1  'True
            Caption         =   "DCB"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   132
            Top             =   2160
            Width           =   315
         End
         Begin VB.Shape shpSThreshOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   2445
            Width           =   255
         End
         Begin VB.Shape shpOverRunOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   765
            Width           =   255
         End
         Begin VB.Shape shpParityOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   525
            Width           =   255
         End
         Begin VB.Shape shpFrameOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   285
            Width           =   255
         End
         Begin VB.Shape shpCTSTOOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1005
            Width           =   255
         End
         Begin VB.Shape shpDSRTOOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1245
            Width           =   255
         End
         Begin VB.Shape shpCDTOOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1485
            Width           =   255
         End
         Begin VB.Shape shpRXOverOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1725
            Width           =   255
         End
         Begin VB.Shape shpYXFullOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   1965
            Width           =   255
         End
         Begin VB.Shape shpDCBOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   2205
            Width           =   255
         End
      End
      Begin VB.ListBox lstErrors 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   2595
         Index           =   1
         ItemData        =   "frmMain.frx":11C62
         Left            =   4800
         List            =   "frmMain.frx":11C64
         TabIndex        =   130
         Top             =   7875
         Width           =   1815
      End
      Begin MSCommLib.MSComm comPort 
         Index           =   1
         Left            =   6120
         Top             =   10200
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
         ParityReplace   =   35
         BaudRate        =   1200
         InputMode       =   1
      End
      Begin VB.Frame fraReceive 
         Caption         =   "Data Received"
         Height          =   4455
         Index           =   1
         Left            =   120
         TabIndex        =   176
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Index           =   1
         Left            =   5160
         Top             =   10320
      End
      Begin VB.Timer tmrPoll 
         Enabled         =   0   'False
         Index           =   1
         Left            =   5640
         Top             =   10320
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   95
         Left            =   6480
         TabIndex        =   224
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   94
         Left            =   6330
         TabIndex        =   223
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   93
         Left            =   6195
         TabIndex        =   222
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   92
         Left            =   6060
         TabIndex        =   221
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   91
         Left            =   5925
         TabIndex        =   220
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   90
         Left            =   5790
         TabIndex        =   219
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   89
         Left            =   5655
         TabIndex        =   218
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   88
         Left            =   5520
         TabIndex        =   217
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   87
         Left            =   5385
         TabIndex        =   216
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   86
         Left            =   5250
         TabIndex        =   215
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   85
         Left            =   5115
         TabIndex        =   214
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   84
         Left            =   4980
         TabIndex        =   213
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   83
         Left            =   4845
         TabIndex        =   212
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   82
         Left            =   4710
         TabIndex        =   211
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   81
         Left            =   4575
         TabIndex        =   210
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   80
         Left            =   4440
         TabIndex        =   209
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   79
         Left            =   4305
         TabIndex        =   208
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   78
         Left            =   4170
         TabIndex        =   207
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   77
         Left            =   4035
         TabIndex        =   206
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   76
         Left            =   3900
         TabIndex        =   205
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   75
         Left            =   3765
         TabIndex        =   204
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   74
         Left            =   3630
         TabIndex        =   203
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   73
         Left            =   3495
         TabIndex        =   202
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   72
         Left            =   3360
         TabIndex        =   201
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   71
         Left            =   3225
         TabIndex        =   200
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   70
         Left            =   3090
         TabIndex        =   199
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   69
         Left            =   2955
         TabIndex        =   198
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   68
         Left            =   2820
         TabIndex        =   197
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   67
         Left            =   2685
         TabIndex        =   196
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   66
         Left            =   2550
         TabIndex        =   195
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   65
         Left            =   2415
         TabIndex        =   194
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   64
         Left            =   2280
         TabIndex        =   193
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   63
         Left            =   2145
         TabIndex        =   192
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   62
         Left            =   2010
         TabIndex        =   191
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   61
         Left            =   1875
         TabIndex        =   190
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   60
         Left            =   1740
         TabIndex        =   189
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   59
         Left            =   1605
         TabIndex        =   188
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   58
         Left            =   1470
         TabIndex        =   187
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   57
         Left            =   1335
         TabIndex        =   186
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   56
         Left            =   1200
         TabIndex        =   185
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   55
         Left            =   1065
         TabIndex        =   184
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   54
         Left            =   930
         TabIndex        =   183
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   53
         Left            =   795
         TabIndex        =   182
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   52
         Left            =   660
         TabIndex        =   181
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   51
         Left            =   525
         TabIndex        =   180
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   50
         Left            =   390
         TabIndex        =   179
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   49
         Left            =   255
         TabIndex        =   178
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   48
         Left            =   120
         TabIndex        =   177
         Top             =   6840
         Width           =   135
      End
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Port 5 Settings 57600, Even, 8, 2, String, XOn/XOff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   240
      TabIndex        =   96
      Top             =   360
      Width           =   6495
      Begin VB.Frame fraSend 
         Caption         =   "Data To Send"
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   117
         Top             =   240
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox txtSend 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   126
            Top             =   360
            Width           =   5175
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Height          =   375
            Index           =   0
            Left            =   5400
            TabIndex        =   125
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdAscii 
            Caption         =   "Char Code"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   124
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtAscii 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   123
            Top             =   960
            Width           =   495
         End
         Begin VB.ComboBox cmbMultiSend 
            Height          =   315
            Index           =   0
            ItemData        =   "frmMain.frx":11C66
            Left            =   3240
            List            =   "frmMain.frx":11C73
            TabIndex        =   122
            Text            =   "Once"
            Top             =   1005
            Width           =   1215
         End
         Begin VB.CommandButton cmdMultiSend 
            Caption         =   "Start"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   4560
            TabIndex        =   121
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtMultiple 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   120
            Text            =   "0"
            Top             =   1035
            Width           =   735
         End
         Begin VB.TextBox txtSendTime 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2520
            MaxLength       =   5
            TabIndex        =   119
            Text            =   "0"
            Top             =   1035
            Width           =   615
         End
         Begin VB.CommandButton cmdSettings 
            Caption         =   "Settings"
            Height          =   375
            Index           =   0
            Left            =   5400
            TabIndex        =   118
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblTimes 
            AutoSize        =   -1  'True
            Caption         =   "Iterations"
            Height          =   195
            Index           =   0
            Left            =   1725
            TabIndex        =   128
            Top             =   840
            Width           =   645
         End
         Begin VB.Label lblSendInterval 
            AutoSize        =   -1  'True
            Caption         =   "Interval"
            Height          =   195
            Index           =   0
            Left            =   2565
            TabIndex        =   127
            Top             =   840
            Width           =   525
         End
      End
      Begin VB.ComboBox cmbParityReplace 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2160
         TabIndex        =   115
         Top             =   1192
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowSend 
         Caption         =   "Send Data"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   5400
         TabIndex        =   114
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   4440
         TabIndex        =   113
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   495
         Index           =   0
         Left            =   3480
         TabIndex        =   112
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame fraInputMode 
         Caption         =   "Input"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   109
         Top             =   960
         Width           =   1815
         Begin VB.OptionButton optBinary 
            Caption         =   "Binary"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   111
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optString 
            Caption         =   "String"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbPort 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":11C93
         Left            =   240
         List            =   "frmMain.frx":11C95
         TabIndex        =   107
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbParity 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":11C97
         Left            =   2280
         List            =   "frmMain.frx":11CAA
         TabIndex        =   101
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbDataBits 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":11CCC
         Left            =   3240
         List            =   "frmMain.frx":11CDF
         TabIndex        =   100
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cmbStopBits 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":11CF2
         Left            =   4080
         List            =   "frmMain.frx":11CFC
         TabIndex        =   99
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cmbBaudRate 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":11D06
         Left            =   1200
         List            =   "frmMain.frx":11D37
         TabIndex        =   98
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbHandShake 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":11D9D
         Left            =   4920
         List            =   "frmMain.frx":11DAD
         TabIndex        =   97
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblParityReplace 
         AutoSize        =   -1  'True
         Caption         =   "Parity Replace"
         Height          =   195
         Index           =   0
         Left            =   2250
         TabIndex        =   116
         Top             =   975
         Width           =   1035
      End
      Begin VB.Label lblPort 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   108
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblParity 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Parity"
         Height          =   195
         Index           =   0
         Left            =   2505
         TabIndex        =   106
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblDataBits 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Data Bits"
         Height          =   195
         Index           =   0
         Left            =   3285
         TabIndex        =   105
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblStopBits 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Stop Bits"
         Height          =   195
         Index           =   0
         Left            =   4125
         TabIndex        =   104
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblBaudRate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Baud Rate"
         Height          =   195
         Index           =   0
         Left            =   1305
         TabIndex        =   103
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblHandShake 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Hand Shaking"
         Height          =   195
         Index           =   0
         Left            =   5077
         TabIndex        =   102
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame fraPort 
      Caption         =   "Communication Port A Is Closed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   25
      Width           =   6735
      Begin VB.ListBox lstErrors 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   2595
         Index           =   0
         ItemData        =   "frmMain.frx":11DD0
         Left            =   4800
         List            =   "frmMain.frx":11DD2
         TabIndex        =   95
         Top             =   7875
         Width           =   1815
      End
      Begin VB.Frame fraComError 
         Caption         =   "Port Errors"
         Height          =   2655
         Index           =   0
         Left            =   3600
         TabIndex        =   84
         Top             =   7800
         Width           =   1225
         Begin VB.Timer tmrErrors 
            Enabled         =   0   'False
            Index           =   0
            Interval        =   250
            Left            =   720
            Top             =   1440
         End
         Begin VB.Shape shpSThresh 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   2445
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpSThreshOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   2445
            Width           =   255
         End
         Begin VB.Label lblSThresh 
            AutoSize        =   -1  'True
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   94
            Top             =   2400
            Width           =   420
         End
         Begin VB.Shape shpFrame 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   285
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblFrame 
            AutoSize        =   -1  'True
            Caption         =   "Frame"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   93
            Top             =   240
            Width           =   525
         End
         Begin VB.Shape shpParity 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblParityError 
            AutoSize        =   -1  'True
            Caption         =   "Parity"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   92
            Top             =   480
            Width           =   630
         End
         Begin VB.Shape shpOverRun 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   765
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblOver 
            AutoSize        =   -1  'True
            Caption         =   "Loss"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   91
            Top             =   720
            Width           =   420
         End
         Begin VB.Shape shpCTSTO 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1005
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblCTSTO 
            AutoSize        =   -1  'True
            Caption         =   "CTS"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   90
            Top             =   960
            Width           =   315
         End
         Begin VB.Shape shpDSRTO 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1245
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblDSRTO 
            AutoSize        =   -1  'True
            Caption         =   "DSR"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   89
            Top             =   1200
            Width           =   315
         End
         Begin VB.Shape shpCDTO 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1485
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblCDTO 
            AutoSize        =   -1  'True
            Caption         =   "CD"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   88
            Top             =   1440
            Width           =   210
         End
         Begin VB.Shape shpRXOver 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1725
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblRxOver 
            AutoSize        =   -1  'True
            Caption         =   "RX"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   87
            Top             =   1680
            Width           =   210
         End
         Begin VB.Shape shpTXFull 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1965
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblTxFull 
            AutoSize        =   -1  'True
            Caption         =   "TX"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   86
            Top             =   1920
            Width           =   210
         End
         Begin VB.Shape shpDCB 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   2205
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblDCB 
            AutoSize        =   -1  'True
            Caption         =   "DCB"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   85
            Top             =   2160
            Width           =   315
         End
         Begin VB.Shape shpOverRunOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   765
            Width           =   255
         End
         Begin VB.Shape shpParityOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   525
            Width           =   255
         End
         Begin VB.Shape shpFrameOff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   285
            Width           =   255
         End
         Begin VB.Shape shpCTSTOOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1005
            Width           =   255
         End
         Begin VB.Shape shpDSRTOOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1245
            Width           =   255
         End
         Begin VB.Shape shpCDTOOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1485
            Width           =   255
         End
         Begin VB.Shape shpRXOverOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1725
            Width           =   255
         End
         Begin VB.Shape shpYXFullOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   1965
            Width           =   255
         End
         Begin VB.Shape shpDCBOff 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   2205
            Width           =   255
         End
      End
      Begin VB.PictureBox picScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   833
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":11DD4
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   431
         TabIndex        =   34
         Top             =   6000
         Width           =   6495
      End
      Begin VB.Frame fraInputLen 
         Caption         =   "Read x Characters At A Time"
         Height          =   615
         Index           =   0
         Left            =   3960
         TabIndex        =   30
         Top             =   7080
         Width           =   2655
         Begin VB.TextBox txtCount 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   33
            Top             =   225
            Width           =   855
         End
         Begin VB.OptionButton optCount 
            Caption         =   "Count Of"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "ALL"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Index           =   0
         Left            =   5160
         Top             =   10320
      End
      Begin VB.Frame fraPoll 
         Caption         =   "Poll Port"
         Height          =   2655
         Index           =   0
         Left            =   2400
         TabIndex        =   23
         Top             =   7800
         Width           =   1095
         Begin VB.TextBox txtBuffer 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   29
            Text            =   "0"
            Top             =   2200
            Width           =   855
         End
         Begin VB.CommandButton cmdApplyPoll 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtPoll 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   25
            Text            =   "0"
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkPoll 
            Caption         =   "Enable"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblReadAt 
            Alignment       =   1  'Right Justify
            Caption         =   "Read When Buffer Has A Count Of :    "
            Height          =   555
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   875
         End
         Begin VB.Label lblInterval 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Time Interval "
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   27
            Top             =   960
            Width           =   960
         End
      End
      Begin VB.Frame fraLineDetect 
         Caption         =   "Detection"
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   9480
         Width           =   1095
         Begin VB.Timer tmrLineDetect 
            Enabled         =   0   'False
            Index           =   0
            Interval        =   250
            Left            =   840
            Top             =   720
         End
         Begin VB.Label lblEOF 
            AutoSize        =   -1  'True
            Caption         =   "EOF"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   83
            Top             =   720
            Width           =   315
         End
         Begin VB.Shape shpEOF 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   765
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpEOFOff 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblRing 
            AutoSize        =   -1  'True
            Caption         =   "Ring"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   21
            Top             =   480
            Width           =   420
         End
         Begin VB.Shape shpRing 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpRingOff 
            FillColor       =   &H00008080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblBreak 
            AutoSize        =   -1  'True
            Caption         =   "Break"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   20
            Top             =   240
            Width           =   525
         End
         Begin VB.Shape shpBreakEvent 
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   285
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpBreakEventOff 
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   285
            Width           =   255
         End
      End
      Begin VB.Frame fraHolding 
         Caption         =   "Holding"
         Height          =   975
         Index           =   0
         Left            =   1320
         TabIndex        =   15
         Top             =   9480
         Width           =   975
         Begin VB.Shape shpCTS 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   285
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpCTSOff 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   285
            Width           =   255
         End
         Begin VB.Label lblCTS 
            AutoSize        =   -1  'True
            Caption         =   "CTS"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   18
            Top             =   240
            Width           =   315
         End
         Begin VB.Shape shpDSR 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   525
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpDSROff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   525
            Width           =   255
         End
         Begin VB.Label lblDSR 
            AutoSize        =   -1  'True
            Caption         =   "DSR"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   17
            Top             =   480
            Width           =   315
         End
         Begin VB.Shape shpCD 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   765
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpCDOff 
            FillColor       =   &H00008080&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lblCD 
            AutoSize        =   -1  'True
            Caption         =   "CD"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   16
            Top             =   720
            Width           =   210
         End
      End
      Begin VB.Frame fraLineControl 
         Caption         =   "Line Toggle"
         Height          =   1695
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   7680
         Width           =   1095
         Begin VB.CommandButton cmdBreak 
            Caption         =   "BREAK"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdRTS 
            Caption         =   "RTS"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdDTR 
            Caption         =   "DTR"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   735
         End
         Begin VB.Shape shpBreak 
            FillColor       =   &H00FF00FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   840
            Top             =   240
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape shpRTS 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   840
            Top             =   720
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape shpDTR 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   840
            Top             =   1200
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Shape shpBreakOff 
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   840
            Top             =   240
            Width           =   135
         End
         Begin VB.Shape shpRTSOff 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   840
            Top             =   720
            Width           =   135
         End
         Begin VB.Shape shpDTROff 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   840
            Top             =   1200
            Width           =   135
         End
      End
      Begin VB.Frame fraRThreshold 
         Caption         =   "Receive Threshold"
         Height          =   615
         Index           =   0
         Left            =   1320
         TabIndex        =   7
         Top             =   7080
         Width           =   2535
         Begin VB.CommandButton cmdApplyRThresh 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   10
            Top             =   225
            Width           =   735
         End
         Begin VB.TextBox txtRThreshold 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   9
            Top             =   225
            Width           =   615
         End
         Begin VB.CheckBox chkRThreshold 
            Caption         =   "Enable"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Frame fraOutput 
         Caption         =   "Output"
         Height          =   855
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   8520
         Width           =   975
         Begin VB.OptionButton optBinaryOut 
            Caption         =   "Binary"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optStringOut 
            Caption         =   "String"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ListBox lstData 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Index           =   0
         ItemData        =   "frmMain.frx":23486
         Left            =   120
         List            =   "frmMain.frx":23488
         TabIndex        =   2
         Top             =   2400
         Width           =   6495
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "Receive"
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   7800
         Width           =   975
      End
      Begin MSCommLib.MSComm comPort 
         Index           =   0
         Left            =   6120
         Top             =   10200
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   0   'False
         ParityReplace   =   35
         BaudRate        =   1200
         InputMode       =   1
      End
      Begin VB.Timer tmrPoll 
         Enabled         =   0   'False
         Index           =   0
         Left            =   5640
         Top             =   10320
      End
      Begin VB.Frame fraReceive 
         Caption         =   "Data Received"
         Height          =   4455
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   47
         Left            =   6480
         TabIndex        =   82
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   46
         Left            =   6330
         TabIndex        =   81
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   45
         Left            =   6195
         TabIndex        =   80
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   44
         Left            =   6060
         TabIndex        =   79
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   43
         Left            =   5925
         TabIndex        =   78
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   42
         Left            =   5790
         TabIndex        =   77
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   41
         Left            =   5655
         TabIndex        =   76
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   40
         Left            =   5520
         TabIndex        =   75
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   39
         Left            =   5385
         TabIndex        =   74
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   38
         Left            =   5250
         TabIndex        =   73
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   37
         Left            =   5115
         TabIndex        =   72
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   36
         Left            =   4980
         TabIndex        =   71
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   35
         Left            =   4845
         TabIndex        =   70
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   34
         Left            =   4710
         TabIndex        =   69
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   33
         Left            =   4575
         TabIndex        =   68
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   32
         Left            =   4440
         TabIndex        =   67
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   31
         Left            =   4305
         TabIndex        =   66
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   30
         Left            =   4170
         TabIndex        =   65
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   29
         Left            =   4035
         TabIndex        =   64
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   28
         Left            =   3900
         TabIndex        =   63
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   27
         Left            =   3765
         TabIndex        =   62
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   26
         Left            =   3630
         TabIndex        =   61
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   25
         Left            =   3495
         TabIndex        =   60
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   3360
         TabIndex        =   59
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   23
         Left            =   3225
         TabIndex        =   58
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   22
         Left            =   3090
         TabIndex        =   57
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   21
         Left            =   2955
         TabIndex        =   56
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   2820
         TabIndex        =   55
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   19
         Left            =   2685
         TabIndex        =   54
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   2550
         TabIndex        =   53
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   2415
         TabIndex        =   52
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   2280
         TabIndex        =   51
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   2145
         TabIndex        =   50
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   2010
         TabIndex        =   49
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   1875
         TabIndex        =   48
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   1740
         TabIndex        =   47
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   1605
         TabIndex        =   46
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   1470
         TabIndex        =   45
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   1335
         TabIndex        =   44
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   1200
         TabIndex        =   43
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   1065
         TabIndex        =   42
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   930
         TabIndex        =   41
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   795
         TabIndex        =   40
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   660
         TabIndex        =   39
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   525
         TabIndex        =   38
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   390
         TabIndex        =   37
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   255
         TabIndex        =   36
         Top             =   6840
         Width           =   135
      End
      Begin VB.Label lblScope 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   6840
         Width           =   135
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScope 
      Caption         =   "&Scope"
      Begin VB.Menu mnuScopeOn 
         Caption         =   "Scope On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuScp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScopeColor 
         Caption         =   "Blue"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuScopeColor 
         Caption         =   "Green"
         Index           =   1
      End
      Begin VB.Menu mnuScopeColor 
         Caption         =   "Purple"
         Index           =   2
      End
      Begin VB.Menu mnuScopeColor 
         Caption         =   "Orange"
         Index           =   3
      End
      Begin VB.Menu mnuScopeColor 
         Caption         =   "Yellow"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   File:
'       frmMain.Frm
'   Author:
'       Tom DeWitt
'   Description:
'       This is a preliminary serial port test utility. The scope of the project has drifted from a simple test app
'   to a full utility. I am releasing the code as a beta test and to attempt to get some feed back. I understand that
'   the structure of the design is not "robust" as this project's scope has wandered. My primary concern is to
'   determine the functionallity likes and dislikes from other programmers needs. The final version will include
'   Text file logging of data sent and received. Please feel free to make suggestions for improvements. The basic
'   functionality of the code has been tested on both Win NT and Win2000. Note that all the error handling is not
'   complete nor has the app been tested by anyone but me.
'-----------------------------------------------------------------------------------------------------------------------
'   Revisions:
'       Original 4/24/2002
'-----------------------------------------------------------------------------------------------------------------------
'   Functions And Subroutines:
'
'-----------------------------------------------------------------------------------------------------------------------
'   Properties:
'-----------------------------------------------------------------------------------------------------------------------
'   Required Functions,Subroutines,Properties,Variables,Etc.:
'
'-----------------------------------------------------------------------------------------------------------------------
'   Variables:
'       Public:
'
'-----------------------------------------------------------------------------------------------------------------------
'       Private:
Private Type PortSettings
    Baud As String
    Parity As String
    DataBits As String
    StopBits As String
End Type

Private Type ScopeColor
    Picture As String
    Trace As Long
End Type

Private typSettings() As PortSettings
Private typDisplayColor() As ScopeColor
Private iPollCount() As Long
Private iMultiSend() As Long
Private bMultiple() As Boolean
Private bFormLoaded As Boolean
Private bScopeOn As Boolean
Private Ports() As Variant
Private sSubKey As String
Private sKeyValue As String
Private sSettings As String
Private sPortNum As String
Private hnd As Long
'-----------------------------------------------------------------------------------------------------------------------
'   Special Notes:
'       1.  Index Key: Index 0 - Port A; Index 1 - Port B  This app was written using control arrays so if another port
'   is desired a simple copy/paste of the fraPort (main frame) will add another Com Port. The lblScope control array is
'   the only exception. This array is all one control array and uses the index property to calculate the offset that is
'   needed to label the scope trace. Note that if you add another Com Port to make sure that the lblScope indexes that
'   were added have the lowest index to the left and increase going to the right (ie 96 - 143) or the display will be
'   inaccurate. Also insure the proper Z-order of the controls as certain objects, mainly the LED type indicators use a
'   back color and the real indicators visible property is set to false.
'       2.  The scope display is a simulation of the data received from the the Com Port as it would be displayed on an
'   Oscilloscope. The main points to remember is that a negative voltage is considered to be a logical true or on bit,
'   and a positive voltage is a logical false or off bit. Also remember that the left side of the scope display is the
'   first data the scope receives therefore the 0 bit is the left most data bit on the scope display. The lblScope will
'   indicate the 'start bit' with a "S" (Upper case) and 'stop bits' will be indicated by "s" (Lower case). The Port
'   settings affect the number of bits displayed and the scope will only display the last 4 bytes of data the Com Port
'   has recieved.
'
'-----------------------------------------------------------------------------------------------------------------------
'   Constants:
'       Private:
Private Const iHigh As Long = 9                             'Scope Trace High Line Value
Private Const iPW As Long = 9                               'Scope Trace Pulse Width
Private Const iDelta As Long = 36                           'Scope Trace Difference from High To Low Line
Private Const sStartBitLabel As String = "S"                'Scope Trace Label For The Start Bit
Private Const sStopBitLabel As String = "s"                 'Scope Trace Label For The Stop Bit
Private Const sParityBitLabel As String = "p"               'Scope Trace Label For The Parity Bit
Private Const lMainKey As Long = HKEY_LOCAL_MACHINE
Private Const lLength As Long = 1024
Private Const sSettingsKey As String = "Settings"           'Registry Key Name For The Port Settings
Private Const sPortKey As String = "Port"                   'Registry Key Name For The Port Number

'-----------------------------------------------------------------------------------------------------------------------
'   Enumeration Constants:
'-----------------------------------------------------------------------------------------------------------------------
'Initialize The Com Port Display With The Proper Values
'---------------------------------------------------------START---------------------------------------------------------
Private Sub Form_Initialize()
    Dim sPortSet As String
    Dim sTmp As String
    Dim sLabel As String
    Dim iComma1 As Long
    Dim iComma2 As Long
    Dim iComma3 As Long
    Dim com As MSComm
    Dim iX As Long
    Dim iC As Long

        VerifyPorts                                         'Parse Registry Entries For System Serial Com Ports
        For Each com In comPort
            ReDim Preserve typSettings(iX)
            ReDim Preserve iPollCount(iX)                   'Polling Timer Interval Array For Each Port
            ReDim Preserve iMultiSend(iX)                   'Polling Iterations Array For Each Port
            ReDim Preserve bMultiple(iX)                    'Multiple Send Mode Flag Array For Each Port
            ReDim Preserve typDisplayColor(iX)              'Scope Trace Color Array For Each Port
            VerifySettings iX                               'Check Registry For Previous Settings
            cmbParityReplace(iX).Text = com.ParityReplace   'Set the Parity Replace Character
            cmbPort(iX).Text = com.CommPort                 'Set The Port Combobox Text
                With typDisplayColor(iX)                    'Set The Scope Trace To Blue
                    .Picture = App.Path & "\ScopeGridBlue.bmp"
                    .Trace = RGB(14, 181, 228)
                    picScope(iX).Picture = LoadPicture(.Picture)
                End With
                sPortSet = comPort(iX).Settings             'Get the port settings
                iComma1 = InStr(1, sPortSet, ",", vbBinaryCompare)
                iComma2 = InStr(iComma1 + 1, sPortSet, ",", vbBinaryCompare)
                iComma3 = InStr(iComma2 + 1, sPortSet, ",", vbBinaryCompare)
                With typSettings(iX)                        'Break Out The Settings From The String
                    .Baud = Mid$(sPortSet, 1, iComma1 - 1)
                    cmbBaudRate(iX).Text = .Baud
                    .Parity = Mid$(sPortSet, iComma1 + 1, 1)
                    Select Case .Parity
                        Case "n"
                            sTmp = "None"
                        Case "o"
                            sTmp = "Odd"
                        Case "e"
                            sTmp = "Even"
                        Case "s"
                            sTmp = "Space"
                        Case "m"
                            sTmp = "Mark"
                    End Select
                    cmbParity(iX).Text = sTmp
                    .DataBits = Mid$(sPortSet, iComma2 + 1, 1)
                    cmbDataBits(iX).Text = .DataBits
                    .StopBits = Mid$(sPortSet, iComma3 + 1)
                    cmbStopBits(iX).Text = .StopBits
                End With
                cmbHandShake(iX).ListIndex = com.Handshaking
                If com.InputMode = comInputModeText Then
                    optString(iX).Value = True
                Else
                    optBinary(iX).Value = True
                End If
                txtRThreshold(iX).Text = com.RThreshold
                For iC = 0 To UBound(Ports)                 'Add Available Ports To The Ports Combo Box
                    cmbPort(iX).AddItem Ports(iC)
                Next
                For iC = 33 To 255                          'Add Characters To The Parity Replace Combo Box
                    cmbParityReplace(iX).AddItem Chr$(iC)
                Next
                SettingsCaption CInt(iX)
                iX = iX + 1
        Next
        bScopeOn = True
        bFormLoaded = True
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Shut Down The App
'---------------------------------------------------------START---------------------------------------------------------
Private Sub mnuExit_Click()
    Dim com As MSComm

        For Each com In comPort
            If com.PortOpen Then com.PortOpen = False
        Next
    Unload Me
    Set frmMain = Nothing
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Manually Open The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdOpen_Click(Index As Integer)

On Error GoTo ErrHndl

        comPort(Index).PortOpen = True
        fraPort(Index).Caption = "Communication Port " & Chr$(Index + 65) & " Is Open"
        cmdOpen(Index).Enabled = False
        cmdClose(Index).Enabled = True
        cmdShowSend(Index).Enabled = True
    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Manually Close The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdClose_Click(Index As Integer)

On Error GoTo ErrHndl

        comPort(Index).PortOpen = False
        fraPort(Index).Caption = "Communication Port " & Chr$(Index + 65) & " Is Closed"
        shpRTS(Index).Visible = False
        shpDTR(Index).Visible = False
        shpCTS(Index).Visible = False
        shpDSR(Index).Visible = False
        shpCD(Index).Visible = False
        shpBreak(Index).Visible = False
        cmdOpen(Index).Enabled = True
        cmdClose(Index).Enabled = False
        cmdShowSend(Index).Enabled = False

    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Manually Read The Com Port Data
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdReceive_Click(Index As Integer)

On Error GoTo ErrHndl

        If optStringOut(Index).Value Then                   'Check Output Display mode
            lstData(Index).AddItem comPort(Index).Input
        Else
            ReadDataBits Index
        End If

    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Manually Send Data To The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdSend_Click(Index As Integer)

On Error GoTo ErrHndl

        comPort(Index).Output = txtSend(Index).Text

    Exit Sub
ErrHndl:
    Select Case Err.Number
        Case comPortNotOpen
            MsgBox "The Com Port Is Not Open. Open The Port And Then Retry.", vbOKOnly, Err.Description
        Case Else
            Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
    End Select
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Nine Subroutines Control Automatic Data Send To The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbMultiSend_Click(Index As Integer)            'Set the Multiple Send Mode
        Select Case cmbMultiSend(Index).ListIndex
            Case 0                                          'Single
                txtMultiple(Index).Enabled = False
                txtSendTime(Index).Enabled = False
                txtMultiple(Index).Text = "0"
                txtSendTime(Index).Text = "0"
                cmdMultiSend(Index).Caption = "Start"
                cmdMultiSend(Index).Enabled = False
                bMultiple(Index) = False
            Case 1                                          'Multiple
                txtMultiple(Index).Enabled = True
                txtSendTime(Index).Enabled = True
                cmdMultiSend(Index).Enabled = True
                txtMultiple(Index).Text = "1"
                txtSendTime(Index).Text = "1"
                bMultiple(Index) = True
            Case 2                                          'Continuous
                txtMultiple(Index).Enabled = False
                txtSendTime(Index).Enabled = True
                txtMultiple(Index).Text = "0"
                txtSendTime(Index).Text = "1"
                cmdMultiSend(Index).Enabled = True
                bMultiple(Index) = False
        End Select
End Sub

Private Sub cmdMultiSend_Click(Index As Integer)
        AutoSendComPort Index
End Sub

Private Sub tmrSend_Timer(Index As Integer)

On Error GoTo ErrHndl

        comPort(Index).Output = txtSend(Index).Text
        If bMultiple(Index) Then                            'Check if Multiple Send Is Enabled
            iMultiSend(Index) = iMultiSend(Index) - 1       'Countdown iterations
            If iMultiSend(Index) = 0 Then                   'Last iteration, Stop polling
                AutoSendComPort Index
            End If
        End If

    Exit Sub
ErrHndl:
    Select Case Err.Number
        Case comPortNotOpen
            MsgBox "The Com Port Is Not Open. Open The Port And Then Retry.", vbOKOnly, Err.Description
        Case Else
            Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
    End Select
End Sub

Private Sub txtMultiple_Change(Index As Integer)
    Dim sSend As String

    sSend = txtMultiple(Index).Text                         'Iterations
    If cmbMultiSend(Index).ListIndex = 1 Then               'Multiple Send Option
        If sSend = vbNullString Then                        'Empty String is Not Allowed
            cmdMultiSend(Index).Enabled = False
        Else
            If CInt(sSend) > 0 Then                         'Iterations Must Be Greater than 0
                cmdMultiSend(Index).Enabled = True
            Else
                cmdMultiSend(Index).Enabled = False
            End If
        End If
    End If
End Sub

Private Sub txtSendTime_Change(Index As Integer)
    Dim sSend As String
    Dim sTime As String

        sSend = txtMultiple(Index).Text                     'Iterations
        sTime = txtSendTime(Index).Text                     'Timer Interval
        Select Case cmbMultiSend(Index).ListIndex
            Case 1                                          'Multiple Must Have Iterations and Timer Interval Values
                If sSend = vbNullString Or sTime = vbNullString Then
                    cmdMultiSend(Index).Enabled = False
                Else                                        'Iterations and Timer Interval must be Greater Than 0
                    If CInt(sSend) > 0 And CInt(sTime) > 0 Then
                        cmdMultiSend(Index).Enabled = True
                    Else
                        cmdMultiSend(Index).Enabled = False
                    End If
                End If
            Case 2                                          'Continuous Must Have Timer Interval Value
                If sTime = vbNullString Then
                    cmdMultiSend(Index).Enabled = False
                Else
                    If CInt(sTime) > 0 Then                 'Timer Interval must be Greater Than 0
                        cmdMultiSend(Index).Enabled = True
                    Else
                        cmdMultiSend(Index).Enabled = False
                    End If
                End If
            End Select
End Sub

Private Sub AutoSendComPort(Index As Integer)

        If tmrSend(Index).Enabled Then                      'Timer Enabled, Stop polling
            cmdMultiSend(Index).Caption = "Start"
            cmbMultiSend(Index).Enabled = True
            If cmbMultiSend(Index).ListIndex = 1 Then       'Multiple Send Option
                txtMultiple(Index).Enabled = True
                bMultiple(Index) = False
            End If
            txtSendTime(Index).Enabled = True
            tmrSend(Index).Enabled = False
        Else                                                'Timer Not Enabled, Start polling
            cmdMultiSend(Index).Caption = "Stop"
            cmbMultiSend(Index).Enabled = False
            txtMultiple(Index).Enabled = False
            txtSendTime(Index).Enabled = False
            If cmbMultiSend(Index).ListIndex = 1 Then       'Multiple Send Option
                iMultiSend(Index) = CLng(txtMultiple(Index).Text)
                txtMultiple(Index).Enabled = False
                bMultiple(Index) = True
            End If
            tmrSend(Index).Interval = CInt(txtSendTime(Index).Text)
            tmrSend(Index).Enabled = True
        End If
End Sub

Private Sub txtSendTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then          'Only Allow Numbers
        KeyAscii = 0
    End If
End Sub

Private Sub cmbMultiSend_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub

Private Sub txtMultiple_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then          'Only Allow Numbers
        KeyAscii = 0
    End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Menu Items
'---------------------------------------------------------START---------------------------------------------------------
Private Sub mnuScopeOn_Click()
    Dim mnu As Menu

        bScopeOn = Not bScopeOn
        mnuScopeOn.Checked = bScopeOn
        For Each mnu In mnuScopeColor
            mnu.Enabled = bScopeOn
        Next

End Sub

Private Sub mnuScopeColor_Click(Index As Integer)
    Dim mnu As Menu
    Dim iX As Long
    Dim iY As Long

        iX = comPort.Count - 1                              'Get The Number of Ports in the Control Array
        For Each mnu In mnuScopeColor                       'Uncheck All The Color Menu Options
            mnu.Checked = False
        Next
        For iY = 0 To iX
            Select Case Index
                Case 0                                      'Blue
                    typDisplayColor(iY).Picture = App.Path & "\ScopeGridBlue.bmp"
                    typDisplayColor(iY).Trace = RGB(14, 181, 228)
                    mnuScopeColor(Index).Checked = True
                Case 1                                      'Green
                    typDisplayColor(iY).Picture = App.Path & "\ScopeGridGreen.bmp"
                    typDisplayColor(iY).Trace = RGB(50, 173, 44)
                    mnuScopeColor(Index).Checked = True
                Case 2                                      'Purple
                    typDisplayColor(iY).Picture = App.Path & "\ScopeGridPurple.bmp"
                    typDisplayColor(iY).Trace = RGB(255, 50, 255)
                    mnuScopeColor(Index).Checked = True
                Case 3                                      'Orange
                    typDisplayColor(iY).Picture = App.Path & "\ScopeGridOrange.bmp"
                    typDisplayColor(iY).Trace = RGB(255, 183, 111)
                    mnuScopeColor(Index).Checked = True
                Case 4                                      'Yellow
                    typDisplayColor(iY).Picture = App.Path & "\ScopeGridYellow.bmp"
                    typDisplayColor(iY).Trace = RGB(255, 255, 0)
                    mnuScopeColor(Index).Checked = True
            End Select
'Set the scope trace background Picture to the Selected Color
            picScope(iY).Picture = LoadPicture(typDisplayColor(iY).Picture)
        Next
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Monitor The Com Port Events
'---------------------------------------------------------START---------------------------------------------------------
Private Sub comPort_OnComm(Index As Integer)

        With comPort(Index)
            Select Case .CommEvent
                Case comEventBreak
                    shpBreakEvent(Index).Visible = True
                    tmrLineDetect(Index).Enabled = True
                Case comEventCTSTO
                    shpCTSTO(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "CTS Time Out"
                Case comEventDSRTO
                    shpDSRTO(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "DRS Time Out"
                Case comEventFrame
                    shpFrame(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "Framing Error"
                Case comEventOverrun
                    shpOverRun(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "Port Overrun"
                Case comEventCDTO
                    shpCDTO(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "CD Time Out"
                Case comEventRxOver
                    shpRXOver(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "RX Buffer Oferflow"
                Case comEventRxParity
                    shpParity(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "Parity Error"
                Case comEventTxFull
                    shpTXFull(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "TX Buffer Full"
                Case comEventDCB
                    shpDCB(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(0).AddItem "Unexpected DCB Error"
                Case comEvSend
                    shpSThresh(Index).Visible = True
                    tmrErrors(Index).Enabled = True
                    lstErrors(Index).AddItem "Too Few Characters"
                Case comEvReceive
                    ReadDataBits Index
                Case comEvCTS
                    shpCTS(Index).Visible = .CTSHolding
                Case comEvDSR
                    shpDSR(Index).Visible = .DSRHolding
                Case comEvCD
                    shpCD(Index).Visible = .CDHolding
                Case comEvRing
                    shpRing(Index).Visible = True
                    tmrLineDetect(Index).Enabled = True
                Case comEvEOF
                    shpEOF(Index).Visible = True
                    tmrLineDetect(Index).Enabled = True
                Case Else
                    Debug.Print .CommEvent & " " & CStr(Index)
            End Select
        End With
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Change the Parity Replace Character Property
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbParityReplace_Click(Index As Integer)
    comPort(Index).ParityReplace = cmbParityReplace(Index).Text
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Toggle the DTR Property On and Off
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdDTR_Click(Index As Integer)
        comPort(Index).DTREnable = Not comPort(Index).DTREnable
        shpDTR(Index).Visible = comPort(Index).DTREnable
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Toggle the RTS Property On and Off
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdRTS_Click(Index As Integer)
        comPort(Index).RTSEnable = Not comPort(Index).RTSEnable
        shpRTS(Index).Visible = comPort(Index).RTSEnable
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Toggle the Break Property On and Off
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdBreak_Click(Index As Integer)
        comPort(Index).Break = Not comPort(Index).Break
        shpBreak(Index).Visible = comPort(Index).Break
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Clear The Data Display Text
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdClear_Click(Index As Integer)
    Dim LabelOffSet As Long
    Dim iZ As Long

        lstData(Index).Clear
        lstErrors(Index).Clear
        LabelOffSet = Index * 48
        For iZ = LabelOffSet To LabelOffSet + 47
            lblScope(iZ).Caption = "-"
        Next
        picScope(Index).Refresh
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Toggle Between The Settings And Send Option Boxes
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdSettings_Click(Index As Integer)
    fraSend(Index).Visible = False
End Sub

Private Sub cmdShowSend_Click(Index As Integer)
    fraSend(Index).Visible = True
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Handle the Com Port Number
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbPort_Click(Index As Integer)
        UpdatePortSettings Index                            'Update The Port Settings
End Sub

Private Sub cmbPort_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Handle the Baud Rate Settings
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbBaudRate_Click(Index As Integer)
        typSettings(Index).Baud = cmbBaudRate(Index).Text
        UpdatePortSettings Index                            'Update The Port Settings
End Sub

Private Sub cmbBaudRate_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Handle the Data Bits Settings
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbDataBits_Click(Index As Integer)
        typSettings(Index).DataBits = cmbDataBits(Index).Text
        UpdatePortSettings Index                            'Update The Port Settings
End Sub

Private Sub cmbDataBits_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Handle the Parity Settings
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbParity_Click(Index As Integer)
        typSettings(Index).Parity = LCase$(Left$(cmbParity(Index).Text, 1))
        UpdatePortSettings Index                            'Update The Port Settings
End Sub

Private Sub cmbParity_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Handle the Stop Bits Settings
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbStopBits_Click(Index As Integer)
        typSettings(Index).StopBits = cmbStopBits(Index).Text
        UpdatePortSettings Index                            'Update The Port Settings
End Sub

Private Sub cmbStopBits_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Handle The Hand Shaking Property For The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmbHandShake_Click(Index As Integer)
    UpdateHandShaking Index, cmbHandShake(Index).ListIndex
End Sub

Private Sub cmbHandShake_KeyPress(Index As Integer, KeyAscii As Integer)
        KeyAscii = 0                                        'Prevent Data Input
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Two Subroutines Control The Display Of The Data Read From The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub optBinary_Click(Index As Integer)
        If bFormLoaded Then
            UpdateInputMode Index, comInputModeBinary       'Update The Input Mode (Binary)
        End If
End Sub

Private Sub optString_Click(Index As Integer)
        If bFormLoaded Then
            UpdateInputMode Index, comInputModeText         'Update The Input Mode (Text)
        End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Three Subroutines Handle Adding Special Keys To The Send String
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdAscii_Click(Index As Integer)
    Dim iChr As Integer

        txtSend(Index).SetFocus
        iChr = CInt(txtAscii(Index).Text)
        If iChr <= 255 Then                                 'Check For Valid Character Code
            SendKeys Chr(iChr)
        Else
            MsgBox "Character Code Must Be 255 or Less", vbInformation, "Invalid Character Code !"
        End If
        txtAscii(Index).Text = vbNullString
        cmdAscii(Index).Enabled = False
End Sub

Private Sub txtAscii_Change(Index As Integer)
    If txtAscii(Index).Text = vbNullString Then             'Disable The Command Button If Ascii Text Is Empty
        cmdAscii(Index).Enabled = False
    Else
        cmdAscii(Index).Enabled = True
    End If
End Sub

Private Sub txtAscii_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then          'Only Allow Numbers
        KeyAscii = 0
    End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Four Subroutines Handle The Receive Threshold Settings For The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdApplyRThresh_Click(Index As Integer)
    If txtRThreshold(Index).Text = vbNullString Then
        MsgBox "The Receive Threshold Property Must Have A Value.", vbInformation, "Invalid Data Entry !"
        Exit Sub
    End If
    cmdApplyRThresh(Index).Enabled = False
    UpdateRThreshold Index, CInt(txtRThreshold(Index).Text) 'Update The RThreshold Setting
End Sub

Private Sub chkRThreshold_Click(Index As Integer)

        txtRThreshold(Index).Enabled = chkRThreshold(Index).Value
        If chkRThreshold(Index).Value = vbChecked Then      'Zero The Property If Disabled
            chkPoll(Index).Value = vbUnchecked
        Else
            bFormLoaded = False                             'Prevent The Text Change Event
            txtRThreshold(Index).Text = 0
            UpdateRThreshold Index, 0
            bFormLoaded = True
        End If
End Sub

Private Sub txtRThreshold_Change(Index As Integer)
        If bFormLoaded Then
            cmdApplyRThresh(Index).Enabled = True
        End If
End Sub

Private Sub txtRThreshold_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then          'Only Allow Numbers
        KeyAscii = 0
    End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Four Subroutines Handle The Input Length Property Of The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub optCount_Click(Index As Integer)
    txtCount(Index).Enabled = True
End Sub

Private Sub optAll_Click(Index As Integer)
    comPort(Index).InputLen = 0
    txtCount(Index).Text = 0
    txtCount(Index).Enabled = False
End Sub

Private Sub txtCount_Change(Index As Integer)
        If bFormLoaded Then
            If Not IsNumeric(txtCount(Index).Text) Then
                MsgBox "The Input Length Property Must Have A Numeric Value.", vbInformation, "Invalid Data Entry !"
                Exit Sub
            End If
            comPort(Index).InputLen = CInt(txtCount(Index).Text)
        End If
End Sub

Private Sub txtCount_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then          'Only Allow Numbers
        KeyAscii = 0
    End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'These Next Five Subroutines Handle The Continous Polling Of The Com Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub cmdApplyPoll_Click(Index As Integer)
    If Not IsNumeric(txtPoll(Index).Text) Then
        MsgBox "The Poll Timer Interval Property Must Have A Numeric Value.", vbInformation, "Invalid Data Entry !"
        Exit Sub
    End If
    If Not IsNumeric(txtBuffer(Index).Text) Then
        MsgBox "The Poll Buffer Property Must Have A Numeric Value.", vbInformation, "Invalid Data Entry !"
        Exit Sub
    End If
    tmrPoll(Index).Interval = CInt(txtPoll(Index).Text)
    tmrPoll(Index).Enabled = True
    iPollCount(Index) = CLng(txtBuffer(Index).Text)
    cmdApplyPoll(Index).Enabled = False
End Sub

Private Sub chkPoll_Click(Index As Integer)

        txtPoll(Index).Enabled = chkPoll(Index).Value
        txtBuffer(Index).Enabled = chkPoll(Index).Value
        If chkPoll(Index).Value = vbChecked Then            'Polling Selected
            chkRThreshold(Index).Value = vbUnchecked        'Deselect Receive Threshold
        Else                                                'Polling Deselected
            bFormLoaded = False                             'Prevent The Text Change Event
            txtPoll(Index).Text = 0
            txtBuffer(Index).Text = 0
            tmrPoll(Index).Enabled = False
            tmrPoll(Index).Interval = 0
            bFormLoaded = True
        End If

End Sub

Private Sub txtPoll_Change(Index As Integer)
        If bFormLoaded Then
            cmdApplyPoll(Index).Enabled = True
        End If
End Sub

Private Sub txtPoll_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then          'Only Allow Numbers
        KeyAscii = 0
    End If
End Sub

Private Sub tmrPoll_Timer(Index As Integer)
    If comPort(Index).InBufferCount >= iPollCount(Index) Then
        ReadDataBits Index
    End If
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Subroutine To Update The Com Port Settings Property
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdatePortSettings(ByVal Index As Integer)
    Dim sPortSet As String

On Error GoTo ErrHndl

        With typSettings(Index)
            sPortSet = .Baud & ","
            sPortSet = sPortSet & .Parity & ","
            sPortSet = sPortSet & .DataBits & ","
            sPortSet = sPortSet & .StopBits
        End With
        If comPort(Index).PortOpen Then
            comPort(Index).PortOpen = False
            DoEvents
            comPort(Index).Settings = sPortSet
            comPort(Index).CommPort = cmbPort(Index).Text
            comPort(Index).PortOpen = True
        Else
            comPort(Index).Settings = sPortSet
            comPort(Index).CommPort = cmbPort(Index).Text
        End If
        SettingsCaption Index
        UpdateSettings CLng(Index)
    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Subroutine To Update The Com Port Input Mode Property
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdateInputMode(ByVal Index As Integer, ByVal Mode As InputModeConstants)

On Error GoTo ErrHndl

        If comPort(Index).PortOpen Then
            comPort(Index).PortOpen = False
            DoEvents
            comPort(Index).InputMode = Mode
            comPort(Index).PortOpen = True
        Else
            comPort(Index).InputMode = Mode
        End If
        SettingsCaption Index
    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Subroutine To Update The Com Port Receive Threshold Setting Property
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdateRThreshold(ByVal Index As Integer, ByVal Value As Integer)

On Error GoTo ErrHndl

        If comPort(Index).PortOpen Then
            comPort(Index).PortOpen = False
            DoEvents
            comPort(Index).RThreshold = Value
            comPort(Index).PortOpen = True
        Else
            comPort(Index).RThreshold = Value
        End If

    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Subroutine To Update The Com Port HandShaking Property
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdateHandShaking(ByVal Index As Integer, ByVal Value As Integer)

On Error GoTo ErrHndl

        If comPort(Index).PortOpen Then
            comPort(Index).PortOpen = False
            DoEvents
            comPort(Index).Handshaking = Value
            comPort(Index).PortOpen = True
        Else
            comPort(Index).Handshaking = Value
        End If
        SettingsCaption Index
    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Subroutine To Update The Port Settings Frame Caption
'---------------------------------------------------------START---------------------------------------------------------
Private Sub SettingsCaption(Index As Integer)
    Dim sStr As String

        sStr = "Port "
        With comPort(Index)
            sStr = sStr & .CommPort & " Settings : "
            sStr = sStr & .Settings & ", "
            Select Case .InputMode
                Case comInputModeBinary
                    sStr = sStr & "Binary Input, "
                Case Else
                      sStr = sStr & "String Input, "
            End Select
        End With
        sStr = sStr & "HandShaking : " & cmbHandShake(Index).Text
        fraSettings(Index).Caption = sStr
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Perform Bitwise Check On A Number
'---------------------------------------------------------START---------------------------------------------------------
Private Function BitOn(Number As Long, Bit As Long) As Boolean
        Dim iX As Long
        Dim iY As Long

        iY = 1
        For iX = 1 To Bit - 1
            iY = iY * 2
        Next
        If Number And iY Then BitOn = True Else BitOn = False
End Function
'----------------------------------------------------------END----------------------------------------------------------
'Read The Data From The Com Port And Display As A Binary String per Byte Received
'---------------------------------------------------------START---------------------------------------------------------
Private Sub ReadDataBits(ByVal Index As Integer)
        Dim bytInput() As Byte
        Dim lngScope() As Long
        Dim bytElement As Byte
        Dim sResult As String
        Dim sLead As String
        Dim sDisplay As String
        Dim sData As String
        Dim sSpace As String
        Dim iBit As Long
        Dim iX As Long
        Dim iY As Long
        Dim iZ As Long
        Dim iB As Long

On Error GoTo ErrHndl

        If comPort(Index).InBufferCount = 0 Then            'Check The Buffer For Data
            lstData(Index).AddItem "There is No Data In The Port Buffer"
            Exit Sub                                        'Exit If The Buffer Is Empty
        End If
        bytInput = comPort(Index).Input
        iX = UBound(bytInput)                               'Get the number of byte elements in the input
        If iX < 3 Then                                      'Set up Array for scope display
            ReDim lngScope(iX)
        Else
            ReDim lngScope(3)
        End If
        For iY = 0 To iX
            bytElement = bytInput(iY)                       'Get the next byte element from the input
            If iY >= iX - 3 Then                            'Get up to the last 4 elements of the input for scope
                lngScope(iZ) = bytInput(iY)
                iZ = iZ + 1
            End If
            sData = Chr$(bytElement)
            If Asc(sData) = 0 Then                          'Check for Null
                sData = Chr(13)
            End If
            sLead = "(" & sData & ")"                       'Create display lead
            For iBit = 1 To 8                               'Bitwise data
                Select Case iBit
                    Case 4
                        sSpace = ","                        'Delineate at the 4th bit
                    Case Else
                        sSpace = vbNullString
                End Select
                sResult = sSpace & Abs(CInt(BitOn(CLng(bytElement), iBit))) & sResult
            Next
            iB = iB + 1                                     'Count the bytes, 4 per display line
            sDisplay = sDisplay & sLead & sResult & " "
            sResult = vbNullString
            If iB = 4 Then                                  'On 4th byte display the data
                lstData(Index).AddItem sDisplay
                sDisplay = vbNullString
                iB = 0
            End If
        Next

        If sDisplay <> vbNullString Then                    'Display remaining data
            lstData(Index).AddItem sDisplay
        End If

        If bScopeOn Then
            ScopeDisplay lngScope, Index                    'Send The Data To The Scope Display
        End If

    Exit Sub
ErrHndl:
    Select Case Err.Number
        Case comPortNotOpen
            MsgBox "The Com Port Is Not Open. Open The Port And Then Retry.", vbOKOnly, Err.Description
        Case Else
            Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
    End Select
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Generate a Flash for the Break, EOF, and Ring Indicator
'---------------------------------------------------------START---------------------------------------------------------
Private Sub tmrLineDetect_Timer(Index As Integer)
    tmrLineDetect(Index).Enabled = False
    shpBreakEvent(Index).Visible = False
    shpRing(Index).Visible = False
    shpEOF(Index).Visible = False
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Generate a Flash for the Error Indicators
'---------------------------------------------------------START---------------------------------------------------------
Private Sub tmrErrors_Timer(Index As Integer)
        tmrErrors(Index).Enabled = False
        shpCTSTO(Index).Visible = False
        shpDSRTO(Index).Visible = False
        shpFrame(Index).Visible = False
        shpOverRun(Index).Visible = False
        shpCDTO(Index).Visible = False
        shpRXOver(Index).Visible = False
        shpParity(Index).Visible = False
        shpTXFull(Index).Visible = False
        shpDCB(Index).Visible = False
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Subroutine To Display The Data As A Scope Trace
'---------------------------------------------------------START---------------------------------------------------------
Private Sub ScopeDisplay(ByRef Data() As Long, Index As Integer)
    Dim iX As Long                                          'Use for the X Coordinate in the Scope Picture
    Dim iY As Long                                          'Use for the Y Coordinate in the Scope Picture
    Dim iZ As Long                                          'Multi use Index Counter
    Dim LabelOffSet As Long
    Dim iLow As Long
    Dim hPic As Long
    Dim iTraceColor As Long
    Dim bOn As Boolean
    Dim iBytes As Long
    Dim bTrace() As Boolean
    Dim iBits As Long
    Dim iStop As Long
    Dim PortSet As PortSettings
    Dim iBitCount As Long
    Dim iOnBits As Long
    Dim bChkParity As Boolean

On Error GoTo ErrHndl

        iLow = iHigh + iDelta                               'Low Scope Trace Value
        LabelOffSet = Index * 48                            'One Control Array For All Ports (offset accordingly)
        picScope(Index).Refresh                             'Clear The Scope Display
        hPic = picScope(Index).hdc                          'Get The Handle of The Scope Display Picture Box
        iBytes = UBound(Data)                               'No Of Bytes In The Data Array To Display
        PortSet = typSettings(0)                            'Get The Ports Settings
        iZ = 0

        With PortSet                                        'Count The Additional Bits From The Ports Settings
            If .Parity <> "n" Then
                iZ = 1
                bChkParity = True
            End If
            iBits = CLng(.DataBits)                         'Number Of Bits To Display From Port Settings
            iStop = CLng(.StopBits)                         'Number Of Stop Bits To Add
            iZ = iZ + iBits + iStop                         'Sum Stop,Parity, and Data Bits
        End With

'Example 2 Stop Bits, No Parity, 7 Data Bits --> Stop + Parity + Data = 9 --> iZ = 9
'Data() contains 3 Bytes --> iBytes = 2 (array index) -->       Note The 2nd (iBytes + 1) Adds The Start Bits
'       iBitCount = (9(iZ) * (2(iBytes) + 1)[No Of Bytes] + (2(iBytes) + 1)[Start Bits] --> iBitCount = 30 bits Total
        iBitCount = (iZ * (iBytes + 1)) + (iBytes + 1)
        ReDim bTrace(iBitCount - 1)                         'Trace Array One less than Total BitCount (array index)

        iZ = 0
        For iX = 0 To iBytes                                'Max iBytes is 3 (Four Byte Trace)
            bTrace(iZ) = True                               'Start Bit
            lblScope(iZ + LabelOffSet).Caption = sStartBitLabel
            iZ = iZ + 1
            iOnBits = 0                                     'Zero Count For On Bits To Check Parity
            For iY = 1 To iBits                             'Parse Each Data Bit
                bOn = BitOn(Data(iX), iY)
                iOnBits = Abs(bOn) + iOnBits
                bTrace(iZ) = Not bOn                        'Invert as Negative Voltage on Scope is Logical True
                lblScope(iZ + LabelOffSet).Caption = CStr(iY - 1)
                iZ = iZ + 1
            Next
'Set The Scope Trace Parity Bits
'Scope Trace is Inverted From Expected Because -3 volts is Logical True, +3 volts is Logical False
            If bChkParity Then                              'True For All But Parity Setting Of "None"
                Select Case PortSet.Parity
                    Case "e"                                'Even Parity
                        If iOnBits Mod 2 Then               'OnBits is Odd Parity Bit Is On
                            bTrace(iZ) = False
                        Else
                            bTrace(iZ) = True               'OnBits is Even Parity Bit Is Off
                        End If
                    Case "o"                                'Odd Parity
                        If iOnBits Mod 2 Then               'OnBits is Odd Parity Bit Is Off
                            bTrace(iZ) = True
                        Else
                            bTrace(iZ) = False              'OnBits is Even Parity Bit Is On
                        End If
                    Case "m"                                'Mark Parity Bit is Always On
                        bTrace(iZ) = False
                    Case "s"
                        bTrace(iZ) = True                   'Space Parity Bit is Always Off
                End Select
                lblScope(iZ + LabelOffSet).Caption = sParityBitLabel
                iZ = iZ + 1
            End If

'Set The Scope Trace Stop Bits
            If iStop = 1 Then                               'One Stop Bit
                bTrace(iZ) = False                          'Stop Bit
                lblScope(iZ + LabelOffSet).Caption = sStopBitLabel
                iZ = iZ + 1
            Else                                            'Two Stop Bits
                bTrace(iZ) = False                          'Stop Bit
                lblScope(iZ + LabelOffSet).Caption = sStopBitLabel
                bTrace(iZ + 1) = False                      'Stop Bit
                lblScope(iZ + LabelOffSet + 1).Caption = sStopBitLabel
                iZ = iZ + 2
            End If
        Next
'Get The Proper Scope Trace To Match The Background Color
        iTraceColor = typDisplayColor(Index).Trace

        iBits = 0                                           'Zero The Bit Count Index
        bOn = bTrace(iBits)                                 'Parse The Trace Array And Draw on Scope Display
        For iX = 1 To picScope(Index).ScaleWidth            'X Coordinate
            iY = iLow - (Abs(bOn) * iDelta)                 'Y Coordinate
            Do While iX Mod iPW                             'Draw Horizontal Line To The Pulse Width
                SetPixel hPic, iX, iY, iTraceColor
                iX = iX + 1
            Loop

            If iBits < UBound(bTrace) Then                  'Check For End Of Trace
                If bTrace(iBits) <> bTrace(iBits + 1) Then  'Check To See If Next Bit Has Changed State
                    For iY = iHigh To iLow                  'Next Bit Change In State Draw Vertical Line
                        SetPixel hPic, iX, iY, iTraceColor
                    Next
                End If
                iBits = iBits + 1                           'Get The Next Bit
                bOn = bTrace(iBits)
            Else                                            'End Of Trace Data Set The Stop Bit
                iX = iX + 1                                 'Skip a Bit for the Stop Bit
                iBits = iBits + 1 + LabelOffSet
                Exit For
            End If
        Next

        Do While iX < picScope(Index).ScaleWidth            'Run The Scope Trace Out
            SetPixel hPic, iX, iLow, iTraceColor
            iX = iX + 1
        Loop
        For iX = iBits To LabelOffSet + 47                  'Set The Remaining Scope Label Captions
            lblScope(iX).Caption = "-"
        Next

    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Open The Local Machine Registry And Get The Serial Ports Available On The Local Machine, Validate Selected Port
'---------------------------------------------------------START---------------------------------------------------------
Private Sub VerifyPorts()
    Dim sPort As String
    Dim iX As Long
    Dim iY As Long
    Dim lngType As Long
    Dim lngValue As Long
    Dim sName As String
    Dim sSwap As String
    ReDim varResult(0 To 1, 0 To 100) As Variant
    Const lNameLen As Long = 260
    Const lDataLen As Long = 4096

        sSubKey = "Hardware\Devicemap\SerialComm"
        If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_READ, hnd) Then Exit Sub
            For iX = 0 To 999999
                If iX > UBound(varResult, 2) Then
                    ReDim Preserve varResult(0 To 1, iX + 99)
                End If
                sName = Space$(lNameLen)
                ReDim binValue(0 To lDataLen - 1) As Byte
                If RegEnumValue(hnd, iX, sName, lNameLen, ByVal 0&, lngType, binValue(0), lDataLen) Then Exit For
                    varResult(0, iX) = Left$(sName, lNameLen)
                    
                    Select Case lngType
                        Case REG_DWORD
                            CopyMemory lngValue, binValue(0), 4
                            varResult(1, iX) = lngValue
                        Case REG_SZ
                            varResult(1, iX) = Left$(StrConv(binValue(), vbUnicode), lDataLen - 1)
                        Case Else
                            ReDim Preserve binValue(0 To lDataLen - 1) As Byte
                            varResult(1, iX) = binValue()
                    End Select
            Next
        If hnd Then RegCloseKey hnd                                             'Close The Registry Key
        ReDim Preserve varResult(0 To 1, iX - 1) As Variant
        ReDim Ports(iX - 1)
        For iX = 0 To UBound(varResult, 2)                                      'Trim 'Port' To Get Just The Number
            sPort = Mid$(varResult(1, iX), 4, 1)
            Ports(iX) = sPort
        Next

        iY = UBound(Ports)                                                       'Arrange The Ports Numbers Low To High
        For iX = 0 To (iY - 1)
            If Ports(iX + 1) < Ports(iX) Then
                sSwap = Ports(iX + 1)
                Ports(iX + 1) = Ports(iX)
                Ports(iX) = sSwap
                iX = -1
            End If
        Next

End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Check The Registy For The Last Used Settings And Sets The MSComm Object Properties. If There Is No Entry It Creates
'One With The Default Setting(Com1 38400n,8,1)
'---------------------------------------------------------START---------------------------------------------------------
Private Sub VerifySettings(Index As Long)
    Dim disposition As Long
    Dim sTmp As String
    Dim iX As Long

        iX = Index
On Error GoTo ErrTrap

        sSettings = comPort(iX).Settings
        sPortNum = comPort(iX).CommPort
        sSubKey = "Software\Damage Inc\ComPort Utility"
        If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_READ, hnd) Then
            If RegCreateKeyEx(lMainKey, sSubKey, 0, 0, 0, 0, 0, hnd, disposition) Then
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Create Registry Key"
            End If
        End If

'The Key Has Been Found/or Created, Now Check To See If Previous Settings Are Present

'Check For The Settings Subkey and Retrieve Value If Present, Then Set ComPort 'Settings' Property

        sKeyValue = Space$(lLength)                                             'Pad The sKeyValue Variable
        If RegQueryValueEx(hnd, sSettingsKey & Chr$(iX + 65), 0, REG_SZ, ByVal sKeyValue, lLength) Then '0 Return OK
            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Settings' Property
                If RegSetValueEx(hnd, sSettingsKey & Chr$(iX + 65), 0, REG_SZ, ByVal sSettings, Len(sSettings)) Then
                    Err.Raise 1001, "VerifySettings() Sub", "Could Not Set Registry Key Settings Value"
                End If
            End If
        Else            'Read Value From Key And Set The Port 'Settings' Property To The Value In The Registry
            comPort(iX).Settings = sKeyValue
        End If

'Check For The Port Subkey and Retrieve Value If Present, Then Set ComPort 'Port' Property

        sKeyValue = Space$(lLength)                                             'Pad The sKeyValue Variable
        If RegQueryValueEx(hnd, sPortKey & Chr$(iX + 65), 0, REG_SZ, ByVal sKeyValue, lLength) Then     '0 Return OK
            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "VerifySettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Port' Property
                If RegSetValueEx(hnd, sPortKey & Chr$(iX + 65), 0, REG_SZ, ByVal sPortNum, Len(sPortNum)) Then
                    Err.Raise 1001, "VerifySettings() Sub", "Could Not Set Registry Key Port Value"
                End If
            End If
        Else            'Read Value From Key And Set The Port 'Port' Property To The Value In The Registry
            comPort(iX).CommPort = sKeyValue
        End If
        RegCloseKey hnd
    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
'Changes The Registry Entries When The User Changes Port Settings
'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdateSettings(Index As Long)
    Dim iX As Long

        iX = Index
On Error GoTo ErrTrap

        sSettings = comPort(iX).Settings
        sPortNum = comPort(iX).CommPort
        sSubKey = "Software\Damage Inc\ComPort Utility"

            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "UpdateSettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Settings' Property
                If RegSetValueEx(hnd, sSettingsKey & Chr$(iX + 65), 0, REG_SZ, ByVal sSettings, Len(sSettings)) Then
                    Err.Raise 1001, "UpdateSettings() Sub", "Could Not Set Registry Key Settings Value"
                End If
            End If

            If RegOpenKeyEx(lMainKey, sSubKey, 0, KEY_WRITE, hnd) Then                      '0 Return if Successful
                Err.Raise 1001, "UpdateSettings() Sub", "Could Not Open Registry Key"
            Else        'The Value Was Not Present, Set To Default Port 'Port' Property
                If RegSetValueEx(hnd, sPortKey & Chr$(iX + 65), 0, REG_SZ, ByVal sPortNum, Len(sPortNum)) Then
                    Err.Raise 1001, "UpdateSettings() Sub", "Could Not Set Registry Key Port Value"
                End If
            End If

    Exit Sub

ErrTrap:
        MsgBox Err.Number & " " & Err.Description & vbCr & " Error Generated By " & Err.Source, vbCritical, _
"System Error Trap !"
End Sub
'----------------------------------------------------------END----------------------------------------------------------
