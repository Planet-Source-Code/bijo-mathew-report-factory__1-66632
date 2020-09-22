VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15270
   HelpContextID   =   1015
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReport.frx":038A
   ScaleHeight     =   8760
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1575
      Left            =   240
      TabIndex        =   45
      Top             =   6120
      Visible         =   0   'False
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      WordWrap        =   -1  'True
      FormatString    =   $"frmReport.frx":3389B
   End
   Begin VB.TextBox txtReportTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      MaxLength       =   50
      TabIndex        =   35
      Text            =   "Report 1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox lstCheck 
      Height          =   450
      Left            =   120
      TabIndex        =   38
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Group By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   240
      Width           =   1380
      Begin VB.ComboBox cboGrpCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboGrpCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin Reports_Factory.ucButtons_H cmdGroupBy 
         Height          =   330
         Left            =   1110
         TabIndex        =   31
         Top             =   -110
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         Focus           =   0   'False
         cGradient       =   16761024
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmReport.frx":339A6
      End
      Begin Reports_Factory.ucButtons_H cmdGrpFilter 
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Filter Data   "
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   14737632
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   3
         Image           =   "frmReport.frx":33CC0
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmReport.frx":3405A
      End
      Begin Reports_Factory.ucButtons_H cmdGrpFilterReset 
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Reset Filter   "
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   14737632
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   3
         Image           =   "frmReport.frx":34374
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmReport.frx":3470E
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   26
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Filter Where"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   1500
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   2760
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cboOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         ItemData        =   "frmReport.frx":34A28
         Left            =   1440
         List            =   "frmReport.frx":34A4A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cboCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2760
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cboOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         ItemData        =   "frmReport.frx":34A93
         Left            =   1440
         List            =   "frmReport.frx":34AB5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2760
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cboOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         ItemData        =   "frmReport.frx":34AFE
         Left            =   1440
         List            =   "frmReport.frx":34B20
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cboCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "frmReport.frx":34B69
         Left            =   1440
         List            =   "frmReport.frx":34B8B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cboOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmReport.frx":34BD4
         Left            =   1440
         List            =   "frmReport.frx":34BF6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboCol 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmReport.frx":34C3F
         Left            =   480
         List            =   "frmReport.frx":34C41
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin Reports_Factory.ucButtons_H cmdFilter 
         Height          =   330
         Left            =   1240
         TabIndex        =   23
         Top             =   -105
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Focus           =   0   'False
         cGradient       =   0
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmReport.frx":34C43
      End
      Begin Reports_Factory.ucButtons_H cmdFilterData 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Filter Data   "
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   14737632
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   3
         Image           =   "frmReport.frx":34F5D
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmReport.frx":352F7
      End
      Begin Reports_Factory.ucButtons_H cmdResetFilter 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Reset Filter   "
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   14737632
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   3
         Image           =   "frmReport.frx":35611
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmReport.frx":359AB
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   345
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      DragIcon        =   "frmReport.frx":35CC5
      Height          =   6975
      Left            =   120
      TabIndex        =   37
      Top             =   840
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12303
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      ForeColor       =   0
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin Reports_Factory.ucButtons_H cmdExit 
      Height          =   375
      Left            =   10800
      TabIndex        =   43
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&E&xit  "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmReport.frx":35FCF
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmReport.frx":36369
   End
   Begin Reports_Factory.ucButtons_H cmdHelp 
      Height          =   375
      Left            =   8760
      TabIndex        =   42
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Help"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmReport.frx":36683
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmReport.frx":39705
   End
   Begin Reports_Factory.ucButtons_H cmdReset 
      Height          =   495
      Left            =   12840
      TabIndex        =   36
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "&Reset All"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmReport.frx":39A1F
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmReport.frx":39DB9
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      DragIcon        =   "frmReport.frx":3A0D3
      Height          =   6975
      Left            =   120
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12303
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin Reports_Factory.ucButtons_H cmdPrevious 
      Height          =   375
      Left            =   2640
      TabIndex        =   39
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "   &Previous"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmReport.frx":3A3DD
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmReport.frx":3A777
   End
   Begin Reports_Factory.ucButtons_H cmdSaveReport 
      Height          =   375
      Left            =   6720
      TabIndex        =   41
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "&Publish Report "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmReport.frx":3AA91
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmReport.frx":3AE2B
   End
   Begin Reports_Factory.ucButtons_H cmdNext 
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   8160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Next   "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   14737632
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   3
      Image           =   "frmReport.frx":3B145
      cBack           =   -2147483633
      mPointer        =   99
      mIcon           =   "frmReport.frx":3B4DF
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6000
      TabIndex        =   34
      Top             =   540
      UseMnemonic     =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblGrpData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grouped Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   600
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblFiltered 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtered Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   600
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options   "
      Begin VB.Menu mnuSetHeadings 
         Caption         =   "&Customize Column Headings"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Set Font"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuFlxOpt 
      Caption         =   "&Flex Grid Options"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCol 
         Caption         =   "&Add Column"
      End
      Begin VB.Menu mnuDelCol 
         Caption         =   "&Delete Column"
      End
      Begin VB.Menu mnuAddCalc 
         Caption         =   "&Add Calculations"
      End
      Begin VB.Menu mnuAddSum 
         Caption         =   "&Add Sum"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolAddData As Boolean
Public lngCurCol As Long

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFilter_Click()
If frmCalculate.Visible = True Then
    MsgBox "You cannot activate filter when the calculation window is active. Please close this to continue", vbExclamation
    Exit Sub
ElseIf lblGrpData.Visible = True Then
    MsgBox "You cannot activate filter when group by is active. Please reset this to continue", vbExclamation
    Exit Sub
End If

If Frame1.Width <> 1500 Then
    Frame1.Width = 1500
    Frame1.Height = 255
    cmdFilter.Caption = "4"
    MSFlexGrid1.Enabled = True
ElseIf Frame1.Width <> 4380 Then
    If Frame2.Width <> 1380 Then
        Frame2.Width = 1380
        Frame2.Height = 255
        cmdGroupBy.Caption = "4"
    End If
    Frame1.Width = 4380
    Frame1.Height = 3615
    cmdFilter.Caption = "3"
    MSFlexGrid1.Enabled = False
    
    On Error Resume Next
    MSFlexGrid1.Parent.Controls.Remove "txt_txt_txt"
End If

SendKeys "{TAB}"
End Sub

Private Sub cmdGroupBy_Click()
If frmCalculate.Visible = True Then
    MsgBox "You cannot activate group by when the calculation window is active. Please close this to continue", vbExclamation
    Exit Sub
ElseIf lblFiltered.Visible = True Then
    MsgBox "You cannot activate group by when filter is active. Please reset this to continue", vbExclamation
    Exit Sub
End If

If Frame2.Width <> 1380 Then
    Frame2.Width = 1380
    Frame2.Height = 255
    cmdGroupBy.Caption = "4"
ElseIf Frame2.Width <> 4500 Then
    If Frame1.Width <> 1500 Then
        Frame1.Width = 1500
        Frame1.Height = 255
        cmdFilter.Caption = "4"
    End If
    Frame2.Width = 4500
    Frame2.Height = 1455
    cmdGroupBy.Caption = "3"
    
    On Error Resume Next
    MSFlexGrid1.Parent.Controls.Remove "txt_txt_txt"
End If

SendKeys "{TAB}"
End Sub

Private Sub cmdGrpFilter_Click()
MSFlexGrid2.Redraw = False
If Len(Trim(cboGrpCol(0).Text)) > 0 Then
    MSFlexGrid1.Visible = False
    MSFlexGrid2.Visible = True
    Dim lngX As Long
    Dim strError As String
    MSFlexGrid2.Clear
    MSFlexGrid2.Rows = 3
    MSFlexGrid2.Cols = 2
    MSFlexGrid2.FixedCols = 1
    MSFlexGrid2.FixedRows = 2
    
    'set headings
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Row = 1
    MSFlexGrid2.CellAlignment = 4
    MSFlexGrid2.CellFontBold = True
    MSFlexGrid2.Text = MSFlexGrid1.TextMatrix(1, cboGrpCol(0).ListIndex)
    
    MSFlexGrid2.Cols = MSFlexGrid2.Cols + 1
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Row = 1
    MSFlexGrid2.CellAlignment = 4
    MSFlexGrid2.CellFontBold = True
    MSFlexGrid2.Text = "SUM(" & MSFlexGrid1.TextMatrix(1, cboGrpCol(0).ListIndex) & ")"
    
    MSFlexGrid2.Cols = MSFlexGrid2.Cols + 1
    MSFlexGrid2.Col = 3
    MSFlexGrid2.Row = 1
    MSFlexGrid2.CellAlignment = 4
    MSFlexGrid2.CellFontBold = True
    MSFlexGrid2.Text = "COUNT(" & MSFlexGrid1.TextMatrix(1, cboGrpCol(0).ListIndex) & ")"
    
    MSFlexGrid2.Cols = MSFlexGrid2.Cols + 1
    MSFlexGrid2.Col = 4
    MSFlexGrid2.Row = 1
    MSFlexGrid2.CellAlignment = 4
    MSFlexGrid2.CellFontBold = True
    MSFlexGrid2.Text = "AVG(" & MSFlexGrid1.TextMatrix(1, cboGrpCol(0).ListIndex) & ")"
    
    If Len(Trim(cboGrpCol(1))) > 0 Then
         'Get distinct col data
         MSFlexGrid1.Col = cboGrpCol(0).ListIndex
         lngX = 2
         lstCheck.Clear
         While lngX + 1 < MSFlexGrid1.Rows
             MSFlexGrid1.Row = lngX
             If chkListMatch(lstCheck, MSFlexGrid1.Text) = False Then
                 lstCheck.AddItem MSFlexGrid1.Text
             End If
             
             lngX = lngX + 1
         Wend
        
        'Add data to grid
         Dim lngChk As Long
         Dim dblSum1 As Double, dblCount1 As Double
         dblSum1 = 0: dblCount1 = 0
         lngChk = 0
         lngX = 2
         While lngChk <= lstCheck.ListCount
            dblSum1 = 0: dblCount1 = 0: strError = ""
            lngX = 2
            MSFlexGrid1.Col = cboGrpCol(0).ListIndex
            While lngX + 1 < MSFlexGrid1.Rows
                MSFlexGrid1.Col = cboGrpCol(0).ListIndex
                MSFlexGrid1.Row = lngX
                If UCase(MSFlexGrid1.Text) = UCase(lstCheck.List(lngChk)) Then
                     MSFlexGrid1.Col = cboGrpCol(1).ListIndex
                     
                     If strError = "" Then
                         On Error GoTo err:
                         dblSum1 = dblSum1 + CDbl(MSFlexGrid1.Text)
                     End If
                     dblCount1 = dblCount1 + 1
                End If
               lngX = lngX + 1
            Wend
            If strError <> "" Then
                MSFlexGrid2.AddItem "" & vbTab & lstCheck.List(lngChk) & vbTab & strError & vbTab _
                & dblCount1 & vbTab & "", 2
            Else
                MSFlexGrid2.AddItem "" & vbTab & lstCheck.List(lngChk) & vbTab & dblSum1 & vbTab _
                & dblCount1 & vbTab & Round(dblSum1 / dblCount1, 3), 2
            End If
            lngChk = lngChk + 1
         Wend
    End If
    
    Call AltFlexiColors(MSFlexGrid2, 2, 1)
    Call setRowNums(Me.MSFlexGrid2)
    Call setColNums(Me.MSFlexGrid2)
    Call FG_AutosizeCols(MSFlexGrid2, Me, , , True)
    Call cmdGroupBy_Click
    lblGrpData.Visible = True
End If
MSFlexGrid2.Redraw = True
Exit Sub

err:
strError = "Err in Sum"
dblSum1 = 0
Resume Next
End Sub

Private Sub cmdGrpFilterReset_Click()
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
lblGrpData.Visible = False

Call cmdGroupBy_Click
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1015)
End Sub

Private Sub cmdNext_Click()
If MSFlexGrid1.Visible = True Then
    MSFlexGrid1.Visible = False
    Call PrintFlexi(txtReportTitle.Text, Me.MSFlexGrid1)
    MSFlexGrid1.Visible = True
ElseIf MSFlexGrid2.Visible = True Then
    MSFlexGrid2.Visible = False
    Call PrintFlexi(txtReportTitle.Text, Me.MSFlexGrid2)
    MSFlexGrid2.Visible = True
End If
Me.Hide
frmPrint.Show
End Sub

Private Sub cmdPrevious_Click()
boolFromPrevious = True
frmConnect.Visible = True
Me.Hide
Unload Me
End Sub

Private Sub cmdReset_Click()
boolConfirm = MsgBox("This will reset the entire report." & vbCrLf & "You will loose" & vbCrLf _
& "1. All added fields" & vbCrLf _
& "2. All calculations" & vbCrLf _
& "3. All customized column headings" & vbCrLf & vbCrLf _
& "Are you sure you want to continue ?", vbExclamation + vbYesNoCancel + vbDefaultButton3)
If boolConfirm = vbYes Then
    frmHeadings.MSFlexGrid1.Rows = 1
    frmHeadings.MSFlexGrid1.Rows = 2
    
    'same as form load but no run options
    Call forFormLoad
    
    MSFlexGrid1.Visible = True
    MSFlexGrid2.Visible = False
    lblFiltered.Visible = False
    lblGrpData.Visible = False
End If
End Sub

Private Sub cmdResetFilter_Click()
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
lblFiltered.Visible = False

Call cmdFilter_Click
End Sub

Private Sub cmdSaveReport_Click()
frmSaveReport.Show , Me
Me.Enabled = False
End Sub

Private Sub Form_Load()
If boolFromRun = True Then
    Call forFormLoad
    MSFlexGrid1.Redraw = False
    Dim lngArr As Long
    Dim lngDelCol As Long
       
    'set table name aliases
    lngArr = 0
    If Len(strAlias(0)) > 0 Then
        frmHeadings.MSFlexGrid1.Rows = 1
        frmHeadings.MSFlexGrid1.Rows = 2
        While lngArr <= UBound(strAlias)
            frmHeadings.MSFlexGrid1.AddItem strAlias(lngArr), 1
            lngArr = lngArr + 1
        Wend
        Call frmHeadings.setHeadings
    End If
    
    'add calculated fields
    Dim strCalc() As String
    Dim lngC As Long
    Dim lngCol As Long
    Dim varValue
    lngArr = 1
    MSFlexGrid3.Rows = 1
    MSFlexGrid3.Rows = 2
    While lngArr < UBound(strCalcField)
        MSFlexGrid3.AddItem strCalcField(lngArr) & vbTab & strCalcField(lngArr + 1), MSFlexGrid3.Rows - 1
        lngArr = lngArr + 2
    Wend
    'add col & values
    lngArr = 1
    While lngArr + 1 < MSFlexGrid3.Rows
        MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = MSFlexGrid1.Cols - 1
        MSFlexGrid1.CellFontBold = True
        
        MSFlexGrid1.Text = MSFlexGrid3.TextMatrix(lngArr, 0)
        MSFlexGrid1.Redraw = True
        
        MSFlexGrid1.Redraw = False
        strCalc = Split(MSFlexGrid3.TextMatrix(lngArr, 1), vbCrLf)
        Load frmCalculate
        frmCalculate.Visible = False
        lngC = 0
        Dim lngControls As Integer
        lngControls = 0
        frmCalculate.lblColInv.Caption = MSFlexGrid1.Col
        While lngC <= UBound(strCalc)
            If Left(UCase(strCalc(lngC)), Len("Report.")) = UCase("Report.") Then
                frmCalculate.cboFrom(lngControls) = "From Report"
                frmCalculate.txtVal(lngControls) = getColPos(Mid(strCalc(lngC), Len("Report.") + 1))
                frmCalculate.txtCol(lngControls) = frmCalculate.txtVal(lngControls)
                If lngC < UBound(strCalc) Then
                    frmCalculate.cboOP(lngControls) = strCalc(lngC + 1)
                End If
                lngControls = lngControls + 1
            ElseIf Left(UCase(strCalc(lngC)), Len("Value.")) = UCase("Value.") Then
                frmCalculate.cboFrom(lngControls) = "Value"
                frmCalculate.txtVal(lngControls) = Mid(strCalc(lngC), Len("Value.") + 1)
                frmCalculate.txtCol(lngControls) = frmCalculate.txtVal(lngControls)
                If lngC < UBound(strCalc) Then
                    frmCalculate.cboOP(lngControls) = strCalc(lngC + 1)
                End If
                lngControls = lngControls + 1
            End If
            'since data array is already traversed
            lngC = lngC + 2
        Wend
'        frmCalculate.Visible = True
        Call frmCalculate.cmdOK_Click
        Unload frmCalculate
        lngArr = lngArr + 1
    Wend
    
    'delete un-wanted cols
    lngArr = 0
    If Len(strDeleteCols(0)) > 0 Then
        MSFlexGrid1.Redraw = False
        While lngArr <= UBound(strDeleteCols)
            lngDelCol = 1
            MSFlexGrid1.Row = 1
            While lngDelCol < MSFlexGrid1.Cols
                MSFlexGrid1.Col = lngDelCol
                If UCase(MSFlexGrid1.Text) = UCase(strDeleteCols(lngArr)) Then
                    Call FG_RemoveColumn(MSFlexGrid1, lngDelCol)
                End If
                lngDelCol = lngDelCol + 1
            Wend
            lngArr = lngArr + 1
        Wend
        MSFlexGrid1.Redraw = True
    End If
    
    Call setRowNums(MSFlexGrid1)
    Call FG_AutosizeRows(MSFlexGrid1, Me, , , True)
    Call SortFlexiArrows(MSFlexGrid1, True, 1)
    Call AltFlexiColors(MSFlexGrid1, 2, 1)
    Call setColNums(MSFlexGrid1)
    Call FG_AutosizeCols(MSFlexGrid1, Me, , , True)
    
    MSFlexGrid1.FixedCols = 1
    MSFlexGrid1.FixedRows = 2
    
    Call loadFilters
    MSFlexGrid1.Redraw = True
Else
    Call forFormLoad
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Visible = True Then
    boolConfirm = MsgBox("Are you sure you want to exit ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirm <> vbYes Then
        Cancel = 1
        Exit Sub
    End If
    Unload frmHeadings
    Call unloadAllForms
    Call unloadAllForms
End If
End Sub

Private Sub mnuAddCalc_Click()
frmCalculate.Show , Me
frmCalculate.lblColInv.Caption = MSFlexGrid1.Col
End Sub

Private Sub mnuAddCol_Click()
Dim strColHeading As String
strColHeading = ""
While Len(Trim(strColHeading)) <= 0
    strColHeading = InputBox("Enter a title for the column", "Reports Factory")
    strColHeading = Trim(strColHeading)
    Dim lngX As Long
    lngX = 1
    MSFlexGrid1.Row = 1
    While lngX < MSFlexGrid1.Cols
        MSFlexGrid1.Col = lngX
        If UCase(MSFlexGrid1.Text) = "[NEW]." & UCase(strColHeading) Then
            strColHeading = ""
            MsgBox "The column heading specified by you already exist. Please specify a new heading", vbExclamation
        End If
        lngX = lngX + 1
    Wend
Wend
MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
MSFlexGrid1.ColPosition(MSFlexGrid1.Cols - 1) = lngCurCol
MSFlexGrid1.Col = lngCurCol
MSFlexGrid1.Row = 1
MSFlexGrid1.Text = "[NEW]." & strColHeading
MSFlexGrid1.CellAlignment = 4
MSFlexGrid1.CellFontBold = True

Call FG_AutosizeCols(MSFlexGrid1, Me, , , True)
Call SortFlexiArrows(MSFlexGrid1, True, 1)
Call AltFlexiColors(Me.MSFlexGrid1, 2, 1)
Call setRowNums(Me.MSFlexGrid1)
Call setColNums(Me.MSFlexGrid1)
Call loadFilters

MSFlexGrid3.AddItem "[NEW]." & strColHeading, MSFlexGrid3.Rows - 1
End Sub

Private Sub mnuAddSum_Click()
On Error Resume Next
If mnuAddSum.Caption = "&Clear Total" Then
    MSFlexGrid1.Text = ""
    Exit Sub
End If
Dim dblSum As Double
Dim lngX As Long

dblSum = 0
lngX = 2
While lngX + 1 < MSFlexGrid1.Rows
    dblSum = dblSum + MSFlexGrid1.TextMatrix(lngX, MSFlexGrid1.Col)
    lngX = lngX + 1
Wend

MSFlexGrid1.Text = "Total=" & Round(dblSum, 4)
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.CellAlignment = 4
End Sub

Private Sub mnuDelCol_Click()
Dim strColHeading As String
strColHeading = ""
If MSFlexGrid1.Col <> 0 Then
    boolConfirm = MsgBox("Are you sure you want to delete this column ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
    If boolConfirm = vbYes Then
        MSFlexGrid1.Row = 1
        strColHeading = MSFlexGrid1.Text
        Call FG_RemoveColumn(MSFlexGrid1, MSFlexGrid1.Col)
        Call SortFlexiArrows(MSFlexGrid1, True, True, 1)
        Call setRowNums(MSFlexGrid1)
        Call setColNums(MSFlexGrid1)
        Call loadFilters
        
        Dim lngX As Long
        lngX = 1
        MSFlexGrid3.Col = 0
        While lngX + 1 < MSFlexGrid3.Rows
            MSFlexGrid3.Row = lngX
            If UCase(strColHeading) = MSFlexGrid3.Text Then
                MSFlexGrid3.RemoveItem lngX
                Exit Sub
            End If
            lngX = lngX + 1
        Wend
    End If
End If
End Sub

Private Sub mnuFind_Click()
frmFind.Show , Me
End Sub

Private Sub mnuFont_Click()
frmFont.Show , Me
End Sub

Private Sub mnuSetHeadings_Click()
If MSFlexGrid1.Visible = False Then
    MsgBox "You cannot rename headers in filter mode. You can click reset filter to continue", vbExclamation
Else
    Me.Enabled = False
    frmHeadings.Visible = True
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
Call SortFlexiArrows(MSFlexGrid1, True, True)
Call setColNums(MSFlexGrid1)
End Sub

Private Sub MSFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
If MSFlexGrid1.MouseCol = 0 Then Exit Sub
If MSFlexGrid1.Tag = "" Then Exit Sub
MSFlexGrid1.Redraw = False
MSFlexGrid1.ColPosition(Val(MSFlexGrid1.Tag)) = MSFlexGrid1.MouseCol
Call setRowNums(MSFlexGrid1)
MSFlexGrid1.Redraw = True
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MSFlexGrid1.Tag = ""
lngCurCol = MSFlexGrid1.Col
If Button = 2 Then
    If Left(MSFlexGrid1.TextMatrix(1, MSFlexGrid1.Col), 6) = "[NEW]." Then
        mnuAddCalc.Enabled = True
    Else
        mnuAddCalc.Enabled = False
    End If
    
    
    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
        mnuAddSum.Enabled = True
        If Len(MSFlexGrid1.Text) > 0 Then
            mnuAddSum.Caption = "&Clear Total"
        Else
            mnuAddSum.Caption = "&Add Total"
        End If
    Else
        mnuAddSum.Enabled = False
    End If
    
    mnuDelCol.Caption = "Delete Column " & MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col)
    PopupMenu mnuFlxOpt
    Exit Sub
End If

If frmCalculate.Visible = True Then
    If frmCalculate.intIndex <> -1 Then
        MSFlexGrid1.Row = 0
        frmCalculate.txtVal(frmCalculate.intIndex).Text = MSFlexGrid1.Text
        If frmCalculate.cboFrom(frmCalculate.intIndex).Text = "From Report" Then
            frmCalculate.txtCol(frmCalculate.intIndex).Text = MSFlexGrid1.Col
        Else
            frmCalculate.txtCol(frmCalculate.intIndex).Text = ""
        End If
    End If
Else
    If MSFlexGrid1.MouseCol = 0 Then Exit Sub
    If MSFlexGrid1.MouseRow <> 0 Then Exit Sub
    MSFlexGrid1.Tag = Str(MSFlexGrid1.MouseCol)
    MSFlexGrid1.Drag 1

End If
End Sub

Sub setRowNums(objMSF As MSFlexGrid)
With objMSF
Dim lngRow As Long, lngCol As Long
lngRow = .Row: lngCol = .Col
.Redraw = False
Dim lngX As Long
Dim strAlpha(0 To 25) As String
strAlpha(0) = "A"
strAlpha(1) = "B"
strAlpha(2) = "C"
strAlpha(3) = "D"
strAlpha(4) = "E"
strAlpha(5) = "F"
strAlpha(6) = "G"
strAlpha(7) = "H"
strAlpha(8) = "I"
strAlpha(9) = "J"
strAlpha(10) = "K"
strAlpha(11) = "L"
strAlpha(12) = "M"
strAlpha(13) = "N"
strAlpha(14) = "O"
strAlpha(15) = "P"
strAlpha(16) = "Q"
strAlpha(17) = "R"
strAlpha(18) = "S"
strAlpha(19) = "T"
strAlpha(20) = "U"
strAlpha(21) = "V"
strAlpha(22) = "W"
strAlpha(23) = "X"
strAlpha(24) = "Y"
strAlpha(25) = "Z"

lngX = 1
.Row = 0
Dim intQ As Integer
Dim intR As Integer
        
While lngX < .Cols
    .Col = lngX
    .CellAlignment = 4
            
        intQ = -1: intR = -1
        intQ = CLng(lngX) \ 26
        intQ = intQ - 1
        
        If CLng(lngX) Mod 26 = 0 Then
            intQ = intQ - 1
            intR = 25
        Else
            intR = CDbl(lngX) - (26 * (CLng(lngX) \ 26))
            intR = intR - 1
        End If
        
        .Text = ""
        If intQ >= 0 Then
            .Text = strAlpha(intQ)
        End If
        If intR >= 0 Then
            .Text = .Text & strAlpha(intR)
        End If
            
    lngX = lngX + 1
Wend

.Row = lngRow: .Col = lngCol
.Redraw = True
End With
End Sub

Sub setColNums(objMSF As MSFlexGrid)
With objMSF
Dim lngRow As Long, lngCol As Long
lngRow = .Row: lngCol = .Col

.Redraw = False
Dim lngX As Long
lngX = 2
.ColAlignment(0) = 4
While lngX + 1 < .Rows
    .TextMatrix(lngX, 0) = lngX - 1
    lngX = lngX + 1
Wend

.Row = lngRow: .Col = lngCol
.Redraw = True
End With
End Sub

Sub loadFilters()
With MSFlexGrid1
Dim lngX As Long
lngX = 1
cboCol(0).Clear: cboCol(1).Clear: cboCol(2).Clear: cboCol(3).Clear: cboCol(4).Clear
cboCol(0).AddItem "": cboCol(1).AddItem "": cboCol(2).AddItem "": cboCol(3).AddItem "": cboCol(4).AddItem ""

cboGrpCol(0).Clear: cboGrpCol(1).Clear
cboGrpCol(0).AddItem "": cboGrpCol(1).AddItem ""

While lngX < .Cols
    cboCol(0).AddItem .TextMatrix(0, lngX)
    cboCol(1).AddItem .TextMatrix(0, lngX)
    cboCol(2).AddItem .TextMatrix(0, lngX)
    cboCol(3).AddItem .TextMatrix(0, lngX)
    cboCol(4).AddItem .TextMatrix(0, lngX)
    
    cboGrpCol(0).AddItem .TextMatrix(0, lngX)
    cboGrpCol(1).AddItem .TextMatrix(0, lngX)
    
    lngX = lngX + 1
Wend
    
End With
End Sub

Private Sub cmdFilterData_Click()
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = True

MSFlexGrid2.Font.Name = MSFlexGrid1.Font.Name
MSFlexGrid2.Font.Size = MSFlexGrid1.Font.Size
    
Dim strText As String
Dim lngRow As Long
Dim lngCol As Long
lngRow = 0: lngCol = 0
'set col headings
MSFlexGrid2.Clear
MSFlexGrid2.Rows = 3
MSFlexGrid2.Cols = MSFlexGrid1.Cols
MSFlexGrid2.FixedCols = 1
MSFlexGrid2.FixedRows = 2
While lngRow <= 1
    MSFlexGrid1.Row = lngRow
    MSFlexGrid2.Row = lngRow
    lngCol = 0
    While lngCol < MSFlexGrid1.Cols
        MSFlexGrid1.Col = lngCol
        MSFlexGrid2.Col = lngCol
        MSFlexGrid2.Text = MSFlexGrid1.Text
        MSFlexGrid2.CellAlignment = MSFlexGrid1.CellAlignment
        MSFlexGrid2.CellFontBold = MSFlexGrid1.CellFontBold
        lngCol = lngCol + 1
    Wend
    lngRow = lngRow + 1
Wend

'set data
boolAddData = True
lngRow = 2: lngCol = 1
While lngRow + 1 < MSFlexGrid1.Rows
    MSFlexGrid1.Row = lngRow
    lngCol = 1
    boolAddData = True
    While lngCol < MSFlexGrid1.Cols
        MSFlexGrid1.Col = lngCol
        If Len(cboCol(0).Text) > 0 And Len(cboOP(0).Text) > 0 Then
            If MSFlexGrid1.Col = cboCol(0).ListIndex Then
                Call validateData(0, MSFlexGrid1.Text)
            End If
        End If
        
        If Len(cboCol(1).Text) > 0 And Len(cboOP(1).Text) > 0 Then
            If MSFlexGrid1.Col = cboCol(1).ListIndex Then
                Call validateData(1, MSFlexGrid1.Text)
            End If
        End If
        
        If Len(cboCol(2).Text) > 0 And Len(cboOP(2).Text) > 0 Then
            If MSFlexGrid1.Col = cboCol(2).ListIndex Then
                Call validateData(2, MSFlexGrid1.Text)
            End If
        End If
        
        If Len(cboCol(3).Text) > 0 And Len(cboOP(3).Text) > 0 Then
            If MSFlexGrid1.Col = cboCol(3).ListIndex Then
                Call validateData(3, MSFlexGrid1.Text)
            End If
        End If
        
        If Len(cboCol(4).Text) > 0 And Len(cboOP(4).Text) > 0 Then
            If MSFlexGrid1.Col = cboCol(4).ListIndex Then
                Call validateData(4, MSFlexGrid1.Text)
            End If
        End If
        
        lngCol = lngCol + 1
    Wend
    lngCol = 1
    strText = ""
    If boolAddData = True Then
        While lngCol < MSFlexGrid1.Cols
            MSFlexGrid1.Col = lngCol
            strText = strText & vbTab & MSFlexGrid1.Text
            lngCol = lngCol + 1
        Wend
        MSFlexGrid2.AddItem strText, MSFlexGrid2.Rows - 1
    End If
    lngRow = lngRow + 1
Wend

Call AltFlexiColors(Me.MSFlexGrid2, 2, 1)
Call setColNums(MSFlexGrid2)
Call cmdFilter_Click
Call FG_AutosizeCols(MSFlexGrid2, Me, , , True)

lblFiltered.Visible = True
End Sub

Sub validateData(intIndex As Integer, strMSFText As String)
On Error GoTo err
Select Case cboOP(intIndex).Text
    Case "<"
        If CDbl(txtValue(intIndex).Text) >= CDbl(strMSFText) Then
            boolAddData = False
        End If
    Case ">"
        If CDbl(txtValue(intIndex).Text) <= CDbl(strMSFText) Then
            boolAddData = False
        End If
    Case "="
        If UCase(txtValue(intIndex).Text) <> UCase(strMSFText) Then
            boolAddData = False
        End If
    Case "<>"
        If txtValue(intIndex).Text = strMSFText Then
            boolAddData = False
        End If
    Case "Begins With"
        If UCase(Left(strMSFText, Len(txtValue(intIndex).Text))) <> UCase(txtValue(intIndex).Text) Then
            boolAddData = False
        End If
    Case "Contains"
        If InStr(UCase(strMSFText), UCase(txtValue(intIndex).Text)) = 0 Then
            boolAddData = False
        End If
    Case "Ends With"
        If UCase(Right(strMSFText, Len(txtValue(intIndex).Text))) <> UCase(txtValue(intIndex).Text) Then
            boolAddData = False
        End If
    Case "Non Blanks"
        If Len(Trim(strMSFText)) <= 0 Then
            boolAddData = False
        End If
    Case "Blanks"
        If Len(Trim(strMSFText)) > 0 Then
            boolAddData = False
        End If
End Select

err:
Exit Sub
End Sub

Sub forFormLoad()
With MSFlexGrid1
.Redraw = False
Dim rsData As New ADODB.Recordset
Dim lngX As Long
Dim strData As String

lngX = 0
.Rows = 3
.Cols = 2
.FixedCols = 1
.FixedRows = 2

If Len(Trim(strReportTitle)) > 0 Then
    txtReportTitle.Text = strReportTitle
Else
    txtReportTitle.Text = "Report 1"
End If

Set rsData = Nothing
rsData.Open "select * from Data", dbLocal, adOpenDynamic, adLockOptimistic
'set cols
If rsData.Fields.Count > 0 Then
    .Cols = rsData.Fields.Count + 1
End If
'set data headings
lngX = 1
.Row = 1
Load frmHeadings
frmHeadings.Visible = False
frmHeadings.MSFlexGrid1.Rows = 1
frmHeadings.MSFlexGrid1.Rows = 2
While lngX <= rsData.Fields.Count
    .Col = lngX
    .Text = Replace(rsData.Fields(lngX - 1).Name, "__", ".")
    frmHeadings.MSFlexGrid1.AddItem .Text & vbTab & .Text, 1
    .CellAlignment = 4
    .CellFontBold = True
    lngX = lngX + 1
Wend

'set data

If rsData.EOF = False Then
    rsData.MoveFirst
    While rsData.EOF = False
        .Row = .Rows - 1
        lngX = 1
        While lngX <= rsData.Fields.Count
            .Col = lngX
            If Len(Trim(rsData(lngX - 1))) > 0 Then
                .Text = rsData(lngX - 1)
                .CellAlignment = 1
            End If
            lngX = lngX + 1
        Wend
        .Rows = .Rows + 1
        rsData.MoveNext
    Wend
End If



Call setRowNums(MSFlexGrid1)
Call FG_AutosizeRows(MSFlexGrid1, Me, , , True)
Call SortFlexiArrows(MSFlexGrid1, True, 1)
Call AltFlexiColors(MSFlexGrid1, 2, 1)
Call setColNums(MSFlexGrid1)
Call FG_AutosizeCols(MSFlexGrid1, Me, , , True)

.FixedCols = 1
.FixedRows = 2

Call loadFilters
.Redraw = True

End With
End Sub

Private Sub MSFlexGrid2_DblClick()
Call SortFlexiArrows(MSFlexGrid2, True, True)
Call setColNums(MSFlexGrid2)
End Sub

Function getColPos(strString As String) As Integer
Dim lngX As Long
Dim lngY As Long
Dim strNewHead As String
getColPos = -1
strNewHead = ""
lngX = 1

While lngX + 1 < frmHeadings.MSFlexGrid1.Rows
    If UCase(frmHeadings.MSFlexGrid1.TextMatrix(lngX, 0)) = UCase(strString) Then
        strNewHead = frmHeadings.MSFlexGrid1.TextMatrix(lngX, 1)
        While lngY + 1 < MSFlexGrid1.Cols
            If UCase(MSFlexGrid1.TextMatrix(1, lngY)) = UCase(strNewHead) Then
                'getColPos = MSFlexGrid1.TextMatrix(0, lngY)
                getColPos = lngY
                Exit Function
            End If
            lngY = lngY + 1
        Wend
    End If
    lngX = lngX + 1
Wend

lngX = 1
While lngX + 1 <= MSFlexGrid1.Cols
    If UCase(strString) = UCase(MSFlexGrid1.TextMatrix(1, lngX)) Then
        getColPos = lngX
        Exit Function
    End If
    lngX = lngX + 1
Wend
End Function
