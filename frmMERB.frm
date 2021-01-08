VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMERB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-EditROM and Multi-ROM Builder and Compare Utility"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cbAllowShort 
      Caption         =   "Allow short files"
      Height          =   345
      Left            =   8850
      TabIndex        =   67
      Top             =   810
      Width           =   2055
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "frmMERB.frx":0000
      Left            =   6510
      List            =   "frmMERB.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   810
      Width           =   2100
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare"
      Height          =   405
      Left            =   7830
      TabIndex        =   48
      Top             =   7560
      Width           =   1755
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   1800
      TabIndex        =   44
      Text            =   "Multi-EditROM Set"
      Top             =   810
      Width           =   4605
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "DOWN"
      Height          =   405
      Left            =   6750
      TabIndex        =   42
      Top             =   7560
      Width           =   885
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Height          =   405
      Left            =   5790
      TabIndex        =   41
      Top             =   7560
      Width           =   885
   End
   Begin VB.CommandButton cmdIns 
      Caption         =   "Insert Entry"
      Height          =   405
      Left            =   3720
      TabIndex        =   40
      Top             =   7560
      Width           =   1755
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   11070
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   15
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   38
      Top             =   7110
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   14
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   36
      Top             =   6720
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   13
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   34
      Top             =   6330
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   12
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   32
      Top             =   5940
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   11
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   30
      Top             =   5550
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   10
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   28
      Top             =   5160
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   9
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   4770
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   8
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   24
      Top             =   4380
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   7
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
      Top             =   3990
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   6
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   3600
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   5
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      Top             =   3210
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   4
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   2820
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   3
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   2430
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   2
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   2040
      Width           =   6825
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   1
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   1650
      Width           =   6825
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build It!"
      Height          =   405
      Left            =   9690
      TabIndex        =   8
      Top             =   7560
      Width           =   1755
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete Entry"
      Height          =   405
      Left            =   1890
      TabIndex        =   7
      Top             =   7560
      Width           =   1755
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Binary..."
      Height          =   405
      Left            =   60
      TabIndex        =   6
      Top             =   7560
      Width           =   1755
   End
   Begin VB.TextBox txtFN 
      Height          =   345
      Index           =   0
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   1260
      Width           =   6825
   End
   Begin VB.CommandButton cmdSaveSet 
      Caption         =   "Save Set"
      Height          =   645
      Left            =   8220
      TabIndex        =   2
      Top             =   60
      Width           =   1605
   End
   Begin VB.CommandButton cmdLoadSet 
      Caption         =   "Load Set"
      Height          =   645
      Left            =   6540
      TabIndex        =   1
      Top             =   60
      Width           =   1605
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   645
      Left            =   9900
      TabIndex        =   0
      Top             =   60
      Width           =   1605
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   10710
      Top             =   1680
      Width           =   165
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   10110
      Top             =   1680
      Width           =   165
   End
   Begin VB.Label Label5 
      Caption         =   "2K          4K         Bad"
      Height          =   255
      Left            =   9660
      TabIndex        =   65
      Top             =   1665
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   195
      Left            =   9465
      Top             =   1680
      Width           =   165
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   7320
      TabIndex        =   64
      Top             =   7110
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   7320
      TabIndex        =   63
      Top             =   6720
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   7320
      TabIndex        =   62
      Top             =   6330
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   7320
      TabIndex        =   61
      Top             =   5940
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   7320
      TabIndex        =   60
      Top             =   5550
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   7320
      TabIndex        =   59
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   7320
      TabIndex        =   58
      Top             =   4770
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   7320
      TabIndex        =   57
      Top             =   4380
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   7320
      TabIndex        =   56
      Top             =   3990
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   7320
      TabIndex        =   55
      Top             =   3600
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   7320
      TabIndex        =   54
      Top             =   3210
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   7320
      TabIndex        =   53
      Top             =   2820
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   7320
      TabIndex        =   52
      Top             =   2430
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   7320
      TabIndex        =   51
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7320
      TabIndex        =   50
      Top             =   1650
      Width           =   345
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7320
      TabIndex        =   49
      Top             =   1260
      Width           =   345
   End
   Begin VB.Label Label4 
      Caption         =   "File Size:"
      Height          =   255
      Left            =   7830
      TabIndex        =   47
      Top             =   1650
      Width           =   645
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   8520
      TabIndex        =   46
      Top             =   1650
      Width           =   90
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000016&
      Caption         =   "Info"
      Height          =   5445
      Left            =   7830
      TabIndex        =   45
      Top             =   1980
      Width           =   3585
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Set Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   43
      Top             =   810
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7830
      TabIndex        =   39
      Top             =   1260
      Width           =   1080
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   30
      TabIndex        =   37
      Top             =   7110
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   30
      TabIndex        =   35
      Top             =   6720
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   30
      TabIndex        =   33
      Top             =   6330
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   30
      TabIndex        =   31
      Top             =   5940
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   30
      TabIndex        =   29
      Top             =   5550
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   30
      TabIndex        =   27
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   30
      TabIndex        =   25
      Top             =   4770
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "09"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   30
      TabIndex        =   23
      Top             =   4380
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   30
      TabIndex        =   21
      Top             =   4020
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   30
      TabIndex        =   19
      Top             =   3630
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   30
      TabIndex        =   17
      Top             =   3240
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   30
      TabIndex        =   15
      Top             =   2850
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   30
      TabIndex        =   13
      Top             =   2460
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   30
      TabIndex        =   11
      Top             =   2070
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   30
      TabIndex        =   9
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmMERB.frx":0029
      Height          =   615
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   6375
   End
End
Attribute VB_Name = "frmMERB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MERB - Multi-EditorROM Builder, (C) 2017-2021 Steve J. Gray
' ====

Dim SelNum As Integer
Dim Cr As String

Private Sub Form_Load()
    Cr = Chr(13)        'Carriage Return
    SelectN 0           'Set First Text Box as selected
    cboMode.ListIndex = 0
End Sub

Private Sub cmdAbout_Click()
    MsgBox "MultiEditROM and MultiROM Builder, (C)2017-2018 Steve J. Gray" & Cr & "Version 1.32 - Nov 29/2018"
End Sub

Private Sub lblN_DblClick(Index As Integer)
    cmdAdd_Click
End Sub

'---
Private Sub txtFN_GotFocus(Index As Integer)
    SelectN Index
End Sub

Private Sub txtFN_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index < 16 Then txtFN(Index + 1).SetFocus
End Sub

'--- Display only Filename when focus is lost
Private Sub txtFN_LostFocus(Index As Integer)
    txtFN(Index).Tag = txtFN(Index).Text
    txtFN(Index).Text = FName(txtFN(Index).Tag)
End Sub


Private Sub lblN_Click(Index As Integer)
    SelectN Index
End Sub

Private Sub lblK_Click(Index As Integer)
    SelectN Index
End Sub
Private Sub cmdAdd_Click()
    Dim Filename As String
    
    Filename = FileOpenSave("", 0, 2, "Add ROM")
    If Filename <> "" Then
        txtFN(SelNum).Tag = Filename
        SelectN SelNum
    End If

End Sub
Private Sub cmdLoadSet_Click()
    Dim Filename As String
    Dim FIO As Integer, i As Integer, Tmp As String
    
    On Local Error Resume Next                          'Allow incomplete set file
    
    Filename = FileOpenSave("", 0, 1, "Load Set")
    If Exists(Filename) = True Then
        FIO = FreeFile
        Open Filename For Input As FIO
        Line Input #FIO, Tmp: txtDesc.Text = Tmp        'Set Description
        For i = 0 To 15
            Tmp = ""
            Line Input #FIO, Tmp                        'Path+Filename
            txtFN(i).Text = FName(Tmp)                  'Filename Only for display
            txtFN(i).Tag = Tmp                          'Path+Filename
        Next i
        Close FIO
        SelectN 0                                       'Select first file slot
    End If
    
End Sub

Private Sub cmdSaveSet_Click()
    Dim Filename As String
    Dim FIO As Integer, i As Integer, Tmp As String
    
    Filename = FileOpenSave("", 1, 1, "Save Set")
    If Overwrite(Filename) = True Then
        FIO = FreeFile
        Open Filename For Output As FIO
        Print #FIO, txtDesc.Text                        'Set Description
        For i = 0 To 15
            Print #FIO, txtFN(i).Tag                    'Path+Filename
        Next i
        Close FIO
    End If
    
End Sub

Private Sub cmdDown_Click()
    Dim Tmp As String, Tmp2 As String, RGB As Long
    
    If SelNum < 15 Then
        Tmp = txtFN(SelNum).Tag
        Tmp2 = txtFN(SelNum).Text
        
        txtFN(SelNum).Tag = txtFN(SelNum + 1).Tag
        txtFN(SelNum + 1).Tag = Tmp
        txtFN(SelNum).Text = txtFN(SelNum + 1).Text
        txtFN(SelNum + 1).Text = Tmp2
                
        RGB = lblK(SelNum).BackColor
        lblK(SelNum).BackColor = lblK(SelNum + 1).BackColor
        lblK(SelNum + 1).BackColor = RGB
        
        SelectN SelNum + 1
    End If
    
End Sub

Private Sub cmdUp_Click()
    Dim Tmp As String, Tmp2 As String, RGB As Long
    
    If SelNum > 0 Then
        Tmp = txtFN(SelNum).Tag
        Tmp2 = txtFN(SelNum).Text
        
        txtFN(SelNum).Tag = txtFN(SelNum - 1).Tag
        txtFN(SelNum - 1).Tag = Tmp
        txtFN(SelNum).Text = txtFN(SelNum - 1).Text
        txtFN(SelNum - 1).Text = Tmp2
        
        RGB = lblK(SelNum).BackColor
        lblK(SelNum).BackColor = lblK(SelNum - 1).BackColor
        lblK(SelNum - 1).BackColor = RGB
        
        SelectN SelNum - 1
    End If
    
End Sub

Private Sub cmdDel_Click()
    Dim Tmp As String, i As Integer, RGB As Long
    
    If SelNum = 16 Then
        txtFN(SelNum).Tag = ""
        txtFN(SelNum).Text = ""
        lblK(SelNum).BackColor = vbBlack
    Else
        For i = SelNum To 14
            txtFN(i).Text = txtFN(i + 1).Text
            txtFN(i).Tag = txtFN(i + 1).Tag
            lblK(i).BackColor = lblK(i + 1).BackColor
        Next
        txtFN(15).Text = ""
        lblK(15).BackColor = vbBlack
    End If
End Sub
Private Sub cmdIns_Click()
    Dim i As Integer, RGB As Long
    
        If SelNum < 15 Then
            For i = 15 To SelNum + 1 Step -1
                txtFN(i).Text = txtFN(i - 1).Text
                txtFN(i).Tag = txtFN(i - 1).Tag
                lblK(i).BackColor = lblK(i - 1).BackColor
            Next
        End If
        txtFN(SelNum).Tag = ""
        txtFN(SelNum).Text = ""
        lblK(SelNum).BackColor = vbBlack
        
End Sub

Private Sub SelectN(ByVal Index As Integer)
    Dim i As Integer
    
    For i = 0 To 15
        If i = Index Then
            lblN(i).BackColor = vbRed           'Selected is made RED
            lblN(i).ForeColor = vbWhite
        Else
            lblN(i).BackColor = vbBlue          'Un-Selected is BLUE
            lblN(i).ForeColor = vbWhite
        End If
    Next
    SelNum = Index                              'Remember it for other operations
    ShowInfo Index                              'Get info from file
    
    txtFN(Index).Text = txtFN(Index).Tag
    DoEvents
End Sub

Private Sub ShowInfo(ByVal Index As Integer)
    Dim Tmp As String, Filename As String, FIO As Integer
    Dim FLen As Integer, Tmp2 As String
    
    Tmp = ""
    Filename = txtFN(Index).Text
    lblSize.Caption = "N/A"                                 'Assume no file
    
    If Exists(Filename) = True Then
        FIO = FreeFile
        Open Filename For Binary As FIO
        FLen = LOF(FIO)
        
        Select Case FLen
            Case 2048, 2050: lblK(Index).BackColor = vbYellow
            Case 4096, 4098: lblK(Index).BackColor = vbGreen
            Case Else: lblK(Index).BackColor = vbRed
        End Select
        
        lblSize.Caption = Str(FLen)
        
        If FLen > 2048 Then
            Tmp2 = Input(2048, FIO)                         'Read and ignore first 2K
            Tmp2 = Input(256, FIO)                          'Read IO Area
            Tmp = StripIt(Tmp2)                             'Extract the text in the IO area"
        End If
        If Tmp = "" Then Tmp = "No info available"          '2K files have no IO area
    Else
        Tmp = "File does not exist!"                        'Couldn't find file
        lblK(Index).BackColor = vbBlack
    End If
    
    Close FIO
    lblInfo.Caption = Tmp                                   'Update Info area
    DoEvents
    
End Sub

'--- Build the ROM
Private Sub cmdBuild_Click()
    Dim Filename As String, FIO As Integer, FIO2 As Integer, FLen As Integer
    Dim i As Integer, j As Integer, Buf As String, Padd As String, Mode As Integer
    
    Padd = Chr(0)
    Mode = cboMode.ListIndex
    
    '--- check that all the files exist
    Flag = 0
    For i = 0 To 15
        Filename = txtFN(i).Text
        If Exists(Filename) = False Then MsgBox "Slot " & Str(i + 1) & " is unspecifiied or does not exist": Exit Sub
        If cbAllowShort.Value = vbUnchecked Then If FileLen(Filename) < 2048 Then MsgBox "The file '" & Filename & "' is < 2K bytes!": Exit Sub
        If FileLen(Filename) > 4096 Then MsgBox "The file '" & Filename & "' is > 4K bytes!": Exit Sub
    Next i

    '--- Get a filename
    Filename = FileOpenSave("", 1, 2, "Add ROM"): If Filename = "" Then Exit Sub
    If Overwrite(Filename) = False Then Exit Sub
    
    '--- Open the Output file
    FIO = FreeFile
    Open Filename For Output As FIO: DoEvents
    
    HideSizes
       
    '--- Process Files
    For i = 0 To 15
        Filename = txtFN(i).Text
        
        lblInfo.Caption = "Writing " & Filename & "...": DoEvents
        
        FIO2 = FreeFile
        Open Filename For Binary As FIO2: FLen = LOF(FIO2)                                      'Open file and get length
        Buf = Input(FLen, FIO2)                                                                 'Read entire file to buffer
        Close FIO2                                                                              'Close it
        
        If (FLen = 2050) Or (FLen = 4098) Then Buf = Mid(Buf, 2)                                'Remove first two bytes (load address)
        Print #FIO, Buf;
        
        '-- Padd short file
        If FLen < 4096 Then
            Select Case Mode
                Case 0 'Padd
                    For j = 1 To 4096 - FLen: Print #FIO, Padd;: Next j                         'Pad the file to 4096 bytes
                Case 1 'Duplicate
                    If FLen < 2048 Then For j = 1 To 2048 - FLen: Print #FIO, Padd;: Next j     'Pad the file to 2048 bytes
                    Print #FIO, Buf;                                                            'Copy the contents
                    If FLen < 2048 Then For j = 1 To 2048 - FLen: Print #FIO, Padd;: Next j     'Pad the file to 4096 bytes
            End Select
        End If
        
        lblK(i).Visible = True: DoEvents

    Next i
    
    Close FIO
    
    MsgBox "File successfully created!!!"
    
End Sub
'--- Compare ROMs
Private Sub cmdCompare_Click()
    Dim Filename As String, FIO As Integer, FIO2 As Integer, FLen As Integer, FLen2 As Integer
    Dim i As Integer, j As Integer, Buf As String, Buf2 As String, Difs As Integer
    Dim FX(15) As Boolean 'File Exists Flags array
    Dim Cr As String, Results As String, B1 As String, B2 As String
    
    Cr = Chr(13)
   
    '--- Check target filename
    Filename = txtFN(SelNum): If Exists(Filename) = False Then MsgBox "You must select a SLOT containing a file!": Exit Sub
        
    '--- Open the Output file
    FIO = FreeFile
    Open Filename For Binary As FIO: FLen = LOF(FIO)    'Open file and get length
    Buf = Input(FLen, FIO)                              'Read entire file to buffer
    Close FIO                                           'Close the file
    
    HideSizes
    
    Results = "Comparing to SLOT " & Format(SelNum + 1) & "..." & Cr          'Initial result text
    
    '--- Process Files
    For i = 0 To 15
        Filename = txtFN(i).Text
            
        lblInfo.Caption = "Reading slot " & Format(i)                           'Show Progress
        DoEvents
        
        If (Exists(Filename) = True) And (i <> SelNum) Then
            FIO2 = FreeFile
            Open Filename For Binary As FIO2: FLen2 = LOF(FIO2)                 'Open file and get length
            Buf2 = Input(FLen2, FIO2)                                           'Read entire file to buffer
            Close FIO2                                                          'Close it
            
            Difs = 0
            Results = Results & "SLOT" & Format(i + 1) & ": "                   'Add slot#
            
            For j = 1 To FLen
                If j > FLen2 Then Results = Results & " is shorter.": Exit For  'done comparing
                B1 = Mid(Buf, j, 1)
                B2 = Mid(Buf2, j, 1)
                If B1 <> B2 Then Difs = Difs + 1
            Next j
            If FLen2 > FLen Then Results = Results & " is longer."              'File is Longer
            
            If Difs = 0 Then
                Results = Results & "MATCHES!"                                  'File is SAME
            Else
                Results = Results & Format(Difs) & " bytes differ."             'File DIFFERS
            End If
            Results = Results & Cr                                              'Add CR
        End If
        
        lblK(i).Visible = True: DoEvents                                        'Show box
        
    Next i
    
    Close FIO
    
    lblInfo.Caption = Results
    
End Sub
'===================
' FUNCTIONS and SUBS
'===================

Private Function Exists(ByVal Filename As String) As Boolean
    Dim FIO As Integer
    
    On Local Error GoTo NoFile              'Open will fail if file does not exist
    FIO = FreeFile
    Open Filename For Input As FIO          'If this works then the file exists
    Close FIO
    Exists = True                           'Return TRUE
    Exit Function

NoFile:
    Close FIO
    Exists = False                          'Return FALSE
    
End Function

'--- Extracts only printable characters from string
Private Function StripIt(ByVal S As String) As String
    Dim S2 As String, M As String, MV As Integer, A As Integer
        
    S2 = ""                                             'Start with empty string
    For A = 1 To Len(S)
        M = Mid(S, A, 1): MV = Asc(M)                   'One byte and it's ascii value
        If (MV > 31) And (MV < 128) Then S2 = S2 & M    'If in range add it
    Next A
    
    StripIt = S2                                        'Return string
    
End Function

'--- Drag and Drop
' To enable, set OLEDropMode to "1 - Manual" for each textFN control
Private Sub txtFN_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Filename As String
      
    If Data.GetFormat(vbCFFiles) Then
        Dim vFn As Variant
        For Each vFn In Data.Files
            Filename = (vFn)                            'vFn is name of file dropped
            txtFN(Index).Text = Filename                'Set the text box to filename
            Index = Index + 1                           'Point to next slot
            If Index > 15 Then Exit For                 'All slots are filled, so done
            SelectN SelNum                              'Get info and set selected
            SelNum = SelNum + 1                         'Make it selected slot
        Next vFn
    End If

End Sub

'--- Provide feedback to user
' If dragging a FILE then accept it, otherwise no.
Private Sub txtFN_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '0=do not allow drop, 1=inform source that data will be copied
    If Data.GetFormat(vbCFFiles) Then Effect = 1 Else Effect = 0
End Sub

'--- Common File Open or Save Dialog
' You can specify a default filename, a File Filter list index (0-1), and Window Title
' MODE: 0=Open, 1=Save
' Returns a filename with full path. If cancelled will return null string
Private Function FileOpenSave(ByVal DefFile As String, ByVal Mode As Integer, FiltSet As Integer, DTitle As String) As String
    Dim Filename As String
    
    CommonDialog.CancelError = True
    On Local Error GoTo NoFile
        
    CommonDialog.DialogTitle = DTitle
    CommonDialog.Flags = cdlOFNHideReadOnly
    CommonDialog.Filename = DefFile
    
    Select Case FiltSet
        Case 0: CommonDialog.Filter = "All files (*.*)|*.*"
        Case 1: CommonDialog.Filter = "Text Files (*.TXT)|*.TXT"
        Case 2: CommonDialog.Filter = "ROM Files (*.bin, *.rom)|*.bin;*.rom"
    End Select
    
    If Mode = 0 Then CommonDialog.ShowOpen Else CommonDialog.ShowSave   'MODE: 0=Open, 1=Save
        
    If CommonDialog.Filename = "" Then Exit Function
    
    FileOpenSave = CommonDialog.Filename
    Exit Function
NoFile:

End Function

'---- Checks for file and prompts to Overwrite if necessary
' Returns TRUE if file does NOT exist, or it EXISTS and user says YES.
' Returns FALSE if file EXISTS but user says NO.
Public Function Overwrite(ByVal Filename As String) As Boolean
    
    Overwrite = True    'assume ok to replace
    
    If Exists(Filename) = True Then
        If MsgBox("The file '" & Filename & "' already exists!" & Cr & "Replace it?", vbYesNo, "Overwrite File") = vbNo Then Overwrite = False
    End If
End Function


Public Function HideSizes()
    Dim i As Integer
    
    For i = 0 To 15: lblK(i).Visible = False: Next
    DoEvents

End Function

' Return the filename only from the end of the path
Public Function FName(ByVal Path As String) As String

Dim j As Integer

j = InStrRev(Path, "\")
If j > 0 Then
    FName = Mid(Path, j + 1)
Else
    FName = Path
End If

End Function
