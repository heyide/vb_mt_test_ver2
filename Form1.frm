VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "12306货运整车预订工具"
   ClientHeight    =   8145
   ClientLeft      =   1305
   ClientTop       =   1185
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   14160
   Begin VB.TextBox txt_pzycfh 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7440
      TabIndex        =   58
      Top             =   7560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtExtcode 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8880
      TabIndex        =   57
      Top             =   7080
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txt_fzyx 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10080
      TabIndex        =   56
      Top             =   7080
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txt_fztmism 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8880
      TabIndex        =   55
      Top             =   7560
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txt_dztmism 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10320
      TabIndex        =   54
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_dzyx 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7440
      TabIndex        =   53
      Top             =   7080
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      Caption         =   "货物信息"
      Height          =   1215
      Left            =   5520
      TabIndex        =   50
      Top             =   2040
      Width           =   8535
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4440
         TabIndex        =   75
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5640
         TabIndex        =   72
         Top             =   360
         Width           =   2415
         Begin VB.OptionButton opt_rdSdz 
            Caption         =   "散堆装"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton opt_rdSdz 
            Caption         =   "非散堆装"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   960
            TabIndex        =   73
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.OptionButton opt_rdPs 
         Caption         =   "不怕湿"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4680
         TabIndex        =   71
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt_rdPs 
         Caption         =   "怕湿"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3960
         TabIndex        =   70
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txt_hwbz 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   68
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txt_hzpm 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   51
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "千克"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6120
         TabIndex        =   77
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "最大单件重量 "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3120
         TabIndex        =   76
         Top             =   765
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "货物特征"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3120
         TabIndex        =   69
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "货物包装"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   67
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "货物名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "发货信息（红字必填，下同）"
      Height          =   1695
      Left            =   5520
      TabIndex        =   37
      Top             =   240
      Width           =   8535
      Begin VB.TextBox txt_fhdwdz 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   49
         Top             =   1200
         Width           =   7455
      End
      Begin VB.TextBox txt_fhdwdh 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7200
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_fhdwmc 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   45
         Top             =   255
         Width           =   3135
      End
      Begin VB.TextBox txt_zcdd 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   40
         Top             =   720
         Width           =   5175
      End
      Begin VB.TextBox txt_fjm 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7200
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_fzhzzm 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4920
         TabIndex        =   38
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "通信地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "移动电话"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6360
         TabIndex        =   46
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "装车地点"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "托运人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "发局"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6720
         TabIndex        =   42
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "发站"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4440
         TabIndex        =   41
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Timer tmrPostOrdertest 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   12480
      Top             =   7080
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   11280
      Pattern         =   "*.dat"
      TabIndex        =   34
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrPostOrder 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   12000
      Top             =   7080
   End
   Begin VB.CommandButton cmd_contractOrder 
      Caption         =   "<< 收起"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   5520
      TabIndex        =   32
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Caption         =   "收货信息"
      Height          =   2055
      Left            =   5520
      TabIndex        =   23
      Top             =   3360
      Width           =   8535
      Begin VB.TextBox txt_shdwdz 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   65
         Top             =   1275
         Width           =   7455
      End
      Begin VB.CheckBox chk_dddxtz 
         Caption         =   "收货人接收到货短信"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   63
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_shdwdh 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6840
         TabIndex        =   61
         Top             =   795
         Width           =   1575
      End
      Begin VB.TextBox txt_dzhzzm 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5040
         TabIndex        =   27
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox txt_djm 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7200
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt_shdwmc 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   25
         Top             =   390
         Width           =   3135
      End
      Begin VB.TextBox txt_xcdd 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   24
         Top             =   795
         Width           =   4335
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "通信地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "收货人手机号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5640
         TabIndex        =   62
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "到站"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4560
         TabIndex        =   31
         Top             =   435
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "到局"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6720
         TabIndex        =   30
         Top             =   435
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "收货人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   435
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "卸车地点"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1080
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "运输信息"
      Height          =   1335
      Left            =   5520
      TabIndex        =   17
      Top             =   5520
      Width           =   8535
      Begin VB.CheckBox chk_tbfs 
         Caption         =   "保价运输"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   66
         Top             =   803
         Width           =   1215
      End
      Begin VB.TextBox txt_ytcs 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3600
         TabIndex        =   59
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chk_ifzzjg 
         Caption         =   "装载加固"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4320
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cbo_cz 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   960
         List            =   "Form1.frx":0016
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_cc 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   18
         Top             =   795
         Width           =   1935
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "车数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   60
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "车种"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "车次"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "自动提报参数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7455
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton cmd_expandOrder 
         Caption         =   ">> 展开订单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         TabIndex        =   82
         Top             =   6000
         Width           =   1605
      End
      Begin VB.ListBox txt_orderlist 
         Height          =   2040
         ItemData        =   "Form1.frx":0048
         Left            =   240
         List            =   "Form1.frx":004A
         TabIndex        =   81
         Top             =   3000
         Width           =   4935
      End
      Begin VB.CommandButton cmdDeAuto 
         Caption         =   "停止自动提交"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         Picture         =   "Form1.frx":004C
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   6720
         Width           =   1605
      End
      Begin VB.CommandButton cmdAuto 
         Caption         =   "开始自动提交"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3480
         Picture         =   "Form1.frx":00DD
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6720
         Width           =   1605
      End
      Begin VB.TextBox txt_AllowAuto 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   35
         Top             =   5475
         Width           =   2505
      End
      Begin VB.CheckBox chk_saveacc 
         Caption         =   "保存"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox txtUsername 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmd_profile 
         Caption         =   "保存订单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "Form1.frx":016F
         TabIndex        =   33
         Top             =   6000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmd_manual 
         Caption         =   "手动提交"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3480
         Picture         =   "Form1.frx":07F6
         TabIndex        =   6
         Top             =   6000
         Width           =   1605
      End
      Begin VB.CommandButton cmd_getorder 
         Caption         =   "获取历史单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   2400
         Width           =   1605
      End
      Begin VB.CommandButton cmd_login 
         Caption         =   "登录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   3
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txt_zyrq 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   4
         Top             =   1665
         Width           =   1740
      End
      Begin VB.OptionButton Option1 
         Caption         =   "今天"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2880
         TabIndex        =   11
         Top             =   1665
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "30天以后加1天"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3585
         TabIndex        =   10
         Top             =   1680
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   795
         Width           =   1740
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "(手动填写请按""2014-01-05""的的格式填写)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   2080
         Width           =   3495
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "自动提交时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   5520
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "未登录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1455
         TabIndex        =   15
         Top             =   1245
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "当前登陆状态:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   1245
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "快捷录入历史单据列表:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "装车日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1710
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "帐号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   405
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "密码:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.Label lblInfo 
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   16
      Top             =   7800
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LocalIP As String
Public http As WinHttp.WinHttpRequest
Public sen As String, sen2 As String, sen3 As String
Public vcodeIndex As Long
Public jsonorder As String, jsonorder2 As String, uuid As String
Public ISAUTO As Boolean, ISLOGIN As Boolean, ISOFFLINE As Boolean
Public JsonselIndex As Integer
Public city As String, testurl As String, testurl2 As String
Public yzmCode As String
Public heartline As Integer


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function LoadLibFromFile Lib "Sunday.dll" (ByVal FilePath As String, ByVal pass As String) As Long
Private Declare Function GetCodeFromBuffer Lib "Sunday.dll" (ByVal CdsFileIndex As Long, ByVal ImgBuffer As Long, ByVal ImgBufLen As Long, ByVal Vcode As String) As Boolean






Private Sub Form_Load()

    'city = "wulmq"
    city = "beij"
    'city = "taiy"
    
    If city = "beij" Then
        testurl = "_test1"  '登陆地址
        testurl2 = "_test1" '"_test2"  '内页地址
    Else
        esturl = ""
        testurl2 = ""
    End If
    JsonselIndex = -1
    Set http = New WinHttp.WinHttpRequest
    http.Option(4) = 13056
    http.Option(6) = False
    http.SetTimeouts 60000, 60000, 60000, 60000
    ISAUTO = False
    ISOFFLINE = False
    yzmCode = ""
    heartline = 0
    
    '提交日期
    If Hour(Now()) >= 8 Then '8点以后第二天提交
        txt_AllowAuto.Text = Format(DateAdd("d", 1, Now()), "yyyy-mm-dd 07:00:00")
    Else '7点之前当天提交
        txt_AllowAuto.Text = Format(Now(), "yyyy-mm-dd 07:00:00")
    End If
    
    '装车日期
    txt_zyrq.Text = Trim(Format(DateAdd("d", 31, Now()), "yyyy-mm-dd"))
    
    '加载账号
    Call bindAccount
    

    Call showinfo(3, "当前帐号尚未登录,请先测试登录!")

End Sub


'加载订单信息
Private Sub txt_profile_click()
    
    Call showinfo(2, "该功能暂不可用,待修复")
    Exit Sub

    If ISLOGIN = True Then
        Call showinfo(2, "当前已登录,无法加载订单信息")
        Exit Sub
    End If
    
    Call loadProfile(txt_profile.List(txt_profile.ListIndex))
    
    Call showinfo(1, "订单信息加载完成!")
End Sub

'选择用户名
Private Sub txtUsername_Click()
    Call bindAccount(txtUsername.Text)
End Sub

'点登录
Private Sub cmd_login_Click()
    Dim funRe As String
    
    Dim username As String, password As String
    If txtUsername.Text = "" Then
        Call showinfo(2, "请输入用户名!")
        txtUsername.SetFocus
        Exit Sub
    End If
    username = Trim(txtUsername.Text)
    
    If txtPassWord.Text = "" Then
        Call showinfo(2, "请输入密码!")
        txtPassWord.SetFocus
        Exit Sub
    End If
    password = Trim(txtPassWord.Text)
    
    
    Call showinfo(3, "登录中,请稍等....")
    cmd_login.Enabled = False
    
    funRe = intiAndLoginFull(username, password)
    
    If CheckFunRe(funRe, 1) <> 1 Then
        Call showinfo(2, "登录失败,错误原因:" & CheckFunRe(funRe, 2))
        cmd_login.Enabled = True
    Else
        Label6.ForeColor = &HD000&
        Label6.Caption = "已登录(" & CheckFunRe(funRe, 2) & ")"
        ISLOGIN = True
        cmd_login.Enabled = True
        
        '成功以后再保存
        If chk_saveacc.Value = 1 Then
            Call saveAccount(txtUsername.Text, txtPassWord.Text)
            Call showinfo(1, "登录成功,账号密码已自动保存!")
        Else
            Call showinfo(1, "登录成功!")
        End If
    End If
End Sub

'点获取预定号
Private Sub cmd_getorder_Click()
    Dim funRe As String
    
    If ISLOGIN = False Then
        Call showinfo(2, "请先登录!")
        Exit Sub
    End If
    
    Call showinfo(3, "获取预定号中,请稍等....")
    cmd_getorder.Enabled = False
    
    funRe = GetOrderNo()
    
    If CheckFunRe(funRe, 1) <> 1 Then
        Call showinfo(2, "获取失败,错误原因:" & CheckFunRe(funRe, 2))
        cmd_getorder.Enabled = True
    Else
        Call showinfo(1, "获取成功,请选择预定号")
        cmd_getorder.Enabled = True
    End If
End Sub

'选择预定号
Private Sub txt_orderlist_Click()
    
    If ISOFFLINE = True Then Exit Sub

    Dim funRe As String
    
    If txt_orderlist.ListCount = 0 Then
        Call showinfo(2, "请先获取预定号!")
        Exit Sub
    End If
    
    Call showinfo(3, "根据预定号获取订单信息中....")
    cmd_getorder.Enabled = False
    
    funRe = GetInfoByOrderNo(txt_orderlist.ListIndex)
    
    If CheckFunRe(funRe, 1) <> 1 Then
        Call showinfo(2, "获取失败,错误原因:" & CheckFunRe(funRe, 2))
        cmd_getorder.Enabled = True
    Else
        Call showinfo(1, "订单填写完成,可以进入手动或自动提交模式")
        cmd_getorder.Enabled = True
    End If
End Sub

'点自动提交
Private Sub cmdAuto_Click()

    Dim offline As Integer

    '早上7点到11点之间 使用离线订单提示

    

    'offline = MsgBox("是否要使用离线订单提交?", vbYesNo, "自动提交")
    
    'If offline = vbNo Then
    '    Exit Sub
    'End If

    If txtUsername.Text = "" Or txtPassWord.Text = "" Or txt_zyrq.Text = "" Or txt_pzycfh.Text = "" Or txt_AllowAuto.Text = "" Or txt_xqslh.Text = "" Then
        Call showinfo(2, "资料填写不完全,请手动登录获取订单信息后再点击自动提交")
        Exit Sub
    End If
    
    If JsonselIndex = -1 Or jsonorder = "" Or jsonorder2 = "" Then
        Call showinfo(2, "订单填写不完全,请选择预定号或手动填写订单信息后再点击自动提交")
        Exit Sub
    End If
    
    If Val(txt_qqcs.Text) > Val(txt_qqcsMax.Text) Then
        Call expandOrder
        txt_qqcs.SetFocus
        Call showinfo(2, "订车数不能超过最大车数！")
        Exit Sub
    End If

    ISAUTO = True
    Call showinfo(3, "自动提交启动中,为避免误操作,请不要点击其他按钮")
    Call SavePage("[" & Now() & "]自动提交启动...", "syslog")
    
    tmrPostOrdertest.Interval = 5000
    tmrPostOrdertest.Enabled = True
    
    Call lockAll
       
End Sub

'点取消自动提交
Private Sub cmdDeAuto_Click()
    ISAUTO = False
    Call showinfo(2, "自动提交关闭")
    tmrPostOrdertest.Enabled = False
    
    Call unlockAll

End Sub

'自动提交流程
Private Sub tmrPostOrder_Timer()

    On Error Resume Next
    DoEvents
    Dim funRe As String
    funRe = 0
    
    Call showinfo(3, "自动提交中,为避免误操作,请不要点击其他按钮")
    
    tmpTime = DateDiff("s", Now(), txt_AllowAuto)
    
    '提前三分钟获取验证码
    If tmpTime > 300 Then
       Call showinfo(2, "未到提交时间,系统待机中,还有" & tmpTime \ 60 & "分开始提交")
       Exit Sub
    ElseIf yzmCode = "" Then
        Call SavePage("[" & Now() & "]自动提交初始化开始", "syslog")
        Do
            funRe = inti(txtUsername.Text)
            
            If CheckFunRe(funRe, 1) <> 1 Then
                Call SavePage("[" & Now() & "]登录初始化失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
            End If
            
            Sleep (1000)
            
        Loop Until CheckFunRe(funRe, 1) = 1
    End If
   
    
    '提前5秒开始提交
    If tmpTime > 0 Then
        Call showinfo(2, "未到提交时间,系统待机中,还有" & tmpTime \ 60 & "分开始提交")
        Exit Sub
    End If
    
    Call SavePage("[" & Now() & "]自动提交开始,开始登录", "syslog")
    
    '登陆
    Do
        funRe = Login(txtUsername.Text, txtPassWord.Text)
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]登陆失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
            If CheckFunRe(funRe, 2) = "系统维护中" Then
                Exit Sub
            End If
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]登陆成功,开始提交", "syslog")
    
    http.SetTimeouts 180000, 180000, 180000, 180000
    
    '登陆完直接提交,跳过检查订单号
    Do
        funRe = PerPost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]预提交失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
            
            If CheckFunRe(funRe, 2) = "超出可预订日期范围" Or CheckFunRe(funRe, 2) = "未找到对应的需求信息" Then
            
                '明确失败
                Call SavePage("[" & Now() & "]" & CheckFunRe(funRe, 2) & ",自动提交关闭", "syslog")
                ISAUTO = False
                Call showinfo(2, "信息填写或时间选择错误,自动提交关闭!")
                tmrPostOrder.Enabled = False
                
                Call unlockAll
                
                Exit Sub
            End If
            
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]预提交成功,uuid=" & uuid & ",开始正式提交", "syslog")
    '正式提交
    Do
        funRe = RePost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]正式提交失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]提交成功,自动提交关闭", "syslog")
    
    
    ISAUTO = False
    Call showinfo(1, "提交完成,自动提交关闭!")
    tmrPostOrder.Enabled = False
    
    Call unlockAll
    
End Sub

'新测试自动提交流程
Private Sub tmrPostOrdertest_Timer()
    On Error Resume Next
    DoEvents
    Dim funRe As String
    funRe = 0
    heartline = heartline + 1
    
    Call showinfo(3, "自动提交中,为避免误操作,请不要点击其他按钮")
    
    tmpTime = DateDiff("s", Now(), txt_AllowAuto)
    
    '提前五分钟不再心跳连接
    If tmpTime > 300 Then
        Call showinfo(2, "未到提交时间,系统待机中,还有" & (tmpTime \ 60) + 1 & "分开始提交")
        
        If heartline > 50 Then
            Call SavePage("[" & Now() & "]心跳连接开始" & sen2, "syslog")
           
            funRe = inti1(txtUsername.Text)
            
            If CheckFunRe(funRe, 1) <> 1 Then
                Call SavePage("[" & Now() & "]心跳连接失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
            End If
            
            heartline = 0
        End If
        
       Exit Sub
    End If
   
    
    '提前5秒开始提交
    If tmpTime > 5 Then
        Call showinfo(2, "未到提交时间,系统待机中,还有" & tmpTime \ 60 & "分开始提交")
        Exit Sub
    End If
    
    Call SavePage("[" & Now() & "]自动提交开始,开始预提交", "syslog")
    
    '登陆完直接提交,跳过检查订单号
    Do
        funRe = PerPost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]预提交失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
            
            If CheckFunRe(funRe, 2) = "超出可预订日期范围" Or CheckFunRe(funRe, 2) = "未找到对应的需求信息" Or CheckFunRe(funRe, 2) = "提报请求车数不能超过最大可用车数" Then
            
                '明确失败
                Call SavePage("[" & Now() & "]" & CheckFunRe(funRe, 2) & ",自动提交关闭", "syslog")
                ISAUTO = False
                Call showinfo(2, "信息填写或时间选择错误,自动提交关闭!")
                tmrPostOrdertest.Enabled = False
                
                Call unlockAll
                
                Exit Sub
            End If
            
            If CheckFunRe(funRe, 2) = "已丢失登录" Then
            
                '明确失败
                Call SavePage("[" & Now() & "]" & CheckFunRe(funRe, 2) & ",自动提交关闭", "syslog")
                ISAUTO = False
                Call showinfo(2, "已丢失登录状态，提交失败，请重新登录")
                tmrPostOrdertest.Enabled = False
                
                Call unlockAll
                ISLOGIN = False
                
                Exit Sub
            End If
            
        End If
        

        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]预提交成功,uuid=" & uuid & ",开始正式提交", "syslog")
    '正式提交
    Do
        funRe = RePost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]正式提交失败,错误原因:" & CheckFunRe(funRe, 2), "syslog")
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]提交成功,自动提交关闭", "syslog")
    
    
    ISAUTO = False
    Call showinfo(1, "提交完成,自动提交关闭!")
    tmrPostOrdertest.Enabled = False
    
    Call unlockAll
End Sub

'点手动提交
Private Sub cmd_manual_Click()
    On Error Resume Next
    
    If ISAUTO = True Then
        MsgBox "自动提交进行中,无法进行操作!"
        Exit Sub
    End If

    If txt_pzycfh.Text = "" Then Call showinfo(2, "资料不完整,请填写完整后再提交"): Exit Sub
    
    If Val(txt_qqcs.Text) > Val(txt_qqcsMax.Text) Then Call showinfo(2, "报车数不能超过最大可用车数，请检查"): Exit Sub

    Call showinfo(3, "提交处理中,请勿反复点击!")
    cmd_manual.Enabled = False
    Dim surl As String, param As String
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_add"
    param = ""
    param = param & "currentPosition=" & "%E9%A2%84%E7%BA%A6%C2%A0%3E%3E%C2%A0%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    param = param & "&" & "djm=" & URLEncodeUTF8(txt_djm.Text)
    param = param & "&" & "dzhzzm=" & URLEncodeUTF8(txt_dzhzzm.Text)
    param = param & "&" & "dztmism=" & txt_dztmism.Text
    param = param & "&" & "dzyx=" & Replace(txt_dzyx.Text, " ", "+")
    param = param & "&" & "fhdwmc=" & URLEncodeUTF8(txt_fhdwmc.Text)
    param = param & "&" & "fjm=" & URLEncodeUTF8(txt_fjm.Text)
    param = param & "&" & "fzhzzm=" & URLEncodeUTF8(txt_fzhzzm.Text)
    param = param & "&" & "fztmism=" & txt_fztmism.Text
    param = param & "&" & "fzyx=" & Replace(txt_fzyx.Text, " ", "+")
    param = param & "&" & "hzpm=" & URLEncodeUTF8(txt_hzpm.Text)
    param = param & "&" & "keyword="
    param = param & "&" & "maxDate=" & Trim(txt_zcrq.Text) '& Format(DateAdd("m", 1, Now()) - 1, "yyyy-mm-dd")
    param = param & "&" & "minDate=" & Format(Now() + 3, "yyyy-mm-dd")
    param = param & "&" & "po.dddxtz=" & chk_dddxtz.Value
    param = param & "&" & "po.hqhw=" & txt_hqhw.Text
    param = param & "&" & "po.pzycfh=" & txt_pzycfh.Text
    param = param & "&" & "po.qqcs=" & txt_qqcs.Text
    param = param & "&" & "po.qqcz=" & Right(cbo_qqcz.Text, 1)
    param = param & "&" & "po.qqds=" & txt_qqds.Text
    param = param & "&" & "po.qqlx=0"
    param = param & "&" & "po.shdwdh=" & txt_shdwdh.Text
    param = param & "&" & "po.uuid=" '8ac086a9441480d4014419d6acbe0064"
    param = param & "&" & "po.xqslh=" & txt_xqslh.Text
    
    param = param & "&" & "po.zcrq=" & Trim(txt_zcrq.Text)
    
    param = param & "&" & "qqcsMax=" & txt_qqcsMax.Text
    param = param & "&" & "shdwmc=" & URLEncodeUTF8(txt_shdwmc.Text)
    param = param & "&" & "xcdd=" & URLEncodeUTF8(txt_xcdd.Text)
    param = param & "&" & "zcdd=" & URLEncodeUTF8(txt_zcdd.Text)
    
    
    Call SavePage("[" & Now() & ":step1]" & param & vbLf, "perpostdata")
    
    http.Open "POST", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        Call showinfo(2, "超时1,请重新提交!")
        cmd_manual.Enabled = True
        Exit Sub
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage("[" & Now() & ":step1]" & body1 & vbLf, "pergetdata")
    
    If InStr(body1, """success"":true") Then
        uuid = mySubstr(body1, "uuid"":""", """")
        
        param = "op=10&uuids=" & uuid & ",&mor_dzsw_security_info=mor_dzsw_security_disabled"
        Call SavePage("[" & Now() & ":step2]" & param & vbLf, "perpostdata")
        surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_operateZcrbjh"
        
        http.Open "POST", surl, False
        http.SetRequestHeader "Connection", "Keep-Alive"
        http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
        http.SetRequestHeader "Cache-Control", "no-cache"
        http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
        http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
        http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
        http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
        http.SetRequestHeader "Cookie", "CASTGC=" & sen3
        http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
        http.SetRequestHeader "Content-Length", Len(param)
        http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
        http.Send param
        
        If Err.Number <> 0 Then
            Err.Clear
            Call showinfo(2, "超时2,请重新提交!")
            cmd_manual.Enabled = True
            Exit Sub
        End If
        
        body2 = BytesToBstr(http.ResponseBody, 2)
        Call SavePage("[" & Now() & ":step2]" & body2 & vbLf, "pergetdata")
        
        If InStr(body2, """success"":true") Then
            Call showinfo(1, "手动提报成功!")
            cmd_manual.Enabled = True
            Exit Sub
        Else
            Call showinfo(2, "提报失败,请检查日志!")
            cmd_manual.Enabled = True
            Exit Sub
        End If
    ElseIf InStr(body1, "超出可预订日期范围") Then
        Call showinfo(2, "超出可预订日期范围!")
        cmd_manual.Enabled = True
        Exit Sub
    ElseIf InStr(body1, "未找到对应的需求信息") Then
        Call showinfo(2, "未找到对应的需求信息,请重新检查所选预定号!")
        cmd_manual.Enabled = True
        Exit Sub
    ElseIf InStr(body1, "提报请求车数不能超过最大可用车数") Then
        Call showinfo(2, "提报请求车数不能超过最大可用车数，请确认")
        cmd_manual.Enabled = True
        Exit Sub
    ElseIf InStr(body1, "正在登录，请稍等...") Then
        Call showinfo(2, "系统错误，已丢失登录状态，请重新登录")
        cmd_manual.Enabled = True
        ISLOGIN = False

        Exit Sub
    Else
       Call showinfo(2, "预提报失败,请检查日志!")
       cmd_manual.Enabled = True
       Exit Sub
    End If
    
End Sub


'保存配置
Private Sub cmd_profile_Click()
    Dim filen As String
    
    If ISLOGIN = False Then
        Call showinfo(2, "请先登录后再保存当前用户名和密码!")
        Exit Sub
    End If
    
    
    If txtUsername.Text = "" Then
        Call showinfo(2, "请输入用户名!")
        txtUsername.SetFocus
        Exit Sub
    End If

    
    If txtPassWord.Text = "" Then
        Call showinfo(2, "请输入密码!")
        txtPassWord.SetFocus
        Exit Sub
    End If
    
    If txt_zyrq.Text = "" Then
        Call showinfo(2, "请输入装车时间!")
        txt_zcrq.SetFocus
        Exit Sub
    End If

    
    If jsonorder = "" Then
        Call showinfo(2, "请先获取预定号!")
        cmd_getorder.SetFocus
        Exit Sub
    End If
    
    If jsonorder2 = "" Or JsonselIndex = -1 Then
        Call showinfo(2, "请选择所需预定号!")
        txt_orderlist.SetFocus
        Exit Sub
    End If
    
    filen = "[" & txtUsername.Text & "]" & txt_zyrq.Text & "_订车数：" & txt_qqcs & "_到站：" & txt_dzhzzm & "_货物：" & txt_hzpm

    filen = InputBox("将所填写提报信息保存为:", "保存设置", filen)
    
    filen = Replace(Replace(Replace(Replace(Replace(Replace(filen, ":", ""), "/", ""), "\", ""), "|", ""), """", ""), "?", "")
    
    Call saveProfile(filen)
    
    Call showinfo(1, "当前订单已保存!")

End Sub

'**************************************************AUTO专用函数区*********************************************************
Function inti(user As String) As String

    On Error Resume Next

    Dim ImgFile As String
    Dim Image() As Byte
    
   
    '直接验证码
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/security/jcaptcha.jpg"
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        inti = "0|获取验证码超时002"
        Exit Function
    End If

    ImgFile = Fun_SaveImgToFile(http.ResponseBody, user & ".jpg", App.Path & "\")
    
    If Err.Number <> 0 Then
        Err.Clear
        inti = "0|验证码获取失败"
        Exit Function
    End If
    
    vcodeIndex = LoadLibFromFile("12306.lib", "123")
    
    If Err.Number <> 0 Then
        Err.Clear
        inti = "0|验证码识别组件加载失败"
        Exit Function
    End If


    If (vcodeIndex = -1) Then
        inti = "0|验证码识别库加载失败"
        Exit Function
    End If
    
    Dim Vcode As String
    Vcode = "      " '必须先对这个变量赋多个空格，空格数量要比验证码字符数量多1
   
    Call MyReadFile(ImgFile, Image)
     '内存接口调用验证码图像并识别
    If (GetCodeFromBuffer(vcodeIndex, VarPtr(Image(0)), UBound(Image), Vcode)) Then
        txtExtcode.Text = Vcode
        yzmCode = Trim(txtExtcode.Text)
        
        head = http.GetAllResponseHeaders
        headers = Split(head, Chr(10))
        
        For ii = LBound(headers) To UBound(headers)
            If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
                p2 = InStr(headers(ii), ";")
                s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
                p2 = InStr(s, "=")
                s1 = Trim(Mid(s, 1, p2 - 1))
                s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
                        
                If s1 = "BIGipServerhyswpt_pool" Then
                    sen = s2
                End If
                
                If s1 = "DZSW_SESSIONID" Then
                    sen2 = s2
                End If
            End If
        Next
        
        inti = "1|识别成功"
    Else
        inti = "0|验证码识别失败"
        Exit Function
    End If

End Function

'保持连接
Function inti1(user As String) As String

    On Error Resume Next
    
   
    '直接验证码
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/security/jcaptcha.jpg"
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        inti1 = "0|心跳连接超时"
        Exit Function
    End If

        
    head = http.GetAllResponseHeaders
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
                    
            If s1 = "BIGipServerhyswpt_pool" Then
                sen = s2
            End If
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
        End If
    Next
    
    inti1 = "1| " & sen2

End Function
'自动登陆
Function Login(user As String, pass As String) As String
    
    On Error Resume Next
    
    Dim username As String, password As String, extcode As String
    Dim param As String
    
    
    username = user
    password = pass
    extcode = Trim(txtExtcode.Text)
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/j_spring_security_check"
    param = "j_username=" & username & "&j_password=" & password & "&j_captcha=" & extcode & "&fromUrl=%2Flogin_bur.jsp"
    
    http.Open "POST", surl, False
    http.Option(WinHttpRequestOption_EnableRedirects) = 0
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"

    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        Login = "0|登录超时"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage("[" & Now() & "]httpStatus:" & http.Status & body1, "login")
    
    
    If InStr(body1, "Dzsw/home.jsp") > 0 Then
        Login = "1|登录成功"
    Else
        Login = "0|自动登录失败"
        Exit Function
    End If
    
    
    '更新cookie
    head = http.GetAllResponseHeaders
            
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
            
            If s1 = "CASTGC" Then
                sen3 = s2
            End If
        End If
    Next
    
    Exit Function
    
End Function

Function intiAndLoginFull(user As String, pass As String) As String
    
    On Error Resume Next

    Dim ImgFile As String
    Dim Image() As Byte
    
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|网络超时001"
        Exit Function
    End If
    
    head = http.GetAllResponseHeaders
            
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
                    
            If s1 = "BIGipServerhyswpt_pool" Then
                sen = s2
            End If
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
        End If
    Next
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    If InStr(body1, "系统正在维护中") > 0 Then
        intiAndLoginFull = "0|系统维护中"
        Exit Function
    End If
    
    '先读并显示验证码
    ' src="/vcode.php?rnd=78475"/>
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/security/jcaptcha.jpg"
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|获取验证码超时002"
        Exit Function
    End If
    

    ImgFile = Fun_SaveImgToFile(http.ResponseBody, user & ".jpg", App.Path & "\")
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|验证码获取失败"
        Exit Function
    End If
    
    vcodeIndex = LoadLibFromFile("12306.lib", "123")
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|验证码识别组件加载失败"
        Exit Function
    End If


    If (vcodeIndex = -1) Then
        intiAndLoginFull = "0|验证码识别库加载失败"
        Exit Function
    End If
    
    Dim Vcode As String
    Vcode = "      " '必须先对这个变量赋多个空格，空格数量要比验证码字符数量多1
   
    Call MyReadFile(ImgFile, Image)
     '内存接口调用验证码图像并识别
    If (GetCodeFromBuffer(vcodeIndex, VarPtr(Image(0)), UBound(Image), Vcode)) Then
        txtExtcode.Text = Vcode
    Else
        intiAndLoginFull = "0|验证码识别失败"
        Exit Function
    End If
    
    Dim username As String, password As String, extcode As String
    Dim param As String
    
    
    username = user
    password = pass
    extcode = Trim(txtExtcode.Text)
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/j_spring_security_check"
    param = "j_username=" & username & "&j_password=" & password & "&j_captcha=" & extcode & "&fromUrl=%2Flogin_bur.jsp"
    
    http.Open "POST", surl, False
    http.Option(WinHttpRequestOption_EnableRedirects) = 1
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"

    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|登录超时"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage(http.Status & body1, "login")
    
    
    If InStr(body1, "margin-left:50px;"">欢迎您：") > 0 Then
    'If InStr(body1, "Dzsw/home.jsp") > 0 Then
        intiAndLoginFull = "1|" & mySubstr(body1, ";white-space:nowrap;margin-left:5px;"">", "</span>")
    ElseIf InStr(body1, "系统正在维护中") > 0 Then
        intiAndLoginFull = "0|系统维护中"
        Exit Function
    ElseIf InStr(body1, "验证码输入不正确") > 0 Then  '验证码输入不正确
        intiAndLoginFull = "0|验证码错误"
        Exit Function
    Else
        intiAndLoginFull = "0|登录失败,请检查用户名与密码"
        Exit Function
    End If
    
    
    '更新cookie
    head = http.GetAllResponseHeaders
            
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
            
            If s1 = "CASTGC" Then
                sen3 = s2
            End If
        End If
    Next
    
    Exit Function
    
End Function

Function GetOrderNo() As String


    On Error Resume Next
    
    If ISLOGIN = False Then
       GetOrderNo = "0|请先登录"
       Exit Function
    End If
    
    Dim i As Integer
    Dim body1 As String, tmpStr As String
    
    If txt_zyrq.Text = "" Then
        GetOrderNo = "0|请选择装车日期"
        Exit Function
    End If
    
    'https://frontier."& city &".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_getYsxq?q=%E7%8E%89%E7%B1%B3&limit=50&timestamp=1389019837982&zcrq=2014-01-08
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_getYsxq?q="
    surl = surl & "&limit=50&timestamp=1389019837982&zcrq="
    surl = surl & Trim(txt_zyrq.Text)
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        GetOrderNo = "0|获取预定号超时"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage(body1, "jsonorder")
    
    jsonorder = body1
    
    If body1 <> "[]" Then
    
        txt_orderlist.Enabled = True
        txt_orderlist.Clear
        For i = 1 To lenJSON(body1)
            tmpStr = ""
            tmpStr = tmpStr & parseJSON(body1, "XQSLH", i)(0) & "("
            'tmpStr = tmpStr & parseJSON(body1, "FZHZZM", i)(0) & "|"
            'tmpStr = tmpStr & parseJSON(body1, "FHDWMC", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "DZHZZM", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "SHDWMC", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "HZPM", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "CZ", i)(0) & "|"
    
            tmpStr = tmpStr & (CLng(parseJSON(body1, "PZCS", i)(0)) - CLng(parseJSON(body1, "JDZC4", i)(0)) - CLng(parseJSON(body1, "YPWZ", i)(0)) - CLng(parseJSON(body1, "YQWP", i)(0)) - CLng(parseJSON(body1, "FACS", i)(0))) & ")"
    
            txt_orderlist.AddItem tmpStr
        Next
        
        GetOrderNo = "1|获取预定号成功"
        Exit Function
    Else
        txt_orderlist.Clear
        GetOrderNo = "0|没有找到任何预定号"
        Exit Function
    End If
End Function


Function GetInfoByOrderNo(selIndex As Integer, Optional line As String = "online") As String
    On Error Resume Next

    Dim i As Integer, sycs As Long
    
    JsonselIndex = selIndex
    
    i = selIndex + 1
    
    If selIndex >= 0 Then
    
        sycs = (CLng(parseJSON(jsonorder, "PZCS", i)(0)) - CLng(parseJSON(jsonorder, "JDZC4", i)(0)) - CLng(parseJSON(jsonorder, "YPWZ", i)(0)) - CLng(parseJSON(jsonorder, "YQWP", i)(0)) - CLng(parseJSON(jsonorder, "FACS", i)(0)))
        
        txt_xqslh.Text = parseJSON(jsonorder, "XQSLH", i)(0)
        txt_fzhzzm.Text = parseJSON(jsonorder, "FZHZZM", i)(0)
        txt_fjm.Text = parseJSON(jsonorder, "FJQC", i)(0)
        txt_fhdwmc.Text = parseJSON(jsonorder, "FHDWMC", i)(0)
        txt_dzhzzm.Text = parseJSON(jsonorder, "DZHZZM", i)(0)
        txt_djm.Text = parseJSON(jsonorder, "DJQC", i)(0)
        txt_shdwmc.Text = parseJSON(jsonorder, "SHDWMC", i)(0)
        txt_hzpm.Text = parseJSON(jsonorder, "HZPM", i)(0)
        txt_hqhw.Text = parseJSON(jsonorder, "HQHW", i)(0)
        txt_zcrq.Text = Trim(txt_zyrq.Text)
        
        txt_qqcs.Text = 1
        txt_qqds.Text = txt_qqcs.Text * 60
        txt_qqcsMax.Text = sycs
        
        txt_pzycfh.Text = parseJSON(jsonorder, "PZYCFH", i)(0)
        txt_dztmism.Text = parseJSON(jsonorder, "DZTMISM", i)(0)
        txt_fztmism.Text = parseJSON(jsonorder, "FZTMISM", i)(0)
        
        
        
        If (parseJSON(jsonorder, "IFZZJG", i)(0)) = 1 Then chk_ifzzjg.Value = 1
        
        
        s = parseJSON(jsonorder, "CZ", i)(0)
        For ii = 0 To cbo_qqcz.ListCount
            If InStr(cbo_qqcz.List(ii), s) Then
                cbo_qqcz.ListIndex = ii
                Exit For
            End If
        Next

        
        If line = "online" Then '在线获取 离线直接加载内存里的jsonorder2
        
            surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_getZyxByPzycfh"
            param = "pzycfh=" & parseJSON(jsonorder, "PZYCFH", i)(0)
            http.Open "POST", surl, False
            http.SetRequestHeader "Connection", "Keep-Alive"
            http.SetRequestHeader "User-Agent", "Mozilla/4.0"
            http.SetRequestHeader "Cache-Control", "no-cache"
            http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
            http.SetRequestHeader "Accept", "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*"
            http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
            http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
            http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"
            http.SetRequestHeader "Content-Length", Len(param)
            http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            http.Send param
            
            If Err.Number <> 0 Then
                Err.Clear
                GetInfoByOrderNo = "0|根据预定号获取详细信息超时"
                Exit Function
            End If
            
            body1 = BytesToBstr(http.ResponseBody, 2)
            jsonorder2 = body1
            
            Call SavePage(body1, "jsonorder")
        
        End If
    
        txt_zcdd.Text = parseJSON(jsonorder2, "zcdd", 1)(0)
        txt_xcdd.Text = parseJSON(jsonorder2, "xcdd", 1)(0)
        txt_dzyx.Text = parseJSON(jsonorder2, "xcdddm", 1)(0)
        txt_fzyx.Text = parseJSON(jsonorder2, "zcdddm", 1)(0)
        
        If parseJSON(jsonorder2, "shdwdh", 1)(0) <> "" Then
            chk_dddxtz.Value = 1
            txt_shdwdh.Text = parseJSON(jsonorder2, "shdwdh", 1)(0)
        Else
           chk_dddxtz.Value = 0
           txt_shdwdh.Text = ""
        End If
            
        GetInfoByOrderNo = "1|订单填写完成"
        
    End If

End Function

'预提交
Function PerPost() As String

    On Error Resume Next

    Dim surl As String, param As String
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_add"
    param = ""
    param = param & "currentPosition=" & "%E9%A2%84%E7%BA%A6%C2%A0%3E%3E%C2%A0%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    param = param & "&" & "djm=" & URLEncodeUTF8(txt_djm.Text)
    param = param & "&" & "dzhzzm=" & URLEncodeUTF8(txt_dzhzzm.Text)
    param = param & "&" & "dztmism=" & txt_dztmism.Text
    param = param & "&" & "dzyx=" & Replace(txt_dzyx.Text, " ", "+")
    param = param & "&" & "fhdwmc=" & URLEncodeUTF8(txt_fhdwmc.Text)
    param = param & "&" & "fjm=" & URLEncodeUTF8(txt_fjm.Text)
    param = param & "&" & "fzhzzm=" & URLEncodeUTF8(txt_fzhzzm.Text)
    param = param & "&" & "fztmism=" & txt_fztmism.Text
    param = param & "&" & "fzyx=" & Replace(txt_fzyx.Text, " ", "+")
    param = param & "&" & "hzpm=" & URLEncodeUTF8(txt_hzpm.Text)
    param = param & "&" & "keyword="
    param = param & "&" & "maxDate=" & Format(DateAdd("m", 1, Now()) - 1, "yyyy-mm-dd")
    param = param & "&" & "minDate=" & Format(Now(), "yyyy-mm-dd")
    param = param & "&" & "po.dddxtz=" & chk_dddxtz.Value
    param = param & "&" & "po.hqhw=" & txt_hqhw.Text
    param = param & "&" & "po.pzycfh=" & txt_pzycfh.Text
    param = param & "&" & "po.qqcs=" & txt_qqcs.Text
    param = param & "&" & "po.qqcz=" & Right(cbo_qqcz.Text, 1)
    param = param & "&" & "po.qqds=" & txt_qqds.Text
    param = param & "&" & "po.qqlx=0"
    param = param & "&" & "po.shdwdh=" & txt_shdwdh.Text
    param = param & "&" & "po.uuid="
    param = param & "&" & "po.xqslh=" & txt_xqslh.Text
    param = param & "&" & "po.zcrq=" & Trim(txt_zcrq.Text)
    param = param & "&" & "qqcsMax=" & txt_qqcsMax.Text
    param = param & "&" & "shdwmc=" & URLEncodeUTF8(txt_shdwmc.Text)
    param = param & "&" & "xcdd=" & URLEncodeUTF8(txt_xcdd.Text)
    param = param & "&" & "zcdd=" & URLEncodeUTF8(txt_zcdd.Text)
    
    
    Call SavePage("[" & Now() & ":step1]" & param & vbLf, "perpostdata")
    
    http.Open "POST", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        PerPost = "0|订单预提交超时"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage("[" & Now() & ":step1]" & body1 & vbLf, "pergetdata")
    
    If InStr(body1, """success"":true") Then
        uuid = mySubstr(body1, "uuid"":""", """")
        If uuid <> "" Then
            PerPost = "1|预提报成功"
            Exit Function
        Else
            PerPost = "0|获取uuid失败"
            Exit Function
        End If
    ElseIf InStr(body1, "超出可预订日期范围") Then
        PerPost = "0|超出可预订日期范围"
        Exit Function
    ElseIf InStr(body1, "未找到对应的需求信息") Then
        PerPost = "0|未找到对应的需求信息"
        Exit Function
    ElseIf InStr(body1, "提报请求车数不能超过最大可用车数") Then
        PerPost = "0|提报请求车数不能超过最大可用车数"
        Exit Function
    ElseIf InStr(body1, "正在登录，请稍等") Then
        PerPost = "0|已丢失登录"
        Exit Function
    Else
        PerPost = "0|预提报失败"
        Exit Function
    End If


End Function


'正式提报
Function RePost() As String

    On Error Resume Next

    Dim surl As String, param As String
    
    param = "op=10&uuids=" & uuid & ",&mor_dzsw_security_info=mor_dzsw_security_disabled"
    Call SavePage("[" & Now() & ":step2]" & param & vbLf, "perpostdata")
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_operateZcrbjh"
    
    http.Open "POST", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        RePost = "0|订单正式提报超时"
        Exit Function
    End If
    
    body2 = BytesToBstr(http.ResponseBody, 2)
    Call SavePage("[" & Now() & ":step2]" & body2 & vbLf, "pergetdata")
    
    If InStr(body2, """success"":true") Then
        RePost = "1|订单正式提报成功"
        Exit Function
    Else
        RePost = "0|订单正式提报失败"
        Exit Function
    End If

End Function



'**************************************************辅助函数区*********************************************************

Private Sub txt_qqcs_Change()
    If txt_qqcs.Text <> "" And IsNumeric(txt_qqcs.Text) = True Then
        txt_qqds.Text = txt_qqcs.Text * 60
    End If
End Sub

Sub showinfo(Result As Integer, info As String)
    If Result = 1 Then   '成功
        lblInfo.ForeColor = &HD000&
        lblInfo.Caption = info
    ElseIf Result = 2 Then '失败
        lblInfo.ForeColor = &HFF&
        lblInfo.Caption = info
    ElseIf Result = 0 Then '处理中
        lblInfo.ForeColor = &HFFFF&
        lblInfo.Caption = info
    ElseIf Result = 3 Then '提示信息
        lblInfo.ForeColor = &HC00000
        lblInfo.Caption = info
    End If
    
    Form1.Refresh
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        txt_zyrq.Text = Trim(Format(Now(), "yyyy-mm-dd"))
    Else
        txt_zyrq.Text = Trim(Format(DateAdd("d", 31, Now()), "yyyy-mm-dd"))
    End If
    
End Sub

'保存账号
Sub saveAccount(user As String, pass As String)

    Dim tout As String, tin As String, flag As Boolean
    tou = ""
    tin = ""
    flag = False
    
    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(App.Path & "/dat/dat000.dat") = False Then
        Fso.CreateTextFile (App.Path & "/dat/dat000.dat")
    End If
    
    
    Open App.Path & "/dat/dat000.dat" For Input As #1
        Do While Not EOF(1)
            Line Input #1, tin
            If mySubstr(tin, "u=", ";") = user Then
               tout = tout & "u=" & user & ";p=" & pass & ";" & Chr(13) & Chr(10)
               flag = True
            Else
               If Len(tin) > 4 Then tout = tout & tin & vbCrLf
            End If
        Loop
    Close #1
    
    If flag = False Then
        tout = tout & "u=" & user & ";p=" & pass & ";" & vbCrLf
    End If
    
    Open App.Path & "/dat/dat000.dat" For Output As #1
        Print #1, tout;
    Close #1
    
End Sub


'读取账号
Sub bindAccount(Optional user As String = "")

    Dim tout As String, tin As String, flag As Boolean
    tou = ""
    tin = ""
    flag = False
    
    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(App.Path & "/dat/dat000.dat") = False Then
        Exit Sub
    End If
    
    If user = "" Then
    
        Open App.Path & "/dat/dat000.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, tin
                If Len(tin) > 4 Then
                   txtUsername.AddItem (mySubstr(tin, "u=", ";"))
                End If
            Loop
        Close #1
        txtUsername.ListIndex = txtUsername.ListCount - 1
        
    Else
    
        Open App.Path & "/dat/dat000.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, tin
                If Len(tin) > 4 And mySubstr(tin, "u=", ";") = user Then
                   txtPassWord.Text = mySubstr(tin, "p=", ";")
                End If
            Loop
        Close #1
    
    End If
    
End Sub


Sub saveProfile(filename As String)

    Dim tout As String
    
    filename = App.Path & "/dat/" & filename & ".dat"

    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(filename) = False Then
        Fso.CreateTextFile (filename)
    End If
    
    tout = ""
    tout = tout & "user=" & Trim(txtUsername.Text) & "" & vbCrLf
    tout = tout & "pass=" & Trim(txtPassWord.Text) & "" & vbCrLf
    tout = tout & "comp=" & Mid(Label6.Caption, 5, Len(Label6.Caption) - 5) & "" & vbCrLf
    tout = tout & "zcrq=" & Trim(txt_zcrq.Text) & "" & vbCrLf
    tout = tout & "jsel=" & JsonselIndex & vbCrLf
    tout = tout & "jod1=" & jsonorder & "" & vbCrLf
    tout = tout & "jod2=" & jsonorder2 & "" & vbCrLf
   

    Open filename For Output As #1
        Print #1, tout;
    Close #1
    
End Sub


Sub loadProfile(filename As String)

    Dim tout As String, tin As String, tmpStr As String
    
    filename = App.Path & "/dat/" & filename & ".dat"

    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(filename) = False Then
        Call showinfo(2, "没有找到对应的订单文件,载入失败!")
        Exit Sub
    End If
    
    

    Open filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, tin
            
            If Left(tin, 4) = "user" Then
                txtUsername.Text = Right(tin, Len(tin) - 5)
                ISOFFLINE = True
                
            ElseIf Left(tin, 4) = "pass" Then
                txtPassWord.Text = Right(tin, Len(tin) - 5)
                
            ElseIf Left(tin, 4) = "comp" Then
                Label6.ForeColor = RGB(0, 0, 255)
                Label6.Caption = "离线订单(" & Right(tin, Len(tin) - 5) & ")"
                
            ElseIf Left(tin, 4) = "zcrq" Then
                txt_zyrq.Text = Right(tin, Len(tin) - 5)
                
            ElseIf Left(tin, 4) = "jsel" Then
                JsonselIndex = Right(tin, Len(tin) - 5)
                
            ElseIf Left(tin, 4) = "jod1" Then
    
                jsonorder = Right(tin, Len(tin) - 5)
                
                txt_orderlist.Clear
               
                tmpStr = ""
                tmpStr = tmpStr & parseJSON(jsonorder, "XQSLH", JsonselIndex + 1)(0) & "("
                tmpStr = tmpStr & parseJSON(jsonorder, "DZHZZM", JsonselIndex + 1)(0) & "|"
                tmpStr = tmpStr & parseJSON(jsonorder, "SHDWMC", JsonselIndex + 1)(0) & "|"
                tmpStr = tmpStr & parseJSON(jsonorder, "HZPM", JsonselIndex + 1)(0) & "|"
                tmpStr = tmpStr & parseJSON(jsonorder, "CZ", JsonselIndex + 1)(0) & "|"
    
                tmpStr = tmpStr & (CLng(parseJSON(jsonorder, "PZCS", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "JDZC4", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "YPWZ", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "YQWP", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "FACS", JsonselIndex + 1)(0))) & ")"
            
                txt_orderlist.AddItem tmpStr
                txt_orderlist.Locked = True
                txt_orderlist.ListIndex = 0
                txt_orderlist.Enabled = False
                
            ElseIf Left(tin, 4) = "jod2" Then
            
                jsonorder2 = Right(tin, Len(tin) - 5)
                Call GetInfoByOrderNo(JsonselIndex, "offline")
                
                Call expandOrder
                
            Else
                
            End If
        Loop
        
    Close #1
    
End Sub


Sub lockAll()
    cmdAuto.Enabled = False
    cmdDeAuto.Enabled = True
    
    cmd_login.Enabled = False
    cmd_getorder.Enabled = False
    
    txt_orderlist.Enabled = False
    txt_zyrq.Enabled = False
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    cmd_manual.Enabled = False
    Form1.Width = 5595
    
    cmd_profile.Enabled = False
    txt_profile.Enabled = False
    
    txt_AllowAuto.Enabled = False
    
    
End Sub

Sub unlockAll()

    cmdAuto.Enabled = True
    cmdDeAuto.Enabled = False
    
    cmd_login.Enabled = True
    cmd_getorder.Enabled = True
    cmd_manual.Enabled = True
    
    txt_orderlist.Enabled = True
    txt_zyrq.Enabled = True
    
    Option1(0).Enabled = True
    Option1(1).Enabled = True
    
    Form1.Width = 5595
    
    cmd_profile.Enabled = True
    txt_profile.Enabled = True
    
    txt_AllowAuto.Enabled = True
    
End Sub


Sub expandOrder() '展开
    Form1.Width = 14220
    cmd_expandOrder.Enabled = False
End Sub

Sub contractOrder() '收起
    Form1.Width = 5595
    cmd_expandOrder.Enabled = True
End Sub


Private Sub cmd_expandOrder_Click()
    Call expandOrder
End Sub

Private Sub cmd_contractOrder_Click()
    Call contractOrder
End Sub
