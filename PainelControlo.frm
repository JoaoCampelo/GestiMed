VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PainelControlo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GestiMed"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   9945
   ControlBox      =   0   'False
   Icon            =   "PainelControlo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   5
      Left            =   1680
      TabIndex        =   149
      Top             =   1680
      Width           =   10095
      Begin VB.CommandButton cmdVender 
         Caption         =   "Vender"
         Height          =   375
         Left            =   5400
         TabIndex        =   173
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRepor6 
         Caption         =   "Repor"
         Height          =   375
         Left            =   6960
         TabIndex        =   155
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar6 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6960
         TabIndex        =   154
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair6 
         Caption         =   "Sair"
         Height          =   375
         Left            =   8520
         TabIndex        =   153
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserir6 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   152
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAlterar6 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1560
         TabIndex        =   151
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdDel6 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         MaskColor       =   &H8000000B&
         TabIndex        =   150
         Top             =   3720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabe110 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":521A
         TabIndex        =   156
         Top             =   4320
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListMedicamentos 
         Height          =   3375
         Left            =   240
         TabIndex        =   157
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel111 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":52A0
         TabIndex        =   158
         Top             =   4680
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel116 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":532A
         TabIndex        =   159
         Top             =   6480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel117 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "PainelControlo.frx":539C
         TabIndex        =   160
         Top             =   4320
         Width           =   7455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel118 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "PainelControlo.frx":53FA
         TabIndex        =   161
         Top             =   4680
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel112 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5458
         TabIndex        =   162
         Top             =   5040
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel119 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "PainelControlo.frx":54E0
         TabIndex        =   163
         Top             =   5040
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel123 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "PainelControlo.frx":553E
         TabIndex        =   164
         Top             =   6480
         Width           =   8535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel113 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":559C
         TabIndex        =   165
         Top             =   5400
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel114 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5626
         TabIndex        =   166
         Top             =   5760
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel120 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "PainelControlo.frx":56AC
         TabIndex        =   167
         Top             =   5400
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel121 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "PainelControlo.frx":570A
         TabIndex        =   168
         Top             =   5760
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel115 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5768
         TabIndex        =   169
         Top             =   6120
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel122 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "PainelControlo.frx":57F4
         TabIndex        =   170
         Top             =   6120
         Width           =   2655
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   4
      Left            =   1320
      TabIndex        =   133
      Top             =   1320
      Width           =   10095
      Begin VB.CommandButton cmdDel5 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         MaskColor       =   &H8000000B&
         TabIndex        =   139
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAlterar5 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1560
         TabIndex        =   138
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdInserir5 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   137
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdSair5 
         Caption         =   "Sair"
         Height          =   375
         Left            =   8520
         TabIndex        =   136
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar5 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6960
         TabIndex        =   135
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRepor5 
         Caption         =   "Repor"
         Height          =   375
         Left            =   6960
         TabIndex        =   134
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel102 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5852
         TabIndex        =   140
         Top             =   4320
         Width           =   1935
      End
      Begin MSComctlLib.ListView ListTratamentos 
         Height          =   3375
         Left            =   240
         TabIndex        =   141
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel103 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":58D6
         TabIndex        =   142
         Top             =   4800
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel105 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":595E
         TabIndex        =   143
         Top             =   5760
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel106 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "PainelControlo.frx":59D0
         TabIndex        =   144
         Top             =   4320
         Width           =   7695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel107 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "PainelControlo.frx":5A2E
         TabIndex        =   145
         Top             =   4800
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel104 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5A8C
         TabIndex        =   146
         Top             =   5280
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel108 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "PainelControlo.frx":5B12
         TabIndex        =   147
         Top             =   5280
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel109 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "PainelControlo.frx":5B70
         TabIndex        =   148
         Top             =   5760
         Width           =   8535
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   3
      Left            =   960
      TabIndex        =   99
      Top             =   960
      Width           =   10095
      Begin VB.CommandButton cmdRepor4 
         Caption         =   "Repor"
         Height          =   375
         Left            =   6960
         TabIndex        =   105
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar4 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6960
         TabIndex        =   104
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair4 
         Caption         =   "Sair"
         Height          =   375
         Left            =   8520
         TabIndex        =   103
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserir4 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   102
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAlterar4 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1560
         TabIndex        =   101
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdDel4 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         MaskColor       =   &H8000000B&
         TabIndex        =   100
         Top             =   3720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel75 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5BCE
         TabIndex        =   106
         Top             =   4200
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListMedicos 
         Height          =   3375
         Left            =   240
         TabIndex        =   107
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel87 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":5C4A
         TabIndex        =   108
         Top             =   6360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel86 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5CC2
         TabIndex        =   109
         Top             =   6360
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel80 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5D2A
         TabIndex        =   110
         Top             =   5280
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel83 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5D9A
         TabIndex        =   111
         Top             =   6000
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel84 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":5E06
         TabIndex        =   112
         Top             =   6000
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel77 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":5E80
         TabIndex        =   113
         Top             =   4560
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel78 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":5F04
         TabIndex        =   114
         Top             =   4920
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel79 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":5F74
         TabIndex        =   115
         Top             =   4920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel89 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "PainelControlo.frx":5FE6
         TabIndex        =   116
         Top             =   4200
         Width           =   7935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel91 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "PainelControlo.frx":6044
         TabIndex        =   117
         Top             =   4560
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel101 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "PainelControlo.frx":60A2
         TabIndex        =   118
         Top             =   6360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel100 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "PainelControlo.frx":6100
         TabIndex        =   119
         Top             =   6360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel92 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":615E
         TabIndex        =   120
         Top             =   4920
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel93 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "PainelControlo.frx":61BC
         TabIndex        =   121
         Top             =   4920
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel94 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":621A
         TabIndex        =   122
         Top             =   5280
         Width           =   8535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel97 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":6278
         TabIndex        =   123
         Top             =   6000
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel98 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "PainelControlo.frx":62D6
         TabIndex        =   124
         Top             =   6000
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel76 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":6334
         TabIndex        =   125
         Top             =   4560
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel90 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "PainelControlo.frx":63B4
         TabIndex        =   126
         Top             =   4560
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel81 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":6412
         TabIndex        =   127
         Top             =   5640
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel95 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "PainelControlo.frx":649C
         TabIndex        =   128
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel82 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":64FA
         TabIndex        =   129
         Top             =   5640
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel96 
         Height          =   255
         Left            =   6120
         OleObjectBlob   =   "PainelControlo.frx":6586
         TabIndex        =   130
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel85 
         Height          =   255
         Left            =   7200
         OleObjectBlob   =   "PainelControlo.frx":65E4
         TabIndex        =   131
         Top             =   6000
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel99 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "PainelControlo.frx":6650
         TabIndex        =   132
         Top             =   6000
         Width           =   1815
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   2
      Left            =   600
      TabIndex        =   63
      Top             =   600
      Width           =   10095
      Begin VB.CommandButton cmdDel3 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         MaskColor       =   &H8000000B&
         TabIndex        =   69
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAlterar3 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1560
         TabIndex        =   68
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdInserir3 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdSair3 
         Caption         =   "Sair"
         Height          =   375
         Left            =   8520
         TabIndex        =   66
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar3 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6960
         TabIndex        =   65
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRepor3 
         Caption         =   "Repor"
         Height          =   375
         Left            =   6960
         TabIndex        =   64
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":66AE
         TabIndex        =   70
         Top             =   4200
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListPacientes 
         Height          =   3375
         Left            =   240
         TabIndex        =   71
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":672E
         TabIndex        =   72
         Top             =   6360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":67A6
         TabIndex        =   73
         Top             =   6360
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":680E
         TabIndex        =   74
         Top             =   5280
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel55 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":687E
         TabIndex        =   75
         Top             =   6000
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel56 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":68EA
         TabIndex        =   76
         Top             =   6000
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel49 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":6964
         TabIndex        =   77
         Top             =   4560
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel50 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":69E8
         TabIndex        =   78
         Top             =   4920
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":6A58
         TabIndex        =   79
         Top             =   4920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel61 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "PainelControlo.frx":6ACA
         TabIndex        =   80
         Top             =   4200
         Width           =   7695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel63 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "PainelControlo.frx":6B28
         TabIndex        =   81
         Top             =   4560
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel73 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "PainelControlo.frx":6B86
         TabIndex        =   82
         Top             =   6360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel72 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "PainelControlo.frx":6BE4
         TabIndex        =   83
         Top             =   6360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel64 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":6C42
         TabIndex        =   84
         Top             =   4920
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel65 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "PainelControlo.frx":6CA0
         TabIndex        =   85
         Top             =   4920
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel66 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":6CFE
         TabIndex        =   86
         Top             =   5280
         Width           =   8535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel69 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":6D5C
         TabIndex        =   87
         Top             =   6000
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel70 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "PainelControlo.frx":6DBA
         TabIndex        =   88
         Top             =   6000
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel48 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":6E18
         TabIndex        =   89
         Top             =   4560
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel62 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "PainelControlo.frx":6E9C
         TabIndex        =   90
         Top             =   4560
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":6EFA
         TabIndex        =   91
         Top             =   5640
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel67 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "PainelControlo.frx":6F84
         TabIndex        =   92
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel54 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":6FE2
         TabIndex        =   93
         Top             =   5640
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel68 
         Height          =   255
         Left            =   6120
         OleObjectBlob   =   "PainelControlo.frx":706E
         TabIndex        =   94
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel57 
         Height          =   255
         Left            =   7200
         OleObjectBlob   =   "PainelControlo.frx":70CC
         TabIndex        =   95
         Top             =   6000
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel71 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "PainelControlo.frx":7138
         TabIndex        =   96
         Top             =   6000
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7196
         TabIndex        =   97
         Top             =   6720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel74 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "PainelControlo.frx":7204
         TabIndex        =   98
         Top             =   6720
         Width           =   8535
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   1
      Left            =   240
      TabIndex        =   29
      Top             =   240
      Width           =   10095
      Begin VB.CommandButton cmdRepor2 
         Caption         =   "Repor"
         Height          =   375
         Left            =   6960
         TabIndex        =   35
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar2 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6960
         TabIndex        =   34
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair2 
         Caption         =   "Sair"
         Height          =   375
         Left            =   8520
         TabIndex        =   33
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserir2 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAlterar2 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdDel2 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         MaskColor       =   &H8000000B&
         TabIndex        =   30
         Top             =   3720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7262
         TabIndex        =   36
         Top             =   4200
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListConsultasTratamento 
         Height          =   3375
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":72E6
         TabIndex        =   38
         Top             =   5280
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7366
         TabIndex        =   39
         Top             =   5640
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabe31 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":73EA
         TabIndex        =   40
         Top             =   6000
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":745A
         TabIndex        =   41
         Top             =   6360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":74D6
         TabIndex        =   42
         Top             =   6720
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":7556
         TabIndex        =   43
         Top             =   4200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":75BE
         TabIndex        =   44
         Top             =   5640
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabe23 
         Height          =   255
         Left            =   6720
         OleObjectBlob   =   "PainelControlo.frx":762E
         TabIndex        =   45
         Top             =   4200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
         Height          =   255
         Left            =   6720
         OleObjectBlob   =   "PainelControlo.frx":7696
         TabIndex        =   46
         Top             =   5640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "PainelControlo.frx":7708
         TabIndex        =   47
         Top             =   4200
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "PainelControlo.frx":7766
         TabIndex        =   48
         Top             =   4200
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "PainelControlo.frx":77C4
         TabIndex        =   49
         Top             =   4200
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "PainelControlo.frx":7822
         TabIndex        =   50
         Top             =   5280
         Width           =   7215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "PainelControlo.frx":7880
         TabIndex        =   51
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "PainelControlo.frx":78DE
         TabIndex        =   52
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
         Height          =   255
         Left            =   7800
         OleObjectBlob   =   "PainelControlo.frx":793C
         TabIndex        =   53
         Top             =   5640
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":799A
         TabIndex        =   54
         Top             =   6000
         Width           =   8055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "PainelControlo.frx":79F8
         TabIndex        =   55
         Top             =   6360
         Width           =   7215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "PainelControlo.frx":7A56
         TabIndex        =   56
         Top             =   6720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7AB4
         TabIndex        =   57
         Top             =   4560
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "PainelControlo.frx":7B38
         TabIndex        =   58
         Top             =   4560
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7B96
         TabIndex        =   59
         Top             =   4920
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "PainelControlo.frx":7C1E
         TabIndex        =   60
         Top             =   4920
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":7C7C
         TabIndex        =   61
         Top             =   4920
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "PainelControlo.frx":7D02
         TabIndex        =   62
         Top             =   4920
         Width           =   1575
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   0
      Left            =   -120
      TabIndex        =   1
      Top             =   -120
      Width           =   10095
      Begin VB.CommandButton cmdDel 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2760
         MaskColor       =   &H8000000B&
         TabIndex        =   8
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   8520
         TabIndex        =   4
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRepor 
         Caption         =   "Repor"
         Height          =   375
         Left            =   6960
         TabIndex        =   2
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7D60
         TabIndex        =   5
         Top             =   4320
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListConsultas 
         Height          =   3375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7DE4
         TabIndex        =   10
         Top             =   4800
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7E64
         TabIndex        =   11
         Top             =   5280
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7EE8
         TabIndex        =   12
         Top             =   5760
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7F58
         TabIndex        =   13
         Top             =   6240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "PainelControlo.frx":7FD4
         TabIndex        =   14
         Top             =   6720
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":8054
         TabIndex        =   15
         Top             =   4320
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "PainelControlo.frx":80BC
         TabIndex        =   16
         Top             =   5280
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   6720
         OleObjectBlob   =   "PainelControlo.frx":812C
         TabIndex        =   17
         Top             =   4320
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   6720
         OleObjectBlob   =   "PainelControlo.frx":8194
         TabIndex        =   18
         Top             =   5280
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "PainelControlo.frx":8206
         TabIndex        =   19
         Top             =   4320
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "PainelControlo.frx":8264
         TabIndex        =   20
         Top             =   4320
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "PainelControlo.frx":82C2
         TabIndex        =   21
         Top             =   4320
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "PainelControlo.frx":8320
         TabIndex        =   22
         Top             =   4800
         Width           =   7215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "PainelControlo.frx":837E
         TabIndex        =   23
         Top             =   5280
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "PainelControlo.frx":83DC
         TabIndex        =   24
         Top             =   5280
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   7800
         OleObjectBlob   =   "PainelControlo.frx":843A
         TabIndex        =   25
         Top             =   5280
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "PainelControlo.frx":8498
         TabIndex        =   26
         Top             =   5760
         Width           =   8055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "PainelControlo.frx":84F6
         TabIndex        =   27
         Top             =   6240
         Width           =   7215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "PainelControlo.frx":8554
         TabIndex        =   28
         Top             =   6720
         Width           =   1575
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3360
      OleObjectBlob   =   "PainelControlo.frx":85B2
      Top             =   -240
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   -240
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12938
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      MultiSelect     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultas de Tratamentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pacientes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Médicos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tratamentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Medicamentos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   375
      Left            =   1440
      TabIndex        =   172
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   0
      TabIndex        =   171
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu ficheiro 
      Caption         =   "Ficheiro"
      Begin VB.Menu bloquear 
         Caption         =   "Bloquear"
      End
      Begin VB.Menu sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu apcoes 
      Caption         =   "Opções"
      Begin VB.Menu verfacturas 
         Caption         =   "Ver Facturas"
      End
      Begin VB.Menu seguranca 
         Caption         =   "Segurança"
         Begin VB.Menu novo_utilizador 
            Caption         =   "Novo Utilizador"
         End
         Begin VB.Menu alterar_password 
            Caption         =   "Alterar Password"
         End
      End
   End
   Begin VB.Menu ajuda 
      Caption         =   "Ajuda"
      Begin VB.Menu topicosajuda 
         Caption         =   "Topicos de Ajuda"
      End
      Begin VB.Menu sobre 
         Caption         =   "Sobre o GestiMed"
      End
   End
End
Attribute VB_Name = "PainelControlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String
Private SelectedTab As Integer

Private Sub alterar_password_Click()
    
    AlterarPassword.Show
    AlterarPassword.Caption = "Alterar Password"
    PainelControlo.Enabled = False
    
End Sub

Private Sub bloquear_Click()
    
    Desbloquear.Show
    PainelControlo.Enabled = False
    
End Sub

Private Sub sobre_Click()

    SobreGestiMed.Show
    PainelControlo.Enabled = False

End Sub

Private Sub topicosajuda_Click()

    frmAjuda.Show
    PainelControlo.Enabled = False
    
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
        For i = 1 To ChoiceFrame.UBound
            ChoiceFrame(i).Move _
            ChoiceFrame(0).Left, _
            ChoiceFrame(0).Top, _
            ChoiceFrame(0).Width, _
            ChoiceFrame(0).Height
        ChoiceFrame(i).Visible = False
    Next i
    
    SelectedTab = 1
    TabStrip1.SelectedItem = TabStrip1.Tabs(SelectedTab)
    ChoiceFrame(SelectedTab - 1).Visible = True

    Dim itemx As ListItem
    Dim str As String
    
    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
    ListConsultas_SetUp
    Preencher_dados
    
    ListConsultasTratamento_SetUp
    Preencher_dados_tratamentos
    
    ListPacientes_SetUp
    Preencher_dados_Paceintes
    
    ListMedicos_SetUp
    Preencher_dados_Medicos
    
    ListTratamentos_SetUp
    Preencher_dados_tratamento
    
    ListMedicamentos_SetUp
    Preencher_dados_Medicamentos
    
    SQL = "select * from consultas"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        cmdAlterar.Enabled = False
        cmdDel.Enabled = False
    End If
    
    SQL = "select * from consultas_tratamentos"
    Set rst = db.OpenRecordset(SQL)
     If rst.BOF = True And rst.EOF = True Then
        cmdAlterar2.Enabled = False
        cmdDel2.Enabled = False
    End If
    
    SQL = "select * from pacientes"
    Set rst = db.OpenRecordset(SQL)
     If rst.BOF = True And rst.EOF = True Then
        cmdAlterar3.Enabled = False
        cmdDel3.Enabled = False
    End If
    
    SQL = "select * from medicos"
    Set rst = db.OpenRecordset(SQL)
     If rst.BOF = True And rst.EOF = True Then
        cmdAlterar4.Enabled = False
        cmdDel4.Enabled = False
    End If
    
    SQL = "select * from tratamentos"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        cmdAlterar5.Enabled = False
        cmdDel5.Enabled = False
    End If
    
    SQL = "select * from medicamentos"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        cmdAlterar6.Enabled = False
        cmdDel6.Enabled = False
        cmdVender.Enabled = False
    End If
    
End Sub

Private Sub novo_utilizador_Click()
    
    NovoUtilizador.Show
    NovoUtilizador.Caption = "Novo Utilizador (GestiMed)"
    PainelControlo.Enabled = False
    
End Sub

Private Sub verfacturas_Click()

    Facturas.Show
    PainelControlo.Enabled = False

End Sub

Private Sub sair_Click()
    
    End
    
End Sub

Private Sub TabStrip1_Click()

    ChoiceFrame(SelectedTab - 1).Visible = False
    SelectedTab = TabStrip1.SelectedItem.Index
    ChoiceFrame(SelectedTab - 1).Visible = True
    
End Sub

Private Sub Timer1_Timer()

    If PainelControlo.Enabled = True Then
        Listconsultas_Click
        ListConsultasTratamentos_Click
        ListPaceintes_Click
        ListMedicos_Click
        ListTratamentos_Click
        ListMedicamentos_Click
    End If
    Label19.Caption = Time
    Label18.Caption = Date
    
End Sub

'-------------------------------------------Código para o separador "Consultas"------------------------------------------

Private Sub cmdAlterar_Click()
    
    SQL = "SELECT * FROM Consultas, seg_Pacientes, seg_Medicos" _
        & " WHERE consultas.cod_cunsulta=" & Val(PainelControlo.ListConsultas.SelectedItem.Tag) _
        & " AND seg_medicos.cod_medico=" & Val(SkinLabel20.Caption) _
        & " AND seg_pacientes.cod_paciente=" & Val(SkinLabel15.Caption)
        
    Set rst = db.OpenRecordset(SQL)
    
    If rst("Data") > DateValue(Label18.Caption) Then
    
        InserirAlterarConsultas.Text1.Text = rst("cod_cunsulta")
        InserirAlterarConsultas.Text2.Text = rst("data")
        InserirAlterarConsultas.Text3.Text = rst("hora")
        InserirAlterarConsultas.cmbPacientes.Text = rst("nome_paciente")
        InserirAlterarConsultas.Text5.Text = rst("seg_pacientes.cod_paciente")
        
        If rst("seg_pacientes.telefone") = Null Then
            InserirAlterarConsultas.Text6.Text = ""
        Else
            x = rst("seg_pacientes.telefone")
            InserirAlterarConsultas.Text6.Text = "" & x
        End If
    
        If rst("seg_pacientes.telemovel") = Null Then
            InserirAlterarConsultas.Text7.Text = ""
        Else
            x = rst("seg_pacientes.telemovel")
            InserirAlterarConsultas.Text7.Text = "" & x
        End If
    
        If rst("seg_pacientes.email") = Null Then
            InserirAlterarConsultas.Text10.Text = ""
        Else
            x = rst("seg_pacientes.email")
            InserirAlterarConsultas.Text10.Text = "" & x
        End If
        
        InserirAlterarConsultas.cmbMedicos.Text = rst("nome_medico")
        InserirAlterarConsultas.Text9.Text = rst("seg_medicos.cod_medico")
        InserirAlterarConsultas.Show
        InserirAlterarConsultas.Caption = "Alterar Dados da Consulta"
        InserirAlterarConsultas.cmdInserir.Caption = "Alterar"
        
        SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
        Set rst = db.OpenRecordset(SQL)
        With rst
            If Not rst.EOF Then
                Do While Not rst.EOF
                    InserirAlterarConsultas.cmbPacientes.AddItem rst("nome_paciente")
                    InserirAlterarConsultas.cmbPacientes.ItemData(InserirAlterarConsultas.cmbPacientes.NewIndex) = rst("cod_Paciente")
                    rst.MoveNext
                Loop
            Else
                InserirAlterarConsultas.cmbPacientes.Text = ""
            End If
        End With
        
        SQL = " SELECT * FROM Medicos ORDER BY nome_medico"
        Set rst = db.OpenRecordset(SQL)
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    InserirAlterarConsultas.cmbMedicos.AddItem rst("nome_medico")
                    InserirAlterarConsultas.cmbMedicos.ItemData(InserirAlterarConsultas.cmbMedicos.NewIndex) = rst("cod_medico")
                    .MoveNext
                Loop
            Else
                InserirAlterarConsultas.cmbMedicos.Text = ""
            End If
        End With
        
    Else
        If rst("Data") = DateValue(Label18.Caption) Then
            If rst("hora") > TimeValue(Label19.Caption) Then
                InserirAlterarConsultas.Text1.Text = rst("cod_cunsulta")
                InserirAlterarConsultas.Text2.Text = rst("data")
                InserirAlterarConsultas.Text3.Text = rst("hora")
                InserirAlterarConsultas.cmbPacientes.Text = rst("nome_paciente")
                InserirAlterarConsultas.Text5.Text = rst("seg_pacientes.cod_paciente")
                
                If rst("seg_pacientes.telefone") = Null Then
                    InserirAlterarConsultas.Text6.Text = ""
                Else
                    x = rst("seg_pacientes.telefone")
                    InserirAlterarConsultas.Text6.Text = "" & x
                End If
    
                If rst("seg_pacientes.telemovel") = Null Then
                    InserirAlterarConsultas.Text7.Text = ""
                Else
                    x = rst("seg_pacientes.telemovel")
                    InserirAlterarConsultas.Text7.Text = "" & x
                End If
    
                If rst("seg_pacientes.email") = Null Then
                    InserirAlterarConsultas.Text10.Text = ""
                Else
                    x = rst("seg_pacientes.email")
                    InserirAlterarConsultas.Text10.Text = "" & x
                End If
                
                InserirAlterarConsultas.cmbMedicos.Text = rst("nome_medico")
                InserirAlterarConsultas.Text9.Text = rst("seg_medicos.cod_medico")
                InserirAlterarConsultas.Show
                InserirAlterarConsultas.Caption = "Alterar Dados da Consulta"
                InserirAlterarConsultas.cmdInserir.Caption = "Alterar"
                
                SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
                Set rst = db.OpenRecordset(SQL)
                With rst
                    If Not rst.EOF Then
                        Do While Not rst.EOF
                            InserirAlterarConsultas.cmbPacientes.AddItem rst("nome_paciente")
                            InserirAlterarConsultas.cmbPacientes.ItemData(InserirAlterarConsultas.cmbPacientes.NewIndex) = rst("cod_Paciente")
                            rst.MoveNext
                        Loop
                    Else
                        InserirAlterarConsultas.cmbPacientes.Text = ""
                    End If
                End With
                
                SQL = " SELECT * FROM Medicos ORDER BY nome_medico"
                Set rst = db.OpenRecordset(SQL)
                With rst
                    If Not .EOF Then
                        Do While Not .EOF
                            InserirAlterarConsultas.cmbMedicos.AddItem rst("nome_medico")
                            InserirAlterarConsultas.cmbMedicos.ItemData(InserirAlterarConsultas.cmbMedicos.NewIndex) = rst("cod_medico")
                            .MoveNext
                        Loop
                    Else
                        InserirAlterarConsultas.cmbMedicos.Text = ""
                    End If
                End With
                Exit Sub
            
            Else
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser alterada porque já foi realizada!"
                FormMsgBoxNormal.Caption = "Alterar (Consultas)"
                PainelControlo.Enabled = False
                Exit Sub
            End If
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser alterada porque já foi realizada!"
            FormMsgBoxNormal.Caption = "Alterar (Consultas)"
            PainelControlo.Enabled = False
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser alterada porque já foi realizada!"
            FormMsgBoxNormal.Caption = "Alterar (Consultas)"
            PainelControlo.Enabled = False
        End If
    End If
    
End Sub

Private Sub cmdDel_Click()

    FormMsgBoxSimNao.Show
    FormMsgBoxSimNao.Label1.Caption = "Tem a certeza que quer eliminar esta consulta?"
    FormMsgBoxSimNao.Caption = "Eliminar (Consultas)"
    PainelControlo.Enabled = False
    
End Sub

Private Sub cmdInserir_Click()

    PainelControlo.Enabled = False
    InserirAlterarConsultas.Show
    InserirAlterarConsultas.Caption = "Marcar Nova Consulta"
    InserirAlterarConsultas.cmdInserir.Caption = "&Inserir"
    InserirAlterarConsultas.cmbPacientes.Enabled = True
    InserirAlterarConsultas.cmbMedicos.Enabled = True
    
    SQL = " SELECT * FROM Consultas"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        InserirAlterarConsultas.Text1.Text = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_cunsulta")
                    .MoveNext
                Loop
            Else
                y = rst("cod_cunsulta")
            End If
        End With
        res = x
        InserirAlterarConsultas.Text1.Text = res + 1
    End If

    SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not rst.EOF Then
            Do While Not rst.EOF
                InserirAlterarConsultas.cmbPacientes.AddItem rst("nome_paciente")
                rst.MoveNext
            Loop
        Else
            InserirAlterarConsultas.cmbPacientes.Text = ""
        End If
    End With
    
    SQL = " SELECT * FROM Medicos ORDER BY nome_medico"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not .EOF Then
            Do While Not .EOF
                InserirAlterarConsultas.cmbMedicos.AddItem rst("nome_medico")
                .MoveNext
            Loop
        Else
            InserirAlterarConsultas.cmbMedicos.Text = ""
        End If
    End With
    
End Sub

Private Sub cmdProcurar_Click()

    InputBoxProcurar.Caption = "Procurar (Consultas)"
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor_Click()

    Dim itemx As ListItem

    SQL = " SELECT consultas.cod_cunsulta, consultas.data, consultas.hora, seg_pacientes.nome_paciente " _
        & " FROM consultas, seg_pacientes" _
        & " WHERE consultas.cod_paciente=seg_pacientes.cod_paciente" _
        & " ORDER BY data desc"
        
    Set rst = db.OpenRecordset(SQL)
    ListConsultas.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListConsultas.ListItems.Add(, , rst("cod_cunsulta"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("nome_paciente")
        itemx.Tag = rst("cod_cunsulta")
        
        rst.MoveNext
    Wend
    
    cmdProcurar.Visible = True
    cmdRepor.Visible = False
    
End Sub

Private Sub cmdSair_Click()

    End
    
End Sub

Private Sub ListConsultas_SetUp()

    Dim itemx As ListItem
    
    SQL = " SELECT consultas.cod_cunsulta, consultas.data, consultas.hora, seg_pacientes.nome_paciente " _
        & " FROM consultas, seg_pacientes" _
        & " WHERE consultas.cod_paciente=seg_pacientes.cod_paciente" _
        & " ORDER BY data desc"
        
    Set rst = db.OpenRecordset(SQL)
     
    ListConsultas.ListItems.Clear
    ListConsultas.ColumnHeaders.Add , , "Código da Consulta", ListConsultas.Width - 8000
    ListConsultas.ColumnHeaders.Add , , "Data", ListConsultas.Width - 8000
    ListConsultas.ColumnHeaders.Add , , "Hora", ListConsultas.Width - 8000
    ListConsultas.ColumnHeaders.Add , , "Nome do Paciente", ListConsultas.Width - 5300
    ListConsultas.View = lvwReport
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListConsultas.ListItems.Add(, , rst("cod_cunsulta"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("nome_paciente")
        itemx.Tag = rst("cod_cunsulta")
        
        rst.MoveNext
    Wend

End Sub

Private Sub Listconsultas_Click()
 
    Preencher_dados
    
End Sub

Private Sub Preencher_dados()

    SQL = "select * from consultas"
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
    SQL = " SELECT consultas.cod_cunsulta, consultas.data, consultas.hora, seg_pacientes.cod_paciente, seg_pacientes.nome_paciente, seg_pacientes.telefone, seg_pacientes.telemovel, seg_pacientes.email, seg_medicos.cod_medico, seg_medicos.nome_medico" _
        & " FROM consultas, seg_pacientes, seg_medicos" _
        & " WHERE consultas.cod_paciente = seg_pacientes.cod_paciente" _
        & " AND consultas.cod_medico = seg_medicos.cod_medico" _
        & " AND cod_cunsulta=" & ListConsultas.SelectedItem.Tag _
        & " ORDER BY data desc"
        
    Set rst = db.OpenRecordset(SQL)
        
    SkinLabel11.Caption = rst("cod_cunsulta")
    SkinLabel12.Caption = rst("Data")
    SkinLabel13.Caption = rst("Hora")
    SkinLabel14.Caption = rst("nome_paciente")
    SkinLabel15.Caption = rst("cod_paciente")
    
    If rst("telefone") = Null Then
        SkinLabel16.Caption = ""
    Else
        x = rst("telefone")
        SkinLabel16.Caption = "" & x
    End If
    
    If rst("telemovel") = Null Then
        SkinLabel17.Caption = ""
    Else
        x = rst("telemovel")
        SkinLabel17.Caption = "" & x
    End If
    
    If rst("email") = Null Then
        SkinLabel18.Caption = ""
    Else
        x = rst("email")
        SkinLabel18.Caption = "" & x
    End If
    
    SkinLabel19.Caption = rst("nome_medico")
    SkinLabel20.Caption = rst("cod_medico")
    
    End If
    
End Sub
Private Sub ListConsultas_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Preencher_dados
    
End Sub
'----------------------------------------Fim código para o separador "Consultas"------------------------------------------

'-------------------------------------------Código para o separador "Consultas de Tratamento"-----------------------------
Private Sub cmdAlterar2_Click()
    
    SQL = "SELECT * FROM Consultas_Tratamentos, seg_Pacientes, seg_Medicos, seg_Tratamentos" _
        & " WHERE consultas_tratamentos.cod_consultatratamento=" & Val(PainelControlo.ListConsultasTratamento.SelectedItem.Tag) _
        & " AND seg_medicos.cod_medico=" & Val(SkinLabel46.Caption) _
        & " AND seg_pacientes.cod_paciente=" & Val(SkinLabel41.Caption) _
        & " AND seg_tratamentos.cod_tratamento=" & Val(SkinLabel38.Caption)
        
    Set rst = db.OpenRecordset(SQL)
    
    If rst("Data") > DateValue(Label18.Caption) Then
    
        InserirAlterarConsultasTratamento.Text1.Text = rst("cod_consultatratamento")
        InserirAlterarConsultasTratamento.Text2.Text = rst("data")
        InserirAlterarConsultasTratamento.Text3.Text = rst("hora")
        InserirAlterarConsultasTratamento.cmbTratamento.Text = rst("nome_tratamento")
        InserirAlterarConsultasTratamento.Text4.Text = rst("seg_tratamentos.cod_tratamento")
        InserirAlterarConsultasTratamento.Text5.Text = rst("preco_tratamento")
        InserirAlterarConsultasTratamento.cmbPacientes.Text = rst("nome_paciente")
        InserirAlterarConsultasTratamento.Text6.Text = rst("seg_pacientes.cod_paciente")
        
        If rst("seg_pacientes.telefone") = Null Then
            InserirAlterarConsultasTratamento.Text7.Text = ""
        Else
            x = rst("seg_pacientes.telefone")
            InserirAlterarConsultasTratamento.Text7.Text = "" & x
        End If
    
        If rst("seg_pacientes.telemovel") = Null Then
            InserirAlterarConsultasTratamento.Text8.Text = ""
        Else
            x = rst("seg_pacientes.telemovel")
            InserirAlterarConsultasTratamento.Text8.Text = "" & x
        End If
    
        If rst("seg_pacientes.email") = Null Then
            InserirAlterarConsultasTratamento.Text9.Text = ""
        Else
            x = rst("seg_pacientes.email")
            InserirAlterarConsultasTratamento.Text9.Text = "" & x
        End If
        
        InserirAlterarConsultasTratamento.cmbMedicos.Text = rst("nome_medico")
        InserirAlterarConsultasTratamento.Text10.Text = rst("seg_medicos.cod_medico")
        InserirAlterarConsultasTratamento.Show
        InserirAlterarConsultasTratamento.Caption = "Alterar Dados da Consulta de Tratamento"
        InserirAlterarConsultasTratamento.cmdInserir.Caption = "Alterar"
      
        SQL = " SELECT * FROM Tratamentos ORDER BY nome_tratamento"
        Set rst = db.OpenRecordset(SQL)
        With rst
            If Not rst.EOF Then
                Do While Not rst.EOF
                    InserirAlterarConsultasTratamento.cmbTratamento.AddItem rst("nome_Tratamento")
                    InserirAlterarConsultasTratamento.cmbTratamento.ItemData(InserirAlterarConsultasTratamento.cmbTratamento.NewIndex) = rst("cod_tratamento")
                    rst.MoveNext
                Loop
            Else
                InserirAlterarConsultasTratamento.cmbTratamento.Text = ""
            End If
        End With
       
        SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
        Set rst = db.OpenRecordset(SQL)
        With rst
            If Not rst.EOF Then
                Do While Not rst.EOF
                    InserirAlterarConsultasTratamento.cmbPacientes.AddItem rst("nome_paciente")
                    InserirAlterarConsultasTratamento.cmbPacientes.ItemData(InserirAlterarConsultasTratamento.cmbPacientes.NewIndex) = rst("cod_Paciente")
                    rst.MoveNext
                Loop
            Else
                InserirAlterarConsultasTratamento.cmbPacientes.Text = ""
            End If
        End With
        
        SQL = " SELECT * FROM Medicos ORDER BY nome_medico"
        Set rst = db.OpenRecordset(SQL)
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    InserirAlterarConsultasTratamento.cmbMedicos.AddItem rst("nome_medico")
                    InserirAlterarConsultasTratamento.cmbMedicos.ItemData(InserirAlterarConsultasTratamento.cmbMedicos.NewIndex) = rst("cod_medico")
                    .MoveNext
                Loop
            Else
                InserirAlterarConsultasTratamento.cmbMedicos.Text = ""
            End If
        End With
        
    Else
        If rst("Data") = DateValue(Label18.Caption) Then
            If rst("hora") > TimeValue(Label19.Caption) Then
            
                InserirAlterarConsultasTratamento.Text1.Text = rst("cod_consultatratamento")
                InserirAlterarConsultasTratamento.Text2.Text = rst("data")
                InserirAlterarConsultasTratamento.Text3.Text = rst("hora")
                InserirAlterarConsultasTratamento.cmbTratamento.Text = rst("nome_tratamento")
                InserirAlterarConsultasTratamento.Text4.Text = rst("seg_tratamentos.cod_tratamento")
                InserirAlterarConsultasTratamento.Text5.Text = rst("preco_tratamento")
                InserirAlterarConsultasTratamento.cmbPacientes.Text = rst("nome_paciente")
                InserirAlterarConsultasTratamento.Text6.Text = rst("seg_pacientes.cod_paciente")
                
                If rst("seg_pacientes.telefone") = Null Then
                    InserirAlterarConsultasTratamento.Text7.Text = ""
                Else
                    x = rst("seg_pacientes.telefone")
                    InserirAlterarConsultasTratamento.Text7.Text = "" & x
                End If
    
                If rst("seg_pacientes.telemovel") = Null Then
                    InserirAlterarConsultasTratamento.Text8.Text = ""
                Else
                    x = rst("seg_pacientes.telemovel")
                    InserirAlterarConsultasTratamento.Text8.Text = "" & x
                End If
    
                If rst("seg_pacientes.email") = Null Then
                    InserirAlterarConsultasTratamento.Text9.Text = ""
                Else
                    x = rst("seg_pacientes.email")
                    InserirAlterarConsultasTratamento.Text9.Text = "" & x
                End If
                
                InserirAlterarConsultasTratamento.cmbMedicos.Text = rst("nome_medico")
                InserirAlterarConsultasTratamento.Text10.Text = rst("seg_medicos.cod_medico")
                InserirAlterarConsultasTratamento.Show
                InserirAlterarConsultasTratamento.Caption = "Alterar Dados da Consulta de Tratamento"
                InserirAlterarConsultasTratamento.cmdInserir.Caption = "Alterar"
                
                SQL = " SELECT * FROM Tratamentos ORDER BY nome_tratamento"
                Set rst = db.OpenRecordset(SQL)
                With rst
                    If Not rst.EOF Then
                        Do While Not rst.EOF
                            InserirAlterarConsultasTratamento.cmbTratamento.AddItem rst("nome_Tratamento")
                            InserirAlterarConsultasTratamento.cmbTratamento.ItemData(InserirAlterarConsultasTratamento.cmbTratamento.NewIndex) = rst("cod_tratamento")
                            rst.MoveNext
                        Loop
                    Else
                        InserirAlterarConsultasTratamento.cmbTratamento.Text = ""
                    End If
                End With
                
                SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
                Set rst = db.OpenRecordset(SQL)
                With rst
                    If Not rst.EOF Then
                        Do While Not rst.EOF
                            InserirAlterarConsultasTratamento.cmbPacientes.AddItem rst("nome_paciente")
                            InserirAlterarConsultasTratamento.cmbPacientes.ItemData(InserirAlterarConsultasTratamento.cmbPacientes.NewIndex) = rst("cod_Paciente")
                            rst.MoveNext
                        Loop
                    Else
                        InserirAlterarConsultasTratamento.cmbPacientes.Text = ""
                    End If
                End With
               
                SQL = " SELECT * FROM Medicos ORDER BY nome_medico"
                Set rst = db.OpenRecordset(SQL)
                With rst
                    If Not .EOF Then
                        Do While Not .EOF
                            InserirAlterarConsultasTratamento.cmbMedicos.AddItem rst("nome_medico")
                            InserirAlterarConsultasTratamento.cmbMedicos.ItemData(InserirAlterarConsultasTratamento.cmbMedicos.NewIndex) = rst("cod_medico")
                            .MoveNext
                        Loop
                    Else
                        InserirAlterarConsultasTratamento.cmbMedicos.Text = ""
                    End If
                End With
                Exit Sub
               
            Else
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser alterada porque já foi realizada!"
                FormMsgBoxNormal.Caption = "Alterar (Consultas de Tratamento)"
                PainelControlo.Enabled = False
                Exit Sub
            End If
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser alterada porque já foi realizada!"
            FormMsgBoxNormal.Caption = "Alterar (Consultas de Tratamento)"
            PainelControlo.Enabled = False
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser alterada porque já foi realizada!"
            FormMsgBoxNormal.Caption = "Alterar (Consultas de Tratamento)"
            PainelControlo.Enabled = False
        End If
    End If
   
End Sub

Private Sub cmdDel2_Click()

    FormMsgBoxSimNao.Show
    FormMsgBoxSimNao.Label1.Caption = "Tem a certeza que quer eliminar esta consulta de tratamento?"
    FormMsgBoxSimNao.Caption = "Eliminar (Consultas de Tratamento)"
    PainelControlo.Enabled = False
   
End Sub

Private Sub cmdInserir2_Click()

    PainelControlo.Enabled = False
    InserirAlterarConsultasTratamento.Show
    InserirAlterarConsultasTratamento.Caption = "Marcar Nova Consulta de Tratamento"
    InserirAlterarConsultasTratamento.cmdInserir.Caption = "&Inserir"
    InserirAlterarConsultasTratamento.cmbTratamento.Enabled = True
    InserirAlterarConsultasTratamento.cmbPacientes.Enabled = True
    InserirAlterarConsultasTratamento.cmbMedicos.Enabled = True
   
    SQL = " SELECT * FROM Consultas_Tratamentos"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        InserirAlterarConsultasTratamento.Text1.Text = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_consultatratamento")
                    .MoveNext
                Loop
            Else
                y = rst("cod_consultatratamento")
            End If
        End With
        res = x
        InserirAlterarConsultasTratamento.Text1.Text = Val(res) + 1
    End If
    
    SQL = " SELECT * FROM Tratamentos ORDER BY nome_tratamento"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not rst.EOF Then
            Do While Not rst.EOF
                InserirAlterarConsultasTratamento.cmbTratamento.AddItem rst("nome_tratamento")
                rst.MoveNext
            Loop
        Else
            InserirAlterarConsultasTratamento.cmbTratamento.Text = ""
        End If
    End With
  
    SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not rst.EOF Then
            Do While Not rst.EOF
                InserirAlterarConsultasTratamento.cmbPacientes.AddItem rst("nome_paciente")
                rst.MoveNext
            Loop
        Else
            InserirAlterarConsultasTratamento.cmbPacientes.Text = ""
        End If
    End With
   
    SQL = " SELECT * FROM Medicos ORDER BY nome_medico"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not .EOF Then
            Do While Not .EOF
                InserirAlterarConsultasTratamento.cmbMedicos.AddItem rst("nome_medico")
                .MoveNext
            Loop
        Else
            InserirAlterarConsultasTratamento.cmbMedicos.Text = ""
        End If
    End With

End Sub

Private Sub cmdProcurar2_Click()

    InputBoxProcurar.Caption = "Procurar (Consultas de Tratamento)"
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor2_Click()

    Dim itemx As ListItem
    
    SQL = " SELECT consultas_tratamentos.cod_consultatratamento, consultas_tratamentos.data, consultas_tratamentos.hora, consultas_tratamentos.cod_tratamento, seg_pacientes.nome_paciente " _
        & " FROM consultas_tratamentos, seg_pacientes, seg_tratamentos" _
        & " WHERE consultas_tratamentos.cod_tratamento = seg_tratamentos.cod_tratamento" _
        & " AND consultas_tratamentos.cod_paciente=seg_pacientes.cod_paciente" _
        & " ORDER BY data desc"
        
    Set rst = db.OpenRecordset(SQL)
    
    ListConsultasTratamento.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListConsultasTratamento.ListItems.Add(, , rst("cod_consultatratamento"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("cod_tratamento")
        itemx.SubItems(4) = rst("nome_paciente")
        itemx.Tag = rst("cod_consultatratamento")
        
        rst.MoveNext
    Wend
    
    cmdProcurar2.Visible = True
    cmdRepor2.Visible = False
    
End Sub

Private Sub cmdSair2_Click()
    
    End
    
End Sub

Private Sub ListConsultasTratamento_SetUp()

    Dim itemx As ListItem
    
    SQL = " SELECT consultas_tratamentos.cod_consultatratamento, consultas_tratamentos.data, consultas_tratamentos.hora, consultas_tratamentos.cod_tratamento, seg_pacientes.nome_paciente " _
        & " FROM consultas_tratamentos, seg_pacientes" _
        & " WHERE consultas_tratamentos.cod_tratamento = cod_tratamento" _
        & " AND consultas_tratamentos.cod_paciente=seg_pacientes.cod_paciente" _
        & " ORDER BY consultas_tratamentos.data desc "
        
    Set rst = db.OpenRecordset(SQL)
     
    ListConsultasTratamento.ListItems.Clear
    ListConsultasTratamento.ColumnHeaders.Add , , "Código da Consulta", ListConsultasTratamento.Width - 8150
    ListConsultasTratamento.ColumnHeaders.Add , , "Data", ListConsultasTratamento.Width - 8400
    ListConsultasTratamento.ColumnHeaders.Add , , "Hora", ListConsultasTratamento.Width - 8550
    ListConsultasTratamento.ColumnHeaders.Add , , "Código do Tratamento", ListConsultasTratamento.Width - 7950
    ListConsultasTratamento.ColumnHeaders.Add , , "Nome do Paciente", ListConsultasTratamento.Width - 5970
    ListConsultasTratamento.View = lvwReport
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListConsultasTratamento.ListItems.Add(, , rst("cod_consultatratamento"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("cod_tratamento")
        itemx.SubItems(4) = rst("nome_paciente")
        itemx.Tag = rst("cod_consultatratamento")
        
        rst.MoveNext
    Wend
   
End Sub

Private Sub ListConsultasTratamentos_Click()
    
    Preencher_dados_tratamentos
    
End Sub

Private Sub Preencher_dados_tratamentos()
    
    SQL = "select * from consultas_tratamentos"
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
    SQL = " SELECT consultas_tratamentos.cod_consultatratamento, consultas_tratamentos.data, consultas_tratamentos.hora, consultas_tratamentos.cod_tratamento, seg_tratamentos.nome_tratamento, seg_tratamentos.preco_tratamento, seg_pacientes.cod_paciente, seg_pacientes.nome_paciente, seg_pacientes.telefone, seg_pacientes.telemovel, seg_pacientes.email, seg_medicos.cod_medico, seg_medicos.nome_medico" _
        & " FROM consultas_tratamentos, seg_pacientes, seg_medicos, seg_tratamentos" _
        & " WHERE consultas_tratamentos.cod_paciente = seg_pacientes.cod_paciente" _
        & " AND consultas_tratamentos.cod_medico = seg_medicos.cod_medico" _
        & " AND consultas_tratamentos.cod_tratamento = seg_tratamentos.cod_tratamento" _
        & " AND cod_consultatratamento=" & ListConsultasTratamento.SelectedItem.Tag _
        & " ORDER BY data, hora"
        
    Set rst = db.OpenRecordset(SQL)
    
    SkinLabel34.Caption = rst("cod_consultatratamento")
    SkinLabel35.Caption = rst("Data")
    SkinLabel36.Caption = rst("Hora")
    SkinLabel37.Caption = rst("nome_tratamento")
    SkinLabel38.Caption = rst("cod_tratamento")
    SkinLabel39.Caption = rst("preco_tratamento")
    SkinLabel40.Caption = rst("nome_paciente")
    SkinLabel41.Caption = rst("cod_paciente")
    
    If rst("telefone") = Null Then
        SkinLabel42.Caption = ""
    Else
        x = rst("telefone")
        SkinLabel42.Caption = "" & x
    End If
    
    If rst("telemovel") = Null Then
        SkinLabel43.Caption = ""
    Else
        x = rst("telemovel")
        SkinLabel43.Caption = "" & x
    End If
    
    If rst("email") = Null Then
        SkinLabel44.Caption = ""
    Else
        x = rst("email")
        SkinLabel44.Caption = "" & x
    End If
    
    SkinLabel45.Caption = rst("nome_medico")
    SkinLabel46.Caption = rst("cod_medico")
    End If
   
End Sub

Private Sub ListConsultasTratamento_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Preencher_dados_tratamentos
    
End Sub
'----------------------------------------Fim código para o separador "Consultas de Tratamento"----------------------------

'-------------------------------------------Código para o separador "Pacientes"-----------------------------
Private Sub cmdAlterar3_Click()
    SQL = "SELECT * FROM Pacientes" _
        & " WHERE cod_paciente=" & Val(PainelControlo.ListPacientes.SelectedItem.Tag)
        
    Set rst = db.OpenRecordset(SQL)
    
    InserirAlterarPacientes.Text1.Text = rst("nome_paciente")
    InserirAlterarPacientes.Text2.Text = rst("cod_paciente")
    InserirAlterarPacientes.Text3.Text = rst("data_nascimento")
    
    If rst("telefone") = Null Then
        InserirAlterarPacientes.Text4.Text = ""
    Else
        x = rst("telefone")
        InserirAlterarPacientes.Text4.Text = "" & x
    End If
    
    If rst("telemovel") = Null Then
        InserirAlterarPacientes.Text5.Text = ""
    Else
        x = rst("telemovel")
        InserirAlterarPacientes.Text5.Text = "" & x
    End If
    
    If rst("email") = Null Then
        InserirAlterarPacientes.Text6.Text = ""
    Else
        x = rst("email")
        InserirAlterarPacientes.Text6.Text = "" & x
    End If
    
    If rst("bilhete_identidade") = "" Then
        InserirAlterarPacientes.Text7.Enabled = True
    End If
    
    If rst("Bilhete_identidade") = Null Then
        InserirAlterarPacientes.Text7.Text = ""
    Else
        x = rst("bilhete_identidade")
        InserirAlterarPacientes.Text7.Text = "" & x
    End If
    
    If rst("num_contribuinte") = "" Then
        InserirAlterarPacientes.Text8.Enabled = True
    End If
    
    If rst("num_contribuinte") = Null Then
        InserirAlterarPacientes.Text8.Text = ""
    Else
        x = rst("num_contribuinte")
        InserirAlterarPacientes.Text8.Text = "" & x
    End If
    
    InserirAlterarPacientes.Text9.Text = rst("morada")
    InserirAlterarPacientes.Text10.Text = rst("cod_postal")
    InserirAlterarPacientes.Text11.Text = rst("cidade")
    InserirAlterarPacientes.Text12.Text = rst("sexo")
    InserirAlterarPacientes.Text13.Text = rst("estado_civil")
    
    If rst("doencas") = Null Then
        InserirAlterarPacientes.Text14.Text = ""
    Else
        x = rst("doencas")
        InserirAlterarPacientes.Text14.Text = "" & x
    End If
    
    InserirAlterarPacientes.Show
    InserirAlterarPacientes.Caption = "Alterar Dados do Paciente"
    InserirAlterarPacientes.cmdInserir.Caption = "Alterar"
    
End Sub

Private Sub cmdDel3_Click()

    FormMsgBoxSimNao.Show
    FormMsgBoxSimNao.Label1.Caption = "Tem a certeza que quer eliminar este paciente?"
    FormMsgBoxSimNao.Caption = "Eliminar (Pacientes)"
    PainelControlo.Enabled = False
    
End Sub

Private Sub cmdInserir3_Click()
    
    PainelControlo.Enabled = False
    InserirAlterarPacientes.Show
    InserirAlterarPacientes.Caption = "Inserir novo Paciente"
    InserirAlterarPacientes.cmdInserir.Caption = "&Inserir"
    InserirAlterarPacientes.Text1.Enabled = True
    InserirAlterarPacientes.Text1.SetFocus
    InserirAlterarPacientes.Text3.Enabled = True
    InserirAlterarPacientes.Text7.Enabled = True
    InserirAlterarPacientes.Text8.Enabled = True
    InserirAlterarPacientes.Text12.Enabled = True
    InserirAlterarPacientes.Text13.Enabled = True
    SQL = " SELECT * FROM seg_Pacientes"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        InserirAlterarPacientes.Text2.Text = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_Paciente")
                    .MoveNext
                Loop
            Else
                y = rst("cod_Paciente")
            End If
        End With
        res = x
        InserirAlterarPacientes.Text2.Text = Val(res) + 1
    End If

End Sub

Private Sub cmdProcurar3_Click()

    InputBoxProcurar.Caption = "Procurar (Pacientes)"
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor3_Click()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM pacientes" _
        & " ORDER BY nome_paciente"
        
    Set rst = db.OpenRecordset(SQL)
    
    ListPacientes.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListPacientes.ListItems.Add(, , rst("cod_Paciente"))
        itemx.SubItems(1) = rst("nome_paciente")
        itemx.SubItems(2) = rst("data_nascimento")
        itemx.Tag = rst("cod_Paciente")
        
        rst.MoveNext
    Wend
    
    cmdProcurar3.Visible = True
    cmdRepor3.Visible = False
    
End Sub

Private Sub cmdSair3_Click()
    
    End
    
End Sub

Private Sub ListPacientes_SetUp()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM pacientes" _
        & " ORDER BY nome_paciente"
        
    Set rst = db.OpenRecordset(SQL)
     
    ListPacientes.ListItems.Clear
    ListPacientes.ColumnHeaders.Add , , "Código do Paciente", ListPacientes.Width - 7500
    ListPacientes.ColumnHeaders.Add , , "Nome do Paciente", ListPacientes.Width - 4550
    ListPacientes.ColumnHeaders.Add , , "Data de Nascimento", ListPacientes.Width - 7500
    ListPacientes.View = lvwReport
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListPacientes.ListItems.Add(, , rst("cod_Paciente"))
        itemx.SubItems(1) = rst("nome_paciente")
        itemx.SubItems(2) = rst("data_nascimento")
        itemx.Tag = rst("cod_Paciente")
        
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListPaceintes_Click()
   
    Preencher_dados_Paceintes
    
End Sub

Private Sub Preencher_dados_Paceintes()

    SQL = " SELECT * FROM pacientes"
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
    SQL = " SELECT * " _
        & " FROM pacientes" _
        & " WHERE cod_paciente = " & ListPacientes.SelectedItem.Tag _
        & " ORDER BY nome_paciente"
        
    Set rst = db.OpenRecordset(SQL)
    
    SkinLabel61.Caption = rst("nome_paciente")
    SkinLabel62.Caption = rst("cod_paciente")
    SkinLabel63.Caption = rst("data_nascimento")
    
    If rst("telefone") = Null Then
        SkinLabel64.Caption = ""
    Else
        x = rst("telefone")
        SkinLabel64.Caption = "" & x
    End If
    
    If rst("telemovel") = Null Then
        SkinLabel65.Caption = ""
    Else
        x = rst("telemovel")
        SkinLabel65.Caption = "" & x
    End If
    
    If rst("email") = Null Then
        SkinLabel66.Caption = ""
    Else
        x = rst("email")
        SkinLabel66.Caption = "" & x
    End If
    
    If rst("bilhete_identidade") = Null Then
        SkinLabel67.Caption = ""
    Else
        x = rst("bilhete_identidade")
        SkinLabel67.Caption = "" & x
    End If
    
    If rst("Num_contribuinte") = Null Then
        SkinLabel68.Caption = ""
    Else
        x = rst("num_contribuinte")
        SkinLabel68.Caption = "" & x
    End If
    
    SkinLabel69.Caption = rst("morada")
    SkinLabel70.Caption = rst("cod_postal")
    SkinLabel71.Caption = rst("cidade")
    SkinLabel72.Caption = rst("sexo")
    SkinLabel73.Caption = rst("estado_civil")
    
    If rst("doencas") = Null Then
        SkinLabel74.Caption = ""
    Else
        x = rst("doencas")
        SkinLabel74.Caption = "" & x
    End If
    End If
    
End Sub

Private Sub ListPacientes_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Preencher_dados_Paceintes
    
End Sub
'-------------------------------------------Fim do código para o separador "Pacientes"-----------------------------

'-------------------------------------------Código para o separador "Médicos"-----------------------------
Private Sub cmdAlterar4_Click()

    SQL = "SELECT * FROM Medicos" _
        & " WHERE cod_medico=" & Val(PainelControlo.ListMedicos.SelectedItem.Tag)
        
    Set rst = db.OpenRecordset(SQL)
    
    InserirAlterarMedicos.Text1.Text = rst("nome_medico")
    InserirAlterarMedicos.Text2.Text = rst("cod_medico")
    InserirAlterarMedicos.Text3.Text = rst("data_nascimento")
    
    If rst("telefone") = Null Then
        InserirAlterarMedicos.Text4.Text = ""
    Else
        x = rst("telefone")
        InserirAlterarMedicos.Text4.Text = "" & x
    End If
    
    If rst("telemovel") = Null Then
        InserirAlterarMedicos.Text5.Text = ""
    Else
        x = rst("telemovel")
        InserirAlterarMedicos.Text5.Text = "" & x
    End If
    
    If rst("email") = Null Then
        InserirAlterarMedicos.Text6.Text = ""
    Else
        x = rst("email")
        InserirAlterarMedicos.Text6.Text = "" & x
    End If
    
    If rst("bilhete_identidade") = "" Then
        InserirAlterarMedicos.Text7.Enabled = True
    End If
    
    If rst("bilhete_identidade") = Null Then
        InserirAlterarMedicos.Text7.Text = ""
    Else
        x = rst("bilhete_identidade")
        InserirAlterarMedicos.Text7.Text = "" & x
    End If
    
    If rst("num_contribuinte") = "" Then
        InserirAlterarMedicos.Text8.Enabled = True
    End If
    
    If rst("num_contribuinte") = Null Then
        InserirAlterarMedicos.Text8.Text = ""
    Else
        x = rst("num_contribuinte")
        InserirAlterarMedicos.Text8.Text = "" & x
    End If
    
    InserirAlterarMedicos.Text9.Text = rst("morada")
    InserirAlterarMedicos.Text10.Text = rst("cod_postal")
    InserirAlterarMedicos.Text11.Text = rst("cidade")
    InserirAlterarMedicos.Text12.Text = rst("sexo")
    InserirAlterarMedicos.Text13.Text = rst("estado_civil")
    InserirAlterarMedicos.Show
    InserirAlterarMedicos.Caption = "Alterar Dados do Médico"
    InserirAlterarMedicos.cmdInserir.Caption = "Alterar"
    
End Sub

Private Sub cmdDel4_Click()

    FormMsgBoxSimNao.Show
    FormMsgBoxSimNao.Label1.Caption = "Tem a certeza que quer eliminar este médico?"
    FormMsgBoxSimNao.Caption = "Eliminar (Médicos)"
    PainelControlo.Enabled = False

End Sub

Private Sub cmdInserir4_Click()
    
    PainelControlo.Enabled = False
    InserirAlterarMedicos.Show
    InserirAlterarMedicos.Caption = "Inserir novo Médico"
    InserirAlterarMedicos.cmdInserir.Caption = "&Inserir"
    InserirAlterarMedicos.Text1.Enabled = True
    InserirAlterarMedicos.Text1.SetFocus
    InserirAlterarMedicos.Text3.Enabled = True
    InserirAlterarMedicos.Text7.Enabled = True
    InserirAlterarMedicos.Text8.Enabled = True
    InserirAlterarMedicos.Text12.Enabled = True
    InserirAlterarMedicos.Text13.Enabled = True
    SQL = " SELECT * FROM seg_Medicos"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        InserirAlterarMedicos.Text2.Text = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_medico")
                    .MoveNext
                Loop
            Else
                y = rst("cod_medico")
            End If
        End With
        res = x
        InserirAlterarMedicos.Text2.Text = Val(res) + 1
    End If

End Sub

Private Sub cmdProcurar4_Click()

    InputBoxProcurar.Caption = "Procurar (Médicos)"
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor4_Click()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM medicos" _
        & " ORDER BY nome_medico"
        
    Set rst = db.OpenRecordset(SQL)
    
    ListMedicos.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListMedicos.ListItems.Add(, , rst("cod_medico"))
        itemx.SubItems(1) = rst("nome_medico")
        itemx.SubItems(2) = rst("data_nascimento")
        itemx.Tag = rst("cod_medico")
        
        rst.MoveNext
    Wend
    
    cmdProcurar4.Visible = True
    cmdRepor4.Visible = False
    
End Sub

Private Sub cmdSair4_Click()
    
    End
    
End Sub

Private Sub ListMedicos_SetUp()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM medicos" _
        & " ORDER BY nome_medico"
        
    Set rst = db.OpenRecordset(SQL)
     
    ListMedicos.ListItems.Clear
    ListMedicos.ColumnHeaders.Add , , "Código do Médico", ListMedicos.Width - 7500
    ListMedicos.ColumnHeaders.Add , , "Nome do Médico", ListMedicos.Width - 4550
    ListMedicos.ColumnHeaders.Add , , "Data de Nascimento", ListMedicos.Width - 7500
    ListMedicos.View = lvwReport
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListMedicos.ListItems.Add(, , rst("cod_medico"))
        itemx.SubItems(1) = rst("nome_medico")
        itemx.SubItems(2) = rst("data_nascimento")
        itemx.Tag = rst("cod_medico")
        
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListMedicos_Click()
   
    Preencher_dados_Medicos
    
End Sub

Private Sub Preencher_dados_Medicos()
    
    SQL = "select * from medicos"
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
    SQL = " SELECT * " _
        & " FROM medicos" _
        & " WHERE cod_medico = " & ListMedicos.SelectedItem.Tag _
        & " ORDER BY nome_medico"
        
    Set rst = db.OpenRecordset(SQL)
    
    SkinLabel89.Caption = rst("nome_medico")
    SkinLabel90.Caption = rst("cod_medico")
    SkinLabel91.Caption = rst("data_nascimento")
    
    If rst("telefone") = Null Then
        SkinLabel92.Caption = ""
    Else
        x = rst("telefone")
        SkinLabel92.Caption = "" & x
    End If
    
    If rst("telemovel") = Null Then
        SkinLabel93.Caption = ""
    Else
        x = rst("telemovel")
        SkinLabel93.Caption = "" & x
    End If
    
    If rst("email") = Null Then
        SkinLabel94.Caption = ""
    Else
        x = rst("email")
        SkinLabel94.Caption = "" & x
    End If
    
    If rst("bilhete_identidade") = Null Then
        SkinLabel95.Caption = ""
    Else
        x = rst("bilhete_identidade")
        SkinLabel95.Caption = "" & x
    End If
    
    If rst("num_contribuinte") = Null Then
        SkinLabel96.Caption = ""
    Else
        x = rst("num_contribuinte")
        SkinLabel96.Caption = "" & x
    End If
    
    SkinLabel97.Caption = rst("morada")
    SkinLabel98.Caption = rst("cod_postal")
    SkinLabel99.Caption = rst("cidade")
    SkinLabel100.Caption = rst("sexo")
    SkinLabel101.Caption = rst("estado_civil")
    End If
    
End Sub

Private Sub ListMedicos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Preencher_dados_Medicos
    
End Sub
'-------------------------------------------Fim do código para o separador "Médicos"-----------------------------

'-------------------------------------------Código para o separador "Tratamentos"-----------------------------
Private Sub cmdAlterar5_Click()

    SQL = "SELECT * FROM Tratamentos" _
        & " WHERE cod_tratamento=" & Val(PainelControlo.ListTratamentos.SelectedItem.Tag)
        
    Set rst = db.OpenRecordset(SQL)
    
    InserirAlterarTratamentos.Text1.Text = rst("nome_tratamento")
    InserirAlterarTratamentos.Text2.Text = rst("cod_tratamento")
    InserirAlterarTratamentos.Text3.Text = rst("preco_tratamento")
    InserirAlterarTratamentos.Text4.Text = rst("descricao")
    InserirAlterarTratamentos.Show
    InserirAlterarTratamentos.Caption = "Alterar Dados do Tratamento"
    InserirAlterarTratamentos.cmdInserir.Caption = "Alterar"
    
End Sub

Private Sub cmdDel5_Click()

    FormMsgBoxSimNao.Show
    FormMsgBoxSimNao.Label1.Caption = "Tem a certeza que quer eliminar este tratamento?"
    FormMsgBoxSimNao.Caption = "Eliminar (Tratamentos)"
    PainelControlo.Enabled = False

End Sub

Private Sub cmdInserir5_Click()

    PainelControlo.Enabled = False
    InserirAlterarTratamentos.Show
    InserirAlterarTratamentos.Caption = "Inserir novo Tratamento"
    InserirAlterarTratamentos.cmdInserir.Caption = "&Inserir"
    InserirAlterarTratamentos.Text1.Enabled = True
    InserirAlterarTratamentos.Text1.SetFocus
    InserirAlterarTratamentos.Text3.Enabled = True
    InserirAlterarTratamentos.Text4.Enabled = True
    SQL = " SELECT * FROM seg_Tratamentos"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        InserirAlterarTratamentos.Text2.Text = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_tratamento")
                    .MoveNext
                Loop
            Else
                y = rst("cod_tratamento")
            End If
        End With
        res = x
        InserirAlterarTratamentos.Text2.Text = Val(res) + 1
    End If

End Sub

Private Sub cmdProcurar5_Click()

    InputBoxProcurar.Caption = "Procurar (Tratamentos)"
    InputBoxProcurar.Label1.Caption = "Indique o nome do tratamento que pretende procurar."
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor5_Click()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM tratamentos" _
        & " ORDER BY nome_tratamento"
        
    Set rst = db.OpenRecordset(SQL)
    
    ListTratamentos.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListTratamentos.ListItems.Add(, , rst("cod_tratamento"))
        itemx.SubItems(1) = rst("nome_tratamento")
        itemx.SubItems(2) = rst("preco_tratamento")
        itemx.Tag = rst("cod_tratamento")
        
        rst.MoveNext
    Wend
    
    cmdProcurar5.Visible = True
    cmdRepor5.Visible = False
    
End Sub

Private Sub cmdSair5_Click()

    End
    
End Sub

Private Sub ListTratamentos_SetUp()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM tratamentos" _
        & " ORDER BY nome_tratamento"
        
    Set rst = db.OpenRecordset(SQL)
     
    ListTratamentos.ListItems.Clear
    ListTratamentos.ColumnHeaders.Add , , "Código do Tratamento", ListTratamentos.Width - 7500
    ListTratamentos.ColumnHeaders.Add , , "Nome do Tratamento", ListTratamentos.Width - 4550
    ListTratamentos.ColumnHeaders.Add , , "Preço do Tratamento", ListTratamentos.Width - 7500
    ListTratamentos.View = lvwReport
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListTratamentos.ListItems.Add(, , rst("cod_Tratamento"))
        itemx.SubItems(1) = rst("nome_tratamento")
        itemx.SubItems(2) = rst("preco_tratamento")
        itemx.Tag = rst("cod_tratamento")
        
        rst.MoveNext
    Wend
    
End Sub
Private Sub ListTratamentos_Click()

    Preencher_dados_tratamento
    
End Sub

Private Sub Preencher_dados_tratamento()

    SQL = " SELECT * FROM tratamentos"
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
    SQL = " SELECT * " _
        & " FROM tratamentos" _
        & " WHERE cod_tratamento = " & ListTratamentos.SelectedItem.Tag _
        & " ORDER BY nome_tratamento"
        
    Set rst = db.OpenRecordset(SQL)
    
    SkinLabel106.Caption = rst("nome_tratamento")
    SkinLabel107.Caption = rst("cod_tratamento")
    SkinLabel108.Caption = rst("preco_tratamento")
    SkinLabel109.Caption = rst("descricao")
    End If
    
End Sub

Private Sub ListTratamentos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Preencher_dados_tratamento
    
End Sub
'-------------------------------------------Fim de código para o separador "Tratamentos"-----------------------------

'-------------------------------------------Código para o separador "Medicamentos"-----------------------------
Private Sub cmdAlterar6_Click()

    SQL = "SELECT * FROM Medicamentos" _
        & " WHERE cod_Medicamento=" & Val(PainelControlo.ListMedicamentos.SelectedItem.Tag)
        
    Set rst = db.OpenRecordset(SQL)
    
    InserirAlterarMedicamentos.Text1.Text = rst("nome")
    InserirAlterarMedicamentos.Text2.Text = rst("cod_medicamento")
    InserirAlterarMedicamentos.Text3.Text = rst("preco_medicamento")
    If rst("comprimidos_caixa") = Null Then
        InserirAlterarMedicamentos.Text4.Text = ""
    Else
        x = rst("comprimidos_caixa")
        InserirAlterarMedicamentos.Text4.Text = "" & x
    End If
    InserirAlterarMedicamentos.Text5.Text = rst("tipo_medicamento")
    InserirAlterarMedicamentos.Text6.Text = rst("emblagem_disponiveis")
    InserirAlterarMedicamentos.Text7.Text = rst("descricao")
    InserirAlterarMedicamentos.Show
    InserirAlterarMedicamentos.Caption = "Alterar Dados do Medicamento"
    InserirAlterarMedicamentos.cmdInserir.Caption = "Alterar"
    
End Sub

Private Sub cmdDel6_Click()

    FormMsgBoxSimNao.Show
    FormMsgBoxSimNao.Label1.Caption = "Tem a certeza que quer eliminar este medicamento?"
    FormMsgBoxSimNao.Caption = "Eliminar (Medicamentos)"
    PainelControlo.Enabled = False

End Sub

Private Sub cmdInserir6_Click()
    
    PainelControlo.Enabled = False
    InserirAlterarMedicamentos.Show
    InserirAlterarMedicamentos.Caption = "Inserir novo Medicamento"
    InserirAlterarMedicamentos.cmdInserir.Caption = "&Inserir"
    InserirAlterarMedicamentos.Text1.Enabled = True
    InserirAlterarMedicamentos.Text1.SetFocus
    InserirAlterarMedicamentos.Text4.Enabled = True
    InserirAlterarMedicamentos.Text5.Enabled = True
    SQL = " SELECT * FROM Medicamentos"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        InserirAlterarMedicamentos.Text2.Text = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_Medicamento")
                    .MoveNext
                Loop
            Else
                y = rst("cod_Medicamento")
            End If
        End With
        res = x
        InserirAlterarMedicamentos.Text2.Text = res + 1
    End If

End Sub

Private Sub cmdProcurar6_Click()

    InputBoxProcurar.Caption = "Procurar (Medicamentos)"
    InputBoxProcurar.Label1.Caption = "Indique o nome do medicamento que pretende procurar."
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor6_Click()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM Medicamentos" _
        & " ORDER BY nome"
        
    Set rst = db.OpenRecordset(SQL)
    
    ListMedicamentos.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListMedicamentos.ListItems.Add(, , rst("cod_Medicamento"))
        itemx.SubItems(1) = rst("nome")
        itemx.SubItems(2) = rst("preco_Medicamento")
        itemx.Tag = rst("cod_Medicamento")
        
        rst.MoveNext
    Wend
    
    cmdProcurar6.Visible = True
    cmdRepor6.Visible = False
    
End Sub

Private Sub cmdSair6_Click()

    End
    
End Sub

Private Sub ListMedicamentos_SetUp()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM Medicamentos" _
        & " ORDER BY nome"
        
    Set rst = db.OpenRecordset(SQL)
     
    ListMedicamentos.ListItems.Clear
    ListMedicamentos.ColumnHeaders.Add , , "Código do Medicamento", ListMedicamentos.Width - 7500
    ListMedicamentos.ColumnHeaders.Add , , "Nome do Medicamento", ListMedicamentos.Width - 4550
    ListMedicamentos.ColumnHeaders.Add , , "Preço do Medicamento", ListMedicamentos.Width - 7500
    ListMedicamentos.View = lvwReport
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
    Set itemx = ListMedicamentos.ListItems.Add(, , rst("cod_Medicamento"))
    itemx.SubItems(1) = rst("nome")
    itemx.SubItems(2) = rst("preco_Medicamento")
    itemx.Tag = rst("cod_Medicamento")
        
    rst.MoveNext
    Wend
    
End Sub

Private Sub ListMedicamentos_Click()

    Preencher_dados_Medicamentos
    
End Sub

Private Sub Preencher_dados_Medicamentos()

    SQL = " SELECT * FROM Medicamentos"
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
    SQL = " SELECT * " _
        & " FROM Medicamentos" _
        & " WHERE cod_Medicamento = " & ListMedicamentos.SelectedItem.Tag _
        & " ORDER BY nome"
        
    Set rst = db.OpenRecordset(SQL)
    
    SkinLabel117.Caption = rst("nome")
    SkinLabel118.Caption = rst("cod_Medicamento")
    SkinLabel119.Caption = rst("preco_Medicamento")
    If rst("comprimidos_caixa") = Null Then
        SkinLabel120.Caption = ""
    Else
        x = rst("comprimidos_caixa")
        SkinLabel120.Caption = "" & x
    End If
    SkinLabel121.Caption = rst("tipo_medicamento")
    SkinLabel122.Caption = rst("emblagem_disponiveis")
    SkinLabel123.Caption = rst("descricao")
    End If
    
End Sub

Private Sub ListMedicamentos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Preencher_dados_Medicamentos
    
End Sub

Private Sub cmdVender_Click()

    SQL = " SELECT * FROM imprimirfaturas"
    Set rst = db.OpenRecordset(SQL)
    SQL = "DELETE FROM imprimirfaturas"
        db.Execute SQL
    VendasMedicamentos.Show
    PainelControlo.Enabled = False
    
End Sub
'-------------------------------------------Fim do código para o separador "Medicamentos"-----------------------------

