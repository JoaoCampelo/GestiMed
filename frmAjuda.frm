VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuda"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ChoiceFrame 
      Height          =   4815
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Width           =   7215
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   1335
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0000
         TabIndex        =   18
         Top             =   600
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":005E
         TabIndex        =   19
         Top             =   240
         Width           =   5655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":010A
         TabIndex        =   20
         Top             =   2040
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   1215
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":01BA
         TabIndex        =   21
         Top             =   2400
         Width           =   6735
      End
   End
   Begin VB.Frame ChoiceFrame 
      Height          =   4815
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   7215
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   1095
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0218
         TabIndex        =   13
         Top             =   600
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0276
         TabIndex        =   14
         Top             =   240
         Width           =   5655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":033E
         TabIndex        =   15
         Top             =   2640
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   615
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":03FA
         TabIndex        =   16
         Top             =   2040
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0458
         TabIndex        =   22
         Top             =   1680
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   1215
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0516
         TabIndex        =   23
         Top             =   3000
         Width           =   6735
      End
   End
   Begin VB.Frame ChoiceFrame 
      Height          =   4815
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   7215
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   1095
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0574
         TabIndex        =   8
         Top             =   600
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":05D2
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0664
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   1575
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":06DE
         TabIndex        =   11
         Top             =   2160
         Width           =   6735
      End
   End
   Begin VB.Frame ChoiceFrame 
      Height          =   4815
      Index           =   0
      Left            =   -120
      TabIndex        =   2
      Top             =   -120
      Width           =   7215
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   1095
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":073C
         TabIndex        =   4
         Top             =   600
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":079A
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":0812
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   1215
         Left            =   240
         OleObjectBlob   =   "frmAjuda.frx":088A
         TabIndex        =   6
         Top             =   2160
         Width           =   6735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmAjuda.frx":08E8
      Top             =   0
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8916
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      MultiSelect     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inserir/Alterar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Segurança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Password do Sistema"
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
End
Attribute VB_Name = "frmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectedTab As Integer

Private Sub cmdOK_Click()

    PainelControlo.Enabled = True
    PainelControlo.SetFocus
    Unload Me

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

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    SkinLabel2.Caption = "   Para inserir dados no programa basta ir ao seperador onde pretende inserir os dados e clicar no botão Inserir, ai irá aparecer uma nova janela onde pode introduzir os dados." & vbCr & "   Depois de introduzir os dados basta clicar no botão Inserir da nova janela e os dados são guardados na base de dados." & vbCr & "   Por questões de segurança alguns dos dados são inseridos automaticamente pelo programa."
    SkinLabel4.Caption = "   Para alterar os dados basta ir ao seperador onde pretende alterar os dados, seleccionar os dados a serem alterados e clicar no botão Alterar, ai irá aparecer uma nova janela com os dados que seleccionou." & vbCr & "    Depois de alterar os dados basta clicar no botão Alterar da nova janela e a alteração será guardada na base de dados." & vbCr & "   Por questões de segurança alguns dos dados não podem ser alterados."
    SkinLabel5.Caption = "   Sim, é possivel eliminar dados da base de dados, mas nem todos os dados podem ser eliminados. Por motivos de segurança existem várias regras no programa que definem os dados que podem ser eliminados e os dados que não podem ser eliminados. "
    SkinLabel8.Caption = "   Para eliminar os dados basta ir ao separador onde pretende eliminar os dados, seleccionar os dados a serem eliminados e clicar no botão Eliminar, ai irá aparecer uma janela para você comfirmar se quer realmente eliminar aqueles dados." & vbCr & "   Se os dados poderem ser eliminados irá aparecer mais uma janela a confirmar que os dados foram eliminados, caso contrário irá aparecer uma janela a dizer que os dados não podem ser eliminados e qual o motivo."
    SkinLabel9.Caption = "   Sim. Para introduzir um novo utilizador no programa existem duas maneiras." & vbCr & "   A primeira e logo a entrar no programa onde se faz login basta clicar no botão Novo Utilizadore adicionar o utilizador, a segunda forma e já no ménu principal ir a Opções, Segurança, Novo Utilizador e introduzir o utilizador."
    SkinLabel12.Caption = "   Não. Para Inserir um novo utilizador é preciso uma password do sistema." & vbCr & "   Essa password é apenas fornecida ao comprador do software."
    SkinLabel13.Caption = "   A password do sistema é fornecida apenas ao comprador do software." & vbCr & "    Para aceder a esta password o comprador do software tem de introduzir o cd de instalação no computador e em seguida tem de ir ao Meu Computador, clicar com o botão direito do rato em cima do cd e seleccionar Abrir." & vbCr & "  Ai irá encontrar um ficheiro de texto com o nome Password so Sistema, é so abrir esse ficheiro e tem acesso á password."
    SkinLabel16.Caption = "   Não. Por motivos de segurança não é possivel alterar esta password."
    SkinLabel18.Caption = "   Sim. Basta no ménu principal clicar em Opções, Segurança, Alterar Password, e preencher os dados lá pedidos." & vbCr & "    Depois de preencher os dados basta clicar no botão Alterar Password e a sua password será alterada."

End Sub

Private Sub TabStrip1_Click()

    ChoiceFrame(SelectedTab - 1).Visible = False
    SelectedTab = TabStrip1.SelectedItem.Index
    ChoiceFrame(SelectedTab - 1).Visible = True
    
End Sub
