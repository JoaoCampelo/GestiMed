VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FormMsgBoxNormal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "FormMsgBoxNormal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FormMsgBoxNormal.frx":521A
      Top             =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "FormMsgBoxNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    
    If FormMsgBoxNormal.Caption = "Login" Then
        Login.Enabled = True
        Login.Text1.Text = ""
        Login.Text2.Text = ""
        Login.Text1.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Alterar (Consultas)") Or (FormMsgBoxNormal.Caption = "Eliminar (Consultas)") Then
        PainelControlo.Enabled = True
        PainelControlo.ListConsultas.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Alterar (Consultas de Tratamento)") Or (FormMsgBoxNormal.Caption = "Eliminar (Consultas de Tratamento)") Then
        PainelControlo.Enabled = True
        PainelControlo.ListConsultasTratamento.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Eliminar (Pacientes)") Then
        PainelControlo.Enabled = True
        PainelControlo.ListPacientes.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Eliminar (Médicos)") Then
        PainelControlo.Enabled = True
        PainelControlo.ListMedicos.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Eliminar (Tratamentos)") Then
        PainelControlo.Enabled = True
        PainelControlo.ListTratamentos.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Eliminar (Medicamentos)") Then
        PainelControlo.Enabled = True
        PainelControlo.ListMedicamentos.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas (Data)") Then
        InserirAlterarConsultas.Enabled = True
        InserirAlterarConsultas.Text2.Text = ""
        InserirAlterarConsultas.Text2.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas (Hora)") Then
        InserirAlterarConsultas.Enabled = True
        InserirAlterarConsultas.Text3.Text = ""
        InserirAlterarConsultas.Text3.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas (Nome do Paciente)") Then
        InserirAlterarConsultas.Enabled = True
        InserirAlterarConsultas.cmbPacientes.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas (Nome do Médico)") Then
        InserirAlterarConsultas.Enabled = True
        InserirAlterarConsultas.cmbMedicos.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas (Alterar)") Then
        InserirAlterarConsultas.Enabled = True
        PainelControlo.ListConsultas.SetFocus
        Unload InserirAlterarConsultas
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas de Tratamento (Data)") Then
        InserirAlterarConsultasTratamento.Enabled = True
        InserirAlterarConsultasTratamento.Text2.Text = ""
        InserirAlterarConsultasTratamento.Text2.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas de Tratamento (Hora)") Then
        InserirAlterarConsultasTratamento.Enabled = True
        InserirAlterarConsultasTratamento.Text3.Text = ""
        InserirAlterarConsultasTratamento.Text3.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas de Tratamento (Nome do Tratamento)") Then
        InserirAlterarConsultasTratamento.Enabled = True
        InserirAlterarConsultasTratamento.cmbPacientes.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas de Tratamento (Nome do Paciente)") Then
        InserirAlterarConsultasTratamento.Enabled = True
        InserirAlterarConsultasTratamento.cmbPacientes.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas de Tratamento (Nome do Médico)") Then
        InserirAlterarConsultasTratamento.Enabled = True
        InserirAlterarConsultasTratamento.cmbMedicos.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Consultas de Tratamento (Alterar)") Then
        InserirAlterarConsultasTratamento.Enabled = True
        PainelControlo.ListConsultasTratamento.SetFocus
        Unload InserirAlterarConsultasTratamento
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Nome)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text1.Text = ""
        InserirAlterarMedicos.Text1.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Data)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text3.Text = ""
        InserirAlterarMedicos.Text3.SetFocus
        Unload Me
        Exit Sub
    
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Telefone)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text4.Text = ""
        InserirAlterarMedicos.Text4.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Telemóvel)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text5.Text = ""
        InserirAlterarMedicos.Text5.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (E-Mail)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text6.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Bilhete de Identidade)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text7.Text = ""
        InserirAlterarMedicos.Text7.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Número de Contribuinte)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text8.Text = ""
        InserirAlterarMedicos.Text8.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Morada)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text9.Text = ""
        InserirAlterarMedicos.Text9.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Código Postal)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text10.Text = ""
        InserirAlterarMedicos.Text10.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Cidade)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text11.Text = ""
        InserirAlterarMedicos.Text11.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Sexo)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text12.Text = ""
        InserirAlterarMedicos.Text12.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Estado Civil)") Then
        InserirAlterarMedicos.Enabled = True
        InserirAlterarMedicos.Text13.Text = ""
        InserirAlterarMedicos.Text13.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Médicos (Alterar)") Then
        InserirAlterarMedicos.Enabled = True
        PainelControlo.ListMedicos.SetFocus
        Unload InserirAlterarMedicos
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Nome)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text1.Text = ""
        InserirAlterarPacientes.Text1.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Data)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text3.Text = ""
        InserirAlterarPacientes.Text3.SetFocus
        Unload Me
        Exit Sub
    
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Telefone)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text4.Text = ""
        InserirAlterarPacientes.Text4.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Telemóvel)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text5.Text = ""
        InserirAlterarPacientes.Text5.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (E-Mail)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text6.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Bilhete de Identidade)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text7.Text = ""
        InserirAlterarPacientes.Text7.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Número de Contribuinte)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text8.Text = ""
        InserirAlterarPacientes.Text8.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Morada)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text9.Text = ""
        InserirAlterarPacientes.Text9.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Código Postal)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text10.Text = ""
        InserirAlterarPacientes.Text10.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Cidade)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text11.Text = ""
        InserirAlterarPacientes.Text11.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Sexo)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text12.Text = ""
        InserirAlterarPacientes.Text12.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Estado Civil)") Then
        InserirAlterarPacientes.Enabled = True
        InserirAlterarPacientes.Text13.Text = ""
        InserirAlterarPacientes.Text13.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Pacientes (Alterar)") Then
        InserirAlterarPacientes.Enabled = True
        PainelControlo.ListPacientes.SetFocus
        Unload InserirAlterarPacientes
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Tratamentos (Nome)") Then
        InserirAlterarTratamentos.Enabled = True
        InserirAlterarTratamentos.Text1.Text = ""
        InserirAlterarTratamentos.Text1.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Tratamentos (Preço do Tratamento)") Then
        InserirAlterarTratamentos.Enabled = True
        InserirAlterarTratamentos.Text3.Text = ""
        InserirAlterarTratamentos.Text3.SetFocus
        Unload Me
        Exit Sub
    
    ElseIf (FormMsgBoxNormal.Caption = "Tratamentos (Descrição)") Then
        InserirAlterarTratamentos.Enabled = True
        InserirAlterarTratamentos.Text4.Text = ""
        InserirAlterarTratamentos.Text4.SetFocus
        Unload Me
        Exit Sub
    
    ElseIf (FormMsgBoxNormal.Caption = "Tratamentos (Alterar)") Then
        InserirAlterarTratamentos.Enabled = True
        PainelControlo.ListTratamentos.SetFocus
        Unload InserirAlterarTratamentos
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Procurar (Consultas)") Then
        InputBoxProcurar.Enabled = True
        PainelControlo.ListConsultas.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Procurar (Consultas de Tratamento)") Then
        InputBoxProcurar.Enabled = True
        PainelControlo.ListConsultasTratamento.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Procurar (Pacientes)") Then
        InputBoxProcurar.Enabled = True
        PainelControlo.ListPacientes.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
        
     ElseIf (FormMsgBoxNormal.Caption = "Procurar (Médicos)") Then
        InputBoxProcurar.Enabled = True
        PainelControlo.ListMedicos.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Procurar (Tratamentos)") Then
        InputBoxProcurar.Enabled = True
        PainelControlo.ListTratamentos.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
    
    ElseIf (FormMsgBoxNormal.Caption = "Procurar (Medicamentos)") Then
        InputBoxProcurar.Enabled = True
        PainelControlo.ListMedicamentos.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Procurar (Facturas)") Then
        InputBoxProcurar.Enabled = True
        Facturas.Enabled = True
        Facturas.ListFacturas.SetFocus
        Unload InputBoxProcurar
        Unload Me
        Exit Sub
    
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Nome)") Then
        InserirAlterarMedicamentos.Enabled = True
        InserirAlterarMedicamentos.Text1.Text = ""
        InserirAlterarMedicamentos.Text1.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Preço do Medicamento)") Then
        InserirAlterarMedicamentos.Enabled = True
        InserirAlterarMedicamentos.Text3.Text = ""
        InserirAlterarMedicamentos.Text3.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Comprimidos por Caixa)") Then
        InserirAlterarMedicamentos.Enabled = True
        InserirAlterarMedicamentos.Text4.Text = ""
        InserirAlterarMedicamentos.Text4.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Tipo de Medicamento)") Then
        InserirAlterarMedicamentos.Enabled = True
        InserirAlterarMedicamentos.Text5.Text = ""
        InserirAlterarMedicamentos.Text5.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Embalagens Disponiveis)") Then
        InserirAlterarMedicamentos.Enabled = True
        InserirAlterarMedicamentos.Text6.Text = ""
        InserirAlterarMedicamentos.Text6.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Descrição)") Then
        InserirAlterarMedicamentos.Enabled = True
        InserirAlterarMedicamentos.Text7.Text = ""
        InserirAlterarMedicamentos.Text7.SetFocus
        Unload Me
        Exit Sub
        
    ElseIf (FormMsgBoxNormal.Caption = "Medicamentos (Alterar)") Then
        InserirAlterarMedicamentos.Enabled = True
        PainelControlo.ListMedicamentos.SetFocus
        Unload InserirAlterarMedicamentos
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Novo Utilizador (Utilizador)") Then
        NovoUtilizador.Enabled = True
        NovoUtilizador.Text1.Text = ""
        NovoUtilizador.Text1.SetFocus
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Novo Utilizador (Password)") Then
        NovoUtilizador.Enabled = True
        NovoUtilizador.Text2.Text = ""
        NovoUtilizador.Text3.Text = ""
        NovoUtilizador.Text2.SetFocus
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Novo Utilizador (Password do Sistema)") Then
        NovoUtilizador.Enabled = True
        NovoUtilizador.Text4.Text = ""
        NovoUtilizador.Text4.SetFocus
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Novo Utilizador (Login)") Then
        Login.Show
        Unload NovoUtilizador
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Novo Utilizador (GestiMed)") Then
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        Unload NovoUtilizador
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Desbloquear") Then
        Desbloquear.Text1.Text = ""
        Desbloquear.Text2.Text = ""
        Desbloquear.Enabled = True
        Desbloquear.SetFocus
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Alterar Password (Utilizador/Password)") Then
        AlterarPassword.Enabled = True
        AlterarPassword.Text1.SetFocus
        AlterarPassword.Text1.Text = ""
        AlterarPassword.Text2.Text = ""
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Alterar Password (Nova Password)") Then
        AlterarPassword.Enabled = True
        AlterarPassword.Text3.SetFocus
        AlterarPassword.Text3.Text = ""
        AlterarPassword.Text4.Text = ""
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Alterar Password (Êxito)") Then
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        Unload AlterarPassword
        Unload Me
        Exit Sub
    ElseIf (FormMsgBoxNormal.Caption = "Stock de Medicamentos") Then
        VendasMedicamentos.Enabled = True
        VendasMedicamentos.Text1.Text = ""
        VendasMedicamentos.cmbMedicamentos.Enabled = False
        VendasMedicamentos.Text1.SetFocus
        Unload Me
    ElseIf (FormMsgBoxNormal.Caption = "Quantidade Elevada") Then
        VendasMedicamentos.Enabled = True
        VendasMedicamentos.Text1.Text = ""
        VendasMedicamentos.cmbMedicamentos.Enabled = False
        VendasMedicamentos.Text1.SetFocus
        Unload Me
    ElseIf (FormMsgBoxNormal.Caption = "Stock a Esgotar") Then
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        Unload Login
        Unload Me
    ElseIf (FormMsgBoxNormal.Caption = "Backup") Then
        End
    End If
    
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd

End Sub
