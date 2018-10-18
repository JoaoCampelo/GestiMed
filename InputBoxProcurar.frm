VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form InputBoxProcurar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5670
   ControlBox      =   0   'False
   Icon            =   "InputBoxProcurar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Tag             =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtProcurar 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5205
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "InputBoxProcurar.frx":521A
      Top             =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Indique o nome da pessoa que pretende procurar."
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "InputBoxProcurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub CmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    Dim Procurar As String
    
    If InputBoxProcurar.Caption = "Procurar (Consultas)" Then
        Procurar = txtProcurar.Text
        SQL = " SELECT consultas.cod_cunsulta, consultas.data, consultas.hora, pacientes.nome_paciente " _
            & " FROM consultas, pacientes" _
            & " WHERE consultas.cod_paciente=pacientes.cod_paciente" _
            & " AND pacientes.nome_paciente LIKE '" & "*" & Procurar & "*" & "'" _
            & " ORDER BY data desc"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "A pessoa que procurou não tem consultas marcadas!"
            FormMsgBoxNormal.Caption = "Procurar (Consultas)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        PainelControlo.ListConsultas.ListItems.Clear
    
        While Not rst.EOF
            Set itemx = PainelControlo.ListConsultas.ListItems.Add(, , rst("cod_cunsulta"))
            itemx.SubItems(1) = rst("data")
            itemx.SubItems(2) = rst("hora")
            itemx.SubItems(3) = rst("nome_paciente")
            itemx.Tag = rst("cod_cunsulta")
            rst.MoveNext
        Wend
        PainelControlo.cmdProcurar.Visible = False
        PainelControlo.cmdRepor.Visible = True
        Unload Me
        PainelControlo.ListConsultas.SetFocus
        
    ElseIf InputBoxProcurar.Caption = "Procurar (Consultas de Tratamento)" Then
        Procurar = txtProcurar.Text
        SQL = " SELECT consultas_tratamentos.cod_consultatratamento, consultas_tratamentos.data, consultas_tratamentos.hora, consultas_tratamentos.cod_tratamento, pacientes.nome_paciente " _
            & " FROM consultas_tratamentos, pacientes" _
            & " WHERE consultas_tratamentos.cod_paciente=pacientes.cod_paciente" _
            & " AND pacientes.nome_paciente LIKE '" & "*" & Procurar & "*" & "'" _
            & " ORDER BY data desc"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "A pessoa que procurou não tem consultas de tratamento marcadas!"
            FormMsgBoxNormal.Caption = "Procurar (Consultas de Tratamento)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        PainelControlo.ListConsultasTratamento.ListItems.Clear
    
        While Not rst.EOF
            Set itemx = PainelControlo.ListConsultasTratamento.ListItems.Add(, , rst("cod_consultatratamento"))
            itemx.SubItems(1) = rst("data")
            itemx.SubItems(2) = rst("hora")
            itemx.SubItems(3) = rst("cod_tratamento")
            itemx.SubItems(4) = rst("nome_paciente")
            itemx.Tag = rst("cod_consultatratamento")
            rst.MoveNext
        Wend
        PainelControlo.cmdProcurar2.Visible = False
        PainelControlo.cmdRepor2.Visible = True
        Unload Me
        PainelControlo.ListConsultasTratamento.SetFocus
        
        ElseIf InputBoxProcurar.Caption = "Procurar (Pacientes)" Then
        Procurar = txtProcurar.Text
        SQL = " SELECT *" _
            & " FROM pacientes" _
            & " WHERE nome_paciente LIKE '" & "*" & Procurar & "*" & "'" _
            & " ORDER BY nome_paciente"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "A pessoa que procurou não existe no sistema!"
            FormMsgBoxNormal.Caption = "Procurar (Pacientes)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        PainelControlo.ListPacientes.ListItems.Clear
    
        While Not rst.EOF
            Set itemx = PainelControlo.ListPacientes.ListItems.Add(, , rst("cod_paciente"))
            itemx.SubItems(1) = rst("nome_paciente")
            itemx.SubItems(2) = rst("data_nascimento")
            itemx.Tag = rst("cod_paciente")
            rst.MoveNext
        Wend
        PainelControlo.cmdProcurar3.Visible = False
        PainelControlo.cmdRepor3.Visible = True
        Unload Me
        PainelControlo.ListPacientes.SetFocus
        
        ElseIf InputBoxProcurar.Caption = "Procurar (Médicos)" Then
        Procurar = txtProcurar.Text
        SQL = " SELECT *" _
            & " FROM Medicos" _
            & " WHERE nome_medico LIKE '" & "*" & Procurar & "*" & "'" _
            & " ORDER BY nome_medico"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "A pessoa que procurou não existe no sistema!"
            FormMsgBoxNormal.Caption = "Procurar (Médicos)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        PainelControlo.ListMedicos.ListItems.Clear
    
        While Not rst.EOF
            Set itemx = PainelControlo.ListMedicos.ListItems.Add(, , rst("cod_medico"))
            itemx.SubItems(1) = rst("nome_medico")
            itemx.SubItems(2) = rst("data_nascimento")
            itemx.Tag = rst("cod_medico")
            rst.MoveNext
        Wend
        PainelControlo.cmdProcurar4.Visible = False
        PainelControlo.cmdRepor4.Visible = True
        Unload Me
        PainelControlo.ListMedicos.SetFocus
        Exit Sub
    
    ElseIf InputBoxProcurar.Caption = "Procurar (Tratamentos)" Then
        Procurar = txtProcurar.Text
        SQL = " SELECT *" _
            & " FROM Tratamentos" _
            & " WHERE nome_tratamento LIKE '" & "*" & Procurar & "*" & "'" _
            & " ORDER BY nome_Tratamento"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O tratamento que procurou não existe no sistema!"
            FormMsgBoxNormal.Caption = "Procurar (Tratamentos)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        PainelControlo.ListTratamentos.ListItems.Clear
    
        While Not rst.EOF
            Set itemx = PainelControlo.ListTratamentos.ListItems.Add(, , rst("cod_tratamento"))
            itemx.SubItems(1) = rst("nome_tratamento")
            itemx.SubItems(2) = rst("preco_tratamento")
            itemx.Tag = rst("cod_tratamento")
            rst.MoveNext
        Wend
        PainelControlo.cmdProcurar5.Visible = False
        PainelControlo.cmdRepor5.Visible = True
        Unload Me
        PainelControlo.ListTratamentos.SetFocus
        Exit Sub
        
    ElseIf InputBoxProcurar.Caption = "Procurar (Medicamentos)" Then
        Procurar = txtProcurar.Text
        SQL = " SELECT *" _
            & " FROM Medicamentos" _
            & " WHERE nome LIKE '" & "*" & Procurar & "*" & "'" _
            & " ORDER BY nome"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O medicamento que procurou não existe no sistema!"
            FormMsgBoxNormal.Caption = "Procurar (Medicamentos)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        PainelControlo.ListMedicamentos.ListItems.Clear
    
        While Not rst.EOF
            Set itemx = PainelControlo.ListMedicamentos.ListItems.Add(, , rst("cod_Medicamento"))
            itemx.SubItems(1) = rst("nome")
            itemx.SubItems(2) = rst("preco_Medicamento")
            itemx.Tag = rst("cod_Medicamento")
            rst.MoveNext
        Wend
        PainelControlo.cmdProcurar6.Visible = False
        PainelControlo.cmdRepor6.Visible = True
        Unload Me
        PainelControlo.ListMedicamentos.SetFocus
        Exit Sub
    
    ElseIf InputBoxProcurar.Caption = "Procurar (Facturas)" Then
        ProcurarNome = txtProcurar.Text
        SQL = " SELECT * " _
            & " FROM faturas" _
            & " where nome_paciente LIKE '" & "*" & ProcurarNome & "*" & "'"
 
        Set rst = db.OpenRecordset(SQL)

        If rst.BOF = True And rst.EOF = True Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Não existem facturas com estes dados!"
            FormMsgBoxNormal.Caption = "Procurar (Facturas)"
            InputBoxProcurar.Enabled = False
            Exit Sub
        End If
        
        Facturas.ListFacturas.ListItems.Clear
    
        While Not rst.EOF
        Set itemx = Facturas.ListFacturas.ListItems.Add(, , rst("cod_fatura"))
        itemx.SubItems(1) = rst("cod_paciente")
        itemx.SubItems(2) = rst("nome_paciente")
        itemx.SubItems(3) = rst("data")
        itemx.SubItems(4) = rst("cod_medicamento")
        itemx.SubItems(5) = rst("nome_medicamento")
        itemx.SubItems(6) = rst("quantidade")
        itemx.SubItems(7) = rst("preco_unidade")
        itemx.Tag = rst("cod_fatura")
        rst.MoveNext
    Wend
        Facturas.cmdProcurar.Visible = False
        Facturas.cmdRepor.Visible = True
        Facturas.Enabled = True
        Facturas.ListFacturas.SetFocus
        Facturas.SetFocus
        Unload Me
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()

    Set db = OpenDatabase(App.Path & "\Clinica.mdb")

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
End Sub
