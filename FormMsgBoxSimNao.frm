VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FormMsgBoxSimNao 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "FormMsgBoxSimNao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNao 
      Caption         =   "Não"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FormMsgBoxSimNao.frx":521A
      Top             =   480
   End
   Begin VB.CommandButton cmdSim 
      Caption         =   "Sim"
      Height          =   375
      Left            =   960
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
Attribute VB_Name = "FormMsgBoxSimNao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub cmdNao_Click()

    If FormMsgBoxSimNao.Caption = "Eliminar (Consultas)" Then
        PainelControlo.Enabled = True
        PainelControlo.ListConsultas.SetFocus
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Consultas de Tratamento)" Then
        PainelControlo.Enabled = True
        PainelControlo.ListConsultasTratamento.SetFocus
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Pacientes)" Then
        PainelControlo.Enabled = True
        PainelControlo.ListPacientes.SetFocus
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Médicos)" Then
        PainelControlo.Enabled = True
        PainelControlo.ListMedicos.SetFocus
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Tratamentos)" Then
        PainelControlo.Enabled = True
        PainelControlo.ListTratamentos.SetFocus
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Medicamentos)" Then
        PainelControlo.Enabled = True
        PainelControlo.ListMedicamentos.SetFocus
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Consultas)" Then
        ListConsultas_Ordena_SetUp
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        PainelControlo.ListConsultas.SetFocus
        Unload InserirAlterarConsultas
        SQL = "select * from consultas"
        Set rst = db.OpenRecordset(SQL)
        If Not (rst.BOF = True And rst.EOF = True) Then
            PainelControlo.cmdAlterar.Enabled = True
            PainelControlo.cmdDel.Enabled = True
        End If
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Consultas de Tratamento)" Then
        ListConsultasTratamentos_Ordena_SetUp
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        PainelControlo.ListConsultasTratamento.SetFocus
        Unload InserirAlterarConsultasTratamento
        SQL = "select * from consultas_tratamentos"
        Set rst = db.OpenRecordset(SQL)
        If Not (rst.BOF = True And rst.EOF = True) Then
            PainelControlo.cmdAlterar2.Enabled = True
            PainelControlo.cmdDel2.Enabled = True
        End If
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Médicos)" Then
        ListMedicos_Ordena_SetUp
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        PainelControlo.ListMedicos.SetFocus
        Unload InserirAlterarMedicos
        SQL = "select * from medicos"
        Set rst = db.OpenRecordset(SQL)
        If Not (rst.BOF = True And rst.EOF = True) Then
            PainelControlo.cmdAlterar4.Enabled = True
            PainelControlo.cmdDel4.Enabled = True
        End If
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Pacientes)" Then
        ListPacientes_Ordena_SetUp
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        PainelControlo.ListPacientes.SetFocus
        Unload InserirAlterarPacientes
        SQL = "select * from pacientes"
        Set rst = db.OpenRecordset(SQL)
        If Not (rst.BOF = True And rst.EOF = True) Then
            PainelControlo.cmdAlterar3.Enabled = True
            PainelControlo.cmdDel3.Enabled = True
        End If
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Tratamentos)" Then
        ListTratamentos_Ordena_SetUp
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        PainelControlo.ListTratamentos.SetFocus
        Unload InserirAlterarTratamentos
        SQL = "select * from tratamentos"
        Set rst = db.OpenRecordset(SQL)
        If Not (rst.BOF = True And rst.EOF = True) Then
            PainelControlo.cmdAlterar5.Enabled = True
            PainelControlo.cmdDel5.Enabled = True
        End If
        Unload Me
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Medicamentos)" Then
        ListMedicamentos_Ordena_SetUp
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        PainelControlo.ListMedicamentos.SetFocus
        Unload InserirAlterarMedicamentos
        SQL = "select * from medicamentos"
        Set rst = db.OpenRecordset(SQL)
        If Not (rst.BOF = True And rst.EOF = True) Then
            PainelControlo.cmdAlterar6.Enabled = True
            PainelControlo.cmdDel6.Enabled = True
            PainelControlo.cmdVender.Enabled = True
        End If
        Unload Me
        
    End If
    
End Sub

Private Sub cmdSim_Click()

    If FormMsgBoxSimNao.Caption = "Eliminar (Consultas)" Then
    
        SQL = " SELECT * FROM Consultas" _
        & " WHERE cod_cunsulta=" & PainelControlo.ListConsultas.SelectedItem.Tag
        
        Set rst = db.OpenRecordset(SQL)
    
        If rst("Data") > DateValue(PainelControlo.Label18.Caption) Then
            SQL = "DELETE FROM Consultas WHERE cod_cunsulta = " & PainelControlo.ListConsultas.SelectedItem.Tag
            db.Execute SQL
            PainelControlo.ListConsultas.ListItems.Remove PainelControlo.ListConsultas.SelectedItem.Index
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "A consulta foi eliminada com exito!"
            FormMsgBoxNormal.Caption = "Eliminar (Consultas)"
            SQL = "select * from consultas"
            Set rst = db.OpenRecordset(SQL)
            If rst.BOF = True And rst.EOF = True Then
                PainelControlo.cmdAlterar.Enabled = False
                PainelControlo.cmdDel.Enabled = False
            End If
            Unload Me
            Exit Sub
        Else
            If rst("Data") = DateValue(PainelControlo.Label18.Caption) Then
                If rst("hora") > TimeValue(PainelControlo.Label19.Caption) Then
                    SQL = "DELETE FROM Consultas WHERE cod_cunsulta = " & PainelControlo.ListConsultas.SelectedItem.Tag
                    db.Execute SQL
                    PainelControlo.ListConsultas.ListItems.Remove PainelControlo.ListConsultas.SelectedItem.Index
                    FormMsgBoxNormal.Show
                    FormMsgBoxNormal.Label1.Caption = "A consulta foi eliminada com exito!"
                    FormMsgBoxNormal.Caption = "Eliminar (Consultas)"
                    SQL = "select * from consultas"
                    Set rst = db.OpenRecordset(SQL)
                    If rst.BOF = True And rst.EOF = True Then
                        PainelControlo.cmdAlterar.Enabled = False
                        PainelControlo.cmdDel.Enabled = False
                    End If
                    Unload Me
                    Exit Sub
                Else
                    FormMsgBoxNormal.Show
                    FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser eliminada porque já foi realizada!"
                    FormMsgBoxNormal.Caption = "Eliminar (Consultas)"
                    Unload Me
                Exit Sub
                End If
            End If
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser eliminada porque já foi realizada!"
                FormMsgBoxNormal.Caption = "Eliminar (Consultas)"
                Unload Me
        End If
        Exit Sub
    
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Consultas de Tratamento)" Then
    
        SQL = " SELECT * FROM Consultas_tratamentos" _
        & " WHERE cod_consultatratamento = " & PainelControlo.ListConsultasTratamento.SelectedItem.Tag
        
        Set rst = db.OpenRecordset(SQL)
    
        If rst("Data") > DateValue(PainelControlo.Label18.Caption) Then
            SQL = "DELETE FROM Consultas_tratamentos WHERE cod_consultatratamento = " & PainelControlo.ListConsultasTratamento.SelectedItem.Tag
            db.Execute SQL
            PainelControlo.ListConsultasTratamento.ListItems.Remove PainelControlo.ListConsultasTratamento.SelectedItem.Index
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "A consulta foi eliminada com exito!"
            FormMsgBoxNormal.Caption = "Eliminar (Consultas de Tratamento)"
            SQL = "select * from consultas_tratamentos"
            Set rst = db.OpenRecordset(SQL)
            If rst.BOF = True And rst.EOF = True Then
                PainelControlo.cmdAlterar2.Enabled = False
                PainelControlo.cmdDel2.Enabled = False
            End If
            Unload Me
            Exit Sub
        Else
            If rst("Data") = DateValue(PainelControlo.Label18.Caption) Then
                If rst("hora") > TimeValue(PainelControlo.Label19.Caption) Then
                    SQL = "DELETE FROM Consultas_tratamentos WHERE cod_consultatratamento = " & PainelControlo.ListConsultasTratamento.SelectedItem.Tag
                    db.Execute SQL
                    PainelControlo.ListConsultasTratamento.ListItems.Remove PainelControlo.ListConsultasTratamento.SelectedItem.Index
                    FormMsgBoxNormal.Show
                    FormMsgBoxNormal.Label1.Caption = "A consulta foi eliminada com exito!"
                    FormMsgBoxNormal.Caption = "Eliminar (Consultas de Tratamento)"
                    SQL = "select * from consultas_tratamentos"
                    Set rst = db.OpenRecordset(SQL)
                    If rst.BOF = True And rst.EOF = True Then
                        PainelControlo.cmdAlterar2.Enabled = False
                        PainelControlo.cmdDel2.Enabled = False
                    End If
                    Unload Me
                    Exit Sub
                Else
                    FormMsgBoxNormal.Show
                    FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser eliminada porque já foi realizada!"
                    FormMsgBoxNormal.Caption = "Eliminar (Consultas de Tratamento)"
                    Unload Me
                    Exit Sub
                End If
            End If
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Esta consulta não pode ser eliminada porque já foi realizada!"
            FormMsgBoxNormal.Caption = "Eliminar (Consultas de Tratamento)"
            Unload Me
        End If
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Pacientes)" Then
    
        SQL = " SELECT * FROM pacientes" _
        & " WHERE cod_paciente = " & PainelControlo.ListPacientes.SelectedItem.Tag
        
        Set rst = db.OpenRecordset(SQL)
        
        SQL = "DELETE FROM Pacientes WHERE cod_paciente = " & PainelControlo.ListPacientes.SelectedItem.Tag
            db.Execute SQL
        PainelControlo.ListPacientes.ListItems.Remove PainelControlo.ListPacientes.SelectedItem.Index
            
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O paciente foi eliminado com exito!"
        FormMsgBoxNormal.Caption = "Eliminar (Pacientes)"
        SQL = "select * from pacientes"
        Set rst = db.OpenRecordset(SQL)
        If rst.BOF = True And rst.EOF = True Then
            PainelControlo.cmdAlterar3.Enabled = False
            PainelControlo.cmdDel3.Enabled = False
        End If
        Unload Me
        Exit Sub
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Médicos)" Then
    
        SQL = " SELECT * FROM Medicos" _
        & " WHERE cod_medico = " & PainelControlo.ListMedicos.SelectedItem.Tag
        
        Set rst = db.OpenRecordset(SQL)
        
        SQL = "DELETE FROM Medicos WHERE cod_medico = " & PainelControlo.ListMedicos.SelectedItem.Tag
            db.Execute SQL
        PainelControlo.ListMedicos.ListItems.Remove PainelControlo.ListMedicos.SelectedItem.Index
            
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O médico foi eliminado com exito!"
        FormMsgBoxNormal.Caption = "Eliminar (Médicos)"
        SQL = "select * from medicos"
        Set rst = db.OpenRecordset(SQL)
        If rst.BOF = True And rst.EOF = True Then
            PainelControlo.cmdAlterar4.Enabled = False
            PainelControlo.cmdDel4.Enabled = False
        End If
        Unload Me
        Exit Sub
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Tratamentos)" Then
    
        SQL = " SELECT * FROM Tratamentos" _
        & " WHERE cod_Tratamento = " & PainelControlo.ListTratamentos.SelectedItem.Tag
        
        Set rst = db.OpenRecordset(SQL)
        
        SQL = "DELETE FROM Tratamentos WHERE cod_tratamento = " & PainelControlo.ListTratamentos.SelectedItem.Tag
            db.Execute SQL
        PainelControlo.ListTratamentos.ListItems.Remove PainelControlo.ListTratamentos.SelectedItem.Index
            
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O tratamento foi eliminado com exito!"
        FormMsgBoxNormal.Caption = "Eliminar (Tratamentos)"
        SQL = "select * from tratamentos"
        Set rst = db.OpenRecordset(SQL)
        If rst.BOF = True And rst.EOF = True Then
            PainelControlo.cmdAlterar5.Enabled = False
            PainelControlo.cmdDel5.Enabled = False
        End If
        Unload Me
        Exit Sub
        
    ElseIf FormMsgBoxSimNao.Caption = "Eliminar (Medicamentos)" Then
    
        SQL = " SELECT * FROM Medicamentos" _
        & " WHERE cod_Medicamento = " & PainelControlo.ListMedicamentos.SelectedItem.Tag
        
        Set rst = db.OpenRecordset(SQL)
        
        SQL = "DELETE FROM Medicamentos WHERE cod_Medicamento = " & PainelControlo.ListMedicamentos.SelectedItem.Tag
            db.Execute SQL
        PainelControlo.ListMedicamentos.ListItems.Remove PainelControlo.ListMedicamentos.SelectedItem.Index
            
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O medicamento foi eliminado com exito!"
        FormMsgBoxNormal.Caption = "Eliminar (Medicamentos)"
        SQL = "select * from medicamentos"
        Set rst = db.OpenRecordset(SQL)
        If rst.BOF = True And rst.EOF = True Then
            PainelControlo.cmdAlterar6.Enabled = False
            PainelControlo.cmdDel6.Enabled = False
            PainelControlo.cmdVender.Enabled = False
        End If
        Unload Me
        Exit Sub
        
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Consultas)" Then
    
        InserirAlterarConsultas.Enabled = True
        
        SQL = " SELECT * FROM Consultas"
            Set rst = db.OpenRecordset(SQL)
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
            InserirAlterarConsultas.Text1.Text = Val(res) + 1
            InserirAlterarConsultas.Text2.Text = ""
            InserirAlterarConsultas.Text3.Text = ""
            InserirAlterarConsultas.cmbPacientes.Text = "Selecione o Paciente"
            InserirAlterarConsultas.Text5.Text = ""
            InserirAlterarConsultas.Text6.Text = ""
            InserirAlterarConsultas.Text7.Text = ""
            InserirAlterarConsultas.Text10.Text = ""
            InserirAlterarConsultas.cmbMedicos.Text = "Selecione o Médico"
            InserirAlterarConsultas.Text9.Text = ""
            InserirAlterarConsultas.Text2.SetFocus
            Unload Me
            
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Consultas de Tratamento)" Then
        
            InserirAlterarConsultasTratamento.Enabled = True
            
            SQL = " SELECT * FROM Consultas_Tratamentos"
            Set rst = db.OpenRecordset(SQL)
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
            InserirAlterarConsultasTratamento.Text2.Text = ""
            InserirAlterarConsultasTratamento.Text3.Text = ""
            InserirAlterarConsultasTratamento.cmbTratamento.Text = "Selecione o Tratamento"
            InserirAlterarConsultasTratamento.Text4.Text = ""
            InserirAlterarConsultasTratamento.Text5.Text = ""
            InserirAlterarConsultasTratamento.cmbPacientes.Text = "Selecione o Paciente"
            InserirAlterarConsultasTratamento.Text6.Text = ""
            InserirAlterarConsultasTratamento.Text7.Text = ""
            InserirAlterarConsultasTratamento.Text8.Text = ""
            InserirAlterarConsultasTratamento.Text9.Text = ""
            InserirAlterarConsultasTratamento.cmbMedicos.Text = "Selecione o Médico"
            InserirAlterarConsultasTratamento.Text10.Text = ""
            InserirAlterarConsultasTratamento.Text2.SetFocus
            Unload Me
            
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Médicos)" Then
            
            InserirAlterarMedicos.Enabled = True
            
            SQL = " SELECT * FROM seg_medicos"
            Set rst = db.OpenRecordset(SQL)
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
            InserirAlterarMedicos.Text1.Text = ""
            InserirAlterarMedicos.Text2.Text = Val(res) + 1
            InserirAlterarMedicos.Text3.Text = ""
            InserirAlterarMedicos.Text4.Text = ""
            InserirAlterarMedicos.Text5.Text = ""
            InserirAlterarMedicos.Text6.Text = ""
            InserirAlterarMedicos.Text7.Text = ""
            InserirAlterarMedicos.Text8.Text = ""
            InserirAlterarMedicos.Text9.Text = ""
            InserirAlterarMedicos.Text10.Text = ""
            InserirAlterarMedicos.Text11.Text = ""
            InserirAlterarMedicos.Text12.Text = ""
            InserirAlterarMedicos.Text13.Text = ""
            InserirAlterarMedicos.Text1.SetFocus
            Unload Me
            
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Pacientes)" Then
            
            InserirAlterarPacientes.Enabled = True
            
            SQL = " SELECT * FROM seg_pacientes"
            Set rst = db.OpenRecordset(SQL)
            With rst
                If Not .EOF Then
                    Do While Not .EOF
                        x = rst("cod_paciente")
                        .MoveNext
                    Loop
                Else
                    y = rst("cod_paciente")
                End If
            End With
            res = x
            InserirAlterarPacientes.Text1.Text = ""
            InserirAlterarPacientes.Text2.Text = Val(res) + 1
            InserirAlterarPacientes.Text3.Text = ""
            InserirAlterarPacientes.Text4.Text = ""
            InserirAlterarPacientes.Text5.Text = ""
            InserirAlterarPacientes.Text6.Text = ""
            InserirAlterarPacientes.Text7.Text = ""
            InserirAlterarPacientes.Text8.Text = ""
            InserirAlterarPacientes.Text9.Text = ""
            InserirAlterarPacientes.Text10.Text = ""
            InserirAlterarPacientes.Text11.Text = ""
            InserirAlterarPacientes.Text12.Text = ""
            InserirAlterarPacientes.Text13.Text = ""
            InserirAlterarPacientes.Text14.Text = ""
            InserirAlterarPacientes.Text1.SetFocus
            Unload Me
            
    ElseIf FormMsgBoxSimNao.Caption = "Inserir (Tratamentos)" Then
        
            InserirAlterarTratamentos.Enabled = True
            
            SQL = " SELECT * FROM seg_Tratamentos"
            Set rst = db.OpenRecordset(SQL)
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
            InserirAlterarTratamentos.Text1.Text = ""
            InserirAlterarTratamentos.Text2.Text = Val(res) + 1
            InserirAlterarTratamentos.Text3.Text = ""
            InserirAlterarTratamentos.Text4.Text = ""
            InserirAlterarTratamentos.Text1.SetFocus
            Unload Me
            
        ElseIf FormMsgBoxSimNao.Caption = "Inserir (Medicamentos)" Then
        
            InserirAlterarMedicamentos.Enabled = True
            
            SQL = " SELECT * FROM Medicamentos"
            Set rst = db.OpenRecordset(SQL)
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
            InserirAlterarMedicamentos.Text1.Text = ""
            InserirAlterarMedicamentos.Text2.Text = Val(res) + 1
            InserirAlterarMedicamentos.Text3.Text = ""
            InserirAlterarMedicamentos.Text4.Text = ""
            InserirAlterarMedicamentos.Text5.Text = ""
            InserirAlterarMedicamentos.Text6.Text = ""
            InserirAlterarMedicamentos.Text7.Text = ""
            InserirAlterarMedicamentos.Text1.SetFocus
            Unload Me
       
    End If

End Sub

Private Sub Form_Load()

    Set db = OpenDatabase(App.Path & "\Clinica.mdb")

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd

End Sub

Private Sub ListConsultas_Ordena_SetUp()
    
    PainelControlo.ListConsultas.ListItems.Clear
    
    SQL = " SELECT consultas.cod_cunsulta, consultas.data, consultas.hora, pacientes.nome_paciente " _
        & " FROM consultas, pacientes" _
        & " WHERE consultas.cod_paciente=pacientes.cod_paciente" _
        & " ORDER BY data desc"
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    While Not rst.EOF
        Set itemx = PainelControlo.ListConsultas.ListItems.Add(, , rst("cod_cunsulta"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("nome_paciente")
        itemx.Tag = rst("cod_cunsulta")
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListConsultasTratamentos_Ordena_SetUp()

    Dim itemx As ListItem

    PainelControlo.ListConsultasTratamento.ListItems.Clear
    
    SQL = " SELECT consultas_tratamentos.cod_consultatratamento, consultas_tratamentos.data, consultas_tratamentos.hora, consultas_tratamentos.cod_tratamento, pacientes.nome_paciente " _
        & " FROM consultas_tratamentos, pacientes" _
        & " WHERE consultas_tratamentos.cod_paciente=pacientes.cod_paciente" _
        & " ORDER BY consultas_tratamentos.data desc"
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListConsultasTratamento.ListItems.Add(, , rst("cod_consultatratamento"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("cod_tratamento")
        itemx.SubItems(4) = rst("nome_paciente")
        itemx.Tag = rst("cod_consultatratamento")
        
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListMedicos_Ordena_SetUp()

    Dim itemx As ListItem
    PainelControlo.ListMedicos.ListItems.Clear
    
    SQL = " SELECT * " _
        & " FROM Medicos" _
        & " ORDER BY nome_medico"
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListMedicos.ListItems.Add(, , rst("cod_medico"))
        itemx.SubItems(1) = rst("nome_medico")
        itemx.SubItems(2) = rst("data_nascimento")
        itemx.Tag = rst("cod_medico")
        
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListPacientes_Ordena_SetUp()

    Dim itemx As ListItem
    PainelControlo.ListPacientes.ListItems.Clear
    
    SQL = " SELECT * " _
        & " FROM Pacientes" _
        & " ORDER BY nome_paciente "
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListPacientes.ListItems.Add(, , rst("cod_paciente"))
        itemx.SubItems(1) = rst("nome_paciente")
        itemx.SubItems(2) = rst("data_nascimento")
        itemx.Tag = rst("cod_paciente")
        
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListTratamentos_Ordena_SetUp()

    Dim itemx As ListItem
  
    PainelControlo.ListTratamentos.ListItems.Clear
    
    SQL = " SELECT * " _
        & " FROM Tratamentos" _
        & " ORDER BY nome_tratamento "
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListTratamentos.ListItems.Add(, , rst("cod_tratamento"))
        itemx.SubItems(1) = rst("nome_tratamento")
        itemx.SubItems(2) = rst("preco_tratamento")
        itemx.Tag = rst("cod_Tratamento")
        
        rst.MoveNext
    Wend
    
End Sub

Private Sub ListMedicamentos_Ordena_SetUp()

    Dim itemx As ListItem
  
    PainelControlo.ListMedicamentos.ListItems.Clear
    
    SQL = " SELECT * " _
        & " FROM Medicamentos" _
        & " ORDER BY nome "
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListMedicamentos.ListItems.Add(, , rst("cod_Medicamento"))
        itemx.SubItems(1) = rst("nome")
        itemx.SubItems(2) = rst("preco_Medicamento")
        itemx.Tag = rst("cod_Medicamento")
        
        rst.MoveNext
    Wend
    
End Sub

