VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VendasMedicamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda de Medicamentos"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   ControlBox      =   0   'False
   Icon            =   "VendasMedicamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   5880
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover Medicamento"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Factura"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox cmbMedicamentos 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   17
      Text            =   "Selecione o Medicamento"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.ComboBox cmbPacientes 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Text            =   "Selecione o Paciente"
      Top             =   720
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "VendasMedicamentos.frx":521A
      Top             =   4080
   End
   Begin MSComctlLib.ListView ListVenderMedicamentos 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   2160
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "0 €"
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Total a Pagar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Nome do Medicamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Código do Paciente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do Paciente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Código da Factura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "VendasMedicamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim ordem As Integer

Private Sub cmbMedicamentos_Click()

    Dim itemx As ListItem
    
    SQL = " SELECT * FROM medicamentos" _
        & " WHERE nome like '" & cmbMedicamentos.Text & "'"
    Set rst = db.OpenRecordset(SQL)
    
    If Val(rst("emblagem_disponiveis")) > 0 Then
        If Val(Text1.Text) <= Val(rst("emblagem_disponiveis")) Then
            ordem = ordem + 1
    
            SQL = " SELECT * FROM medicamentos" _
                & " WHERE nome like '" & cmbMedicamentos.Text & "'"
            Set rst = db.OpenRecordset(SQL)

            While Not rst.EOF
                Set itemx = ListVenderMedicamentos.ListItems.Add(, , ordem)
                itemx.SubItems(1) = rst("cod_medicamento")
                itemx.SubItems(2) = rst("nome")
                itemx.SubItems(3) = rst("Preco_medicamento")
                itemx.SubItems(4) = Text1.Text
                itemx.SubItems(5) = (CCur(Text1.Text) * CCur(rst("Preco_medicamento")))
                itemx.Tag = ordem
                Label13.Caption = CCur(Label13.Caption) + ((CCur(Text1.Text) * CCur(rst("Preco_medicamento")))) & " €"
                rst.MoveNext
            Wend
    
            SQL = "INSERT INTO Faturas(cod_fatura,cod_paciente,nome_paciente,quantidade,ordem,nome_medicamento,cod_medicamento,data,hora,preco_unidade,preco_total)" _
                & "VALUES('" & Label8.Caption & "'" _
                & ", '" & Label9.Caption & "'" _
                & ", '" & cmbPacientes.Text & "'" _
                & ", '" & Text1.Text & "'" _
                & ", '" & itemx & "'" _
                & ", '" & cmbMedicamentos.Text & "'" _
                & ", '" & itemx.SubItems(1) & "'" _
                & ", '" & Label10.Caption & "'" _
                & ", '" & Label11.Caption & "'" _
                & ", '" & itemx.SubItems(3) & "'" _
                & ", '" & itemx.SubItems(5) & "')"
            db.Execute SQL
            
            SQL1 = "INSERT INTO ImprimirFaturas(cod_fatura,cod_paciente,nome_paciente,quantidade,ordem,nome_medicamento,cod_medicamento,data,hora,preco_unidade,preco_total)" _
                & "VALUES('" & Label8.Caption & "'" _
                & ", '" & Label9.Caption & "'" _
                & ", '" & cmbPacientes.Text & "'" _
                & ", '" & Text1.Text & "'" _
                & ", '" & itemx & "'" _
                & ", '" & cmbMedicamentos.Text & "'" _
                & ", '" & itemx.SubItems(1) & "'" _
                & ", '" & Label10.Caption & "'" _
                & ", '" & Label11.Caption & "'" _
                & ", '" & itemx.SubItems(3) & "'" _
                & ", '" & itemx.SubItems(5) & "')"
            db.Execute SQL1
    
            SQL = " SELECT * FROM medicamentos" _
                & " WHERE nome like '" & cmbMedicamentos.Text & "'"
            Set rst = db.OpenRecordset(SQL)
            SQL = "UPDATE Medicamentos SET emblagem_disponiveis = '" & Val(rst("emblagem_disponiveis")) - Val(Text1.Text) & "'" _
                & " WHERE cod_medicamento = " & itemx.SubItems(1)
            db.Execute SQL
    
            Text1.Text = ""
            cmbMedicamentos.Enabled = False
            Text1.SetFocus
            SQL = "select * from imprimirfaturas"
            Set rst = db.OpenRecordset(SQL)
            If Not (rst.BOF = True And rst.EOF = True) Then
                cmdRemover.Enabled = True
            End If
            Exit Sub
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Caption = "Quantidade Elevada"
            FormMsgBoxNormal.Label1.Caption = "A quantidade que introduziu é maior que  stock existente!" & vbCr & "Introduza uma quantidade mais baixa."
            VendasMedicamentos.Enabled = False
            Exit Sub
        End If
    Else
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Caption = "Stock de Medicamentos"
        FormMsgBoxNormal.Label1.Caption = "Já não tem embalagens disponiveis do medicamento que escolheu!"
        VendasMedicamentos.Enabled = False
        Exit Sub
    End If
    
End Sub

Private Sub cmdImprimir_Click()
            
    Unload DataEnvironment1
    Load DataEnvironment1
    rptImprimirFaturas.Show
    
End Sub

Private Sub cmdRemover_Click()

    SQL2 = " SELECT * FROM imprimirfaturas" _
        & " WHERE ordem = " & ListVenderMedicamentos.SelectedItem.Tag
    Set rst = db.OpenRecordset(SQL2)
    SQL3 = "DELETE FROM imprimirfaturas WHERE ordem = " & ListVenderMedicamentos.SelectedItem.Tag
        db.Execute SQL3

    SQL = " SELECT * FROM faturas" _
        & " WHERE ordem = " & ListVenderMedicamentos.SelectedItem.Tag
    Set rst = db.OpenRecordset(SQL)
    SQL = "DELETE FROM faturas WHERE ordem = " & ListVenderMedicamentos.SelectedItem.Tag
        db.Execute SQL
    ListVenderMedicamentos.ListItems.Remove ListVenderMedicamentos.SelectedItem.Index
    
    SQL = "select * from imprimirfaturas"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        cmdRemover.Enabled = False
        Label13.Caption = "0 €"
    End If
    
End Sub

Private Sub cmdSair_Click()

    SQL = " SELECT * FROM imprimirfaturas"
    Set rst = db.OpenRecordset(SQL)
    SQL = "DELETE FROM imprimirfaturas"
        db.Execute SQL
    PainelControlo.Enabled = True
    PainelControlo.SetFocus
    PainelControlo.ListMedicamentos.SetFocus
    Unload Me
    
End Sub

Private Sub Form_Load()

    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    ListVenderMedicamentos_SetUp
    
    SQL = " SELECT * FROM faturas"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        Label8.Caption = 1
    Else
        With rst
            If Not .EOF Then
                Do While Not .EOF
                    x = rst("cod_fatura")
                    .MoveNext
                Loop
            Else
                y = rst("cod_fatura")
            End If
        End With
        res = x
        Label8.Caption = res + 1
    End If

    SQL = " SELECT * FROM Pacientes ORDER BY nome_paciente"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not rst.EOF Then
            Do While Not rst.EOF
                cmbPacientes.AddItem rst("nome_paciente")
                rst.MoveNext
            Loop
        Else
            cmbPacientes.Text = ""
        End If
    End With
    
    SQL = " SELECT * FROM Medicamentos ORDER BY nome"
    Set rst = db.OpenRecordset(SQL)
    With rst
        If Not .EOF Then
            Do While Not .EOF
                cmbMedicamentos.AddItem rst("nome")
                .MoveNext
            Loop
        Else
            cmbMedicamentos.Text = ""
        End If
    End With
    
    Label10.Caption = PainelControlo.Label18
    Label11.Caption = PainelControlo.Label19
    
    ordem = 0
    
    SQL = "select * from imprimirfaturas"
    Set rst = db.OpenRecordset(SQL)
    If rst.BOF = True And rst.EOF = True Then
        cmdRemover.Enabled = False
    End If

End Sub

Private Sub cmbPacientes_Click()

    pacientes = cmbPacientes.Text
    SQL = " SELECT * FROM pacientes" _
        & " WHERE nome_paciente LIKE " & "'" & pacientes & "'"
        
    Set rst = db.OpenRecordset(SQL)
    
    Label9.Caption = rst("cod_paciente")
    
    cmbPacientes.Enabled = False
    Text1.Enabled = True
    Text1.SetFocus

End Sub

Private Sub ListVenderMedicamentos_SetUp()
     
    ListVenderMedicamentos.ListItems.Clear
    ListVenderMedicamentos.ColumnHeaders.Add , , "Ordem", ListVenderMedicamentos.Width - 9050
    ListVenderMedicamentos.ColumnHeaders.Add , , "Código do Medicamento", ListVenderMedicamentos.Width - 7800
    ListVenderMedicamentos.ColumnHeaders.Add , , "Nome do Medicamento", ListVenderMedicamentos.Width - 6100
    ListVenderMedicamentos.ColumnHeaders.Add , , "Preço Unidade", ListVenderMedicamentos.Width - 8400
    ListVenderMedicamentos.ColumnHeaders.Add , , "Quantidade", ListVenderMedicamentos.Width - 8700
    ListVenderMedicamentos.ColumnHeaders.Add , , "Preço Total", ListVenderMedicamentos.Width - 8700
    
    ListVenderMedicamentos.View = lvwReport
 
End Sub

Private Sub Text1_Change()
    
    cmbMedicamentos.Enabled = True
    
End Sub
