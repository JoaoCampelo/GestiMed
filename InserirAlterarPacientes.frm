VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form InserirAlterarPacientes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8310
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "InserirAlterarPacientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdInserir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Height          =   1005
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   6240
      Width           =   4095
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   4440
      Width           =   4215
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "InserirAlterarPacientes.frx":521A
      Top             =   7560
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label14 
      Caption         =   "Doenças:"
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
      Left            =   240
      TabIndex        =   29
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Estado Civil:"
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
      Left            =   2520
      TabIndex        =   28
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Sexo:"
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
      Left            =   240
      TabIndex        =   27
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Código Postal:"
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
      Left            =   240
      TabIndex        =   26
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Cidade:"
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
      Left            =   2880
      TabIndex        =   25
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Morada:"
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
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Número de Contribuinte:"
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
      Left            =   240
      TabIndex        =   23
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Bilhete de Identidade:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "E - Mail:"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Telemóvel:"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Telefone:"
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
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Data de Nascimento:"
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
      Left            =   240
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "InserirAlterarPacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String
Dim SQL1 As String

Private Sub cmdInserir_Click()
   
    If Len(Trim(Text1.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Paciente» é obrigatório."
        FormMsgBoxNormal.Caption = "Pacientes (Nome)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    contnome = Len(Text1.Text)
    For i = 1 To contnome
        If IsNumeric(Mid(Text1.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Paciente» não pode conter números!"
            FormMsgBoxNormal.Caption = "Pacientes (Nome)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    Next i

    If Not IsDate(Text3.Text) Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Data de Nascimento» tem que ser uma data (dd/mm/aa)!"
        FormMsgBoxNormal.Caption = "Pacientes (Data)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
 
    If Text4.Text <> "" Then
        If Len(Text4.Text) = 9 Then
            If Not IsNumeric(Text4.Text) Then
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "O campo «Telefone» tem que ser um número!"
                FormMsgBoxNormal.Caption = "Pacientes (Telefone)"
                InserirAlterarPacientes.Enabled = False
                Exit Sub
            End If
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Telefone» tem que ter nove números!"
            FormMsgBoxNormal.Caption = "Pacientes (Telefone)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    End If
    
    If Text5.Text <> "" Then
        If Len(Text5.Text) = 9 Then
            If Not IsNumeric(Text5.Text) Then
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "O campo «Telemóvel» tem que ser um número!"
                FormMsgBoxNormal.Caption = "Pacientes (Telemóvel)"
                InserirAlterarPacientes.Enabled = False
                Exit Sub
            End If
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Telemóvel» tem que ter nove números!"
            FormMsgBoxNormal.Caption = "Pacientes (Telemóvel)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    End If
    
    If Text6.Text <> "" Then
        Dim VALID As Boolean
        VALID = IsValidEmail(Text6.Text)
        If VALID = False Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O e-mail introduzido é inválido!"
            FormMsgBoxNormal.Caption = "Pacientes (E-Mail)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    End If
   
    If Text7.Text <> "" Then
        If Len(Text7.Text) <= 8 Then
            If Not IsNumeric(Text7.Text) Then
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "O campo «Bilhete de Identidade» tem que ser um número!"
                FormMsgBoxNormal.Caption = "Pacientes (Bilhete de Identidade)"
                InserirAlterarPacientes.Enabled = False
                Exit Sub
            End If
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Bilhete de Identidade» não pode ter mais de oito números!"
            FormMsgBoxNormal.Caption = "Pacientes (Bilhete de Identidade)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    End If
    
    If Text8.Text <> "" Then
        If Not Len(Text8.Text) <> 9 Then
            If Not IsNumeric(Text8.Text) Then
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "O campo «Número de Contribuinte» tem que ser um número!"
                FormMsgBoxNormal.Caption = "Pacientes (Número de Contribuinte)"
                InserirAlterarPacientes.Enabled = False
                Exit Sub
            End If
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Número de Contribuinte» tem que ter nove números!"
            FormMsgBoxNormal.Caption = "Pacientes (Número de Contribuinte)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    End If
    
    If Len(Trim(Text9.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Morada» é obrigatório."
        FormMsgBoxNormal.Caption = "Pacientes (Morada)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    
    cont = Len(Text10.Text)
    For i = 1 To cont
        If i <> 5 Then
            If Not IsNumeric(Mid(Text10.Text, i, 1)) Then
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "O campo «Código Postal» tem que ser um código postal (NNNN-NNN) !"
                FormMsgBoxNormal.Caption = "Pacientes (Código Postal)"
                InserirAlterarPacientes.Enabled = False
                Exit Sub
            End If
        Else
            If i = 5 Then
                If Mid(Text10.Text, 5, 1) <> "-" Then
                    FormMsgBoxNormal.Show
                    FormMsgBoxNormal.Label1.Caption = "O campo «Código Postal» tem que ser um código postal (NNNN-NNN) !"
                    FormMsgBoxNormal.Caption = "Pacientes (Código Postal)"
                    InserirAlterarPacientes.Enabled = False
                    Exit Sub
                End If
            End If
        End If
    Next i
    If Len(Text10.Text) <> 8 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Código Postal» tem que ser um código postal (NNNN-NNN) !"
        FormMsgBoxNormal.Caption = "Pacientes (Código Postal)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    
    If Len(Trim(Text11.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Cidade» é obrigatório."
        FormMsgBoxNormal.Caption = "Pacientes (Cidade)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    contcidade = Len(Text11.Text)
    For i = 1 To contcidade
        If IsNumeric(Mid(Text11.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Cidade» não pode conter números!"
            FormMsgBoxNormal.Caption = "Pacientes (Cidade)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    Next i
    
    If Len(Trim(Text12.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Sexo» é obrigatório."
        FormMsgBoxNormal.Caption = "Pacientes (Sexo)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    contsexo = Len(Text12.Text)
    For i = 1 To contsexo
        If IsNumeric(Mid(Text12.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Sexo» não pode conter números!"
            FormMsgBoxNormal.Caption = "Pacientes (Sexo)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    Next i
    If Text12.Text <> "Masculino" And Text12.Text <> "Feminino" Then
        FormMsgBoxNormal.Height = 2150
        FormMsgBoxNormal.Width = 5750
        FormMsgBoxNormal.Label1.Height = 1000
        FormMsgBoxNormal.Label1.Width = 5055
        FormMsgBoxNormal.cmdOK.Top = 1200
        FormMsgBoxNormal.cmdOK.Left = 2100
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Sexo» tem de obedecer a uma das seguintes condições:" & vbCr & "" & vbCr & "Masculino" & vbCr & "Feminino"
        FormMsgBoxNormal.Caption = "Pacientes (Sexo)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
        
    If Len(Trim(Text13.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Estado Civil» é obrigatório."
        FormMsgBoxNormal.Caption = "Pacientes (Estado Civil)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    contestciv = Len(Text13.Text)
    For i = 1 To contestciv
        If IsNumeric(Mid(Text13.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Estado Civil» não pode conter números!"
            FormMsgBoxNormal.Caption = "Pacientes (Estado Civil)"
            InserirAlterarPacientes.Enabled = False
            Exit Sub
        End If
    Next i
    If (Text13.Text <> "Solteiro" And Text13.Text <> "Solteira" And Text13.Text <> "Casado" And Text13.Text <> "Casada" And Text13.Text <> "Separado" And Text13.Text <> "Separada" And Text13.Text <> "Divorciado" And Text13.Text <> "Divorciada" And Text13.Text <> "Viúvo" And Text13.Text <> "Viúva") Then
        FormMsgBoxNormal.Height = 2800
        FormMsgBoxNormal.Width = 5750
        FormMsgBoxNormal.Label1.Height = 1400
        FormMsgBoxNormal.Label1.Width = 5055
        FormMsgBoxNormal.cmdOK.Top = 1850
        FormMsgBoxNormal.cmdOK.Left = 2100
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Estado Civil» tem de obedecer a um dos seguintes estados:" & vbCr & "" & vbCr & "Solteiro(a)" & vbCr & "Casado(a)" & vbCr & "Separado(a)" & vbCr & "Divorciado(a)" & vbCr & "Viúvo(a)"
        FormMsgBoxNormal.Caption = "Pacientes (Estado Civil)"
        InserirAlterarPacientes.Enabled = False
        Exit Sub
    End If
    
    If cmdInserir.Caption = "&Inserir" Then
        SQL = "INSERT INTO Pacientes(nome_paciente, cod_paciente, data_nascimento, telefone, telemovel, email, bilhete_identidade, num_contribuinte, morada, cod_postal, cidade, sexo, estado_civil, doencas) " _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text4.Text & "'" _
                        & ", '" & Text5.Text & "'" _
                        & ", '" & Text6.Text & "'" _
                        & ", '" & Text7.Text & "'" _
                        & ", '" & Text8.Text & "'" _
                        & ", '" & Text9.Text & "'" _
                        & ", '" & Text10.Text & "'" _
                        & ", '" & Text11.Text & "'" _
                        & ", '" & Text12.Text & "'" _
                        & ", '" & Text13.Text & "'" _
                        & ", '" & Text14.Text & "')"
        db.Execute SQL
        
        SQL1 = "INSERT INTO seg_Pacientes(nome_paciente, cod_paciente, data_nascimento, telefone, telemovel, email, bilhete_identidade, num_contribuinte, morada, cod_postal, cidade, sexo, estado_civil, doencas) " _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text4.Text & "'" _
                        & ", '" & Text5.Text & "'" _
                        & ", '" & Text6.Text & "'" _
                        & ", '" & Text7.Text & "'" _
                        & ", '" & Text8.Text & "'" _
                        & ", '" & Text9.Text & "'" _
                        & ", '" & Text10.Text & "'" _
                        & ", '" & Text11.Text & "'" _
                        & ", '" & Text12.Text & "'" _
                        & ", '" & Text13.Text & "'" _
                        & ", '" & Text14.Text & "')"
        db.Execute SQL1
        
        FormMsgBoxSimNao.Show
        FormMsgBoxSimNao.Label1.Caption = "O paciente foi inserido com exito!" & vbCr & "Deseja inserir outro paciente?"
        FormMsgBoxSimNao.Caption = "Inserir (Pacientes)"
        InserirAlterarPacientes.Enabled = False
        
    Else
        SQL = "UPDATE Pacientes SET telefone = '" & Text4.Text & "'" _
            & ", telemovel = '" & Text5.Text & "'" _
            & ", email = '" & Text6.Text & "'" _
            & ", morada = '" & Text9.Text & "'" _
            & ", cod_postal = '" & Text10.Text & "'" _
            & ", cidade = '" & Text11.Text & "'" _
            & ", estado_civil = '" & Text13.Text & "'" _
            & ", doencas = '" & Text14.Text & "'" _
            & " WHERE cod_paciente = " & PainelControlo.ListPacientes.SelectedItem.Tag
        db.Execute SQL
        
        SQL1 = "UPDATE seg_Pacientes SET telefone = '" & Text4.Text & "'" _
            & ", telemovel = '" & Text5.Text & "'" _
            & ", email = '" & Text6.Text & "'" _
            & ", morada = '" & Text9.Text & "'" _
            & ", cod_postal = '" & Text10.Text & "'" _
            & ", cidade = '" & Text11.Text & "'" _
            & ", estado_civil = '" & Text13.Text & "'" _
            & ", doencas = '" & Text14.Text & "'" _
            & " WHERE cod_paciente = " & PainelControlo.ListPacientes.SelectedItem.Tag
        db.Execute SQL1
        
        ListPacientes_Ordena_SetUp
        
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Os dados do paciente foram alterados com êxito!"
        FormMsgBoxNormal.Caption = "Pacientes (Alterar)"
        InserirAlterarPacientes.Enabled = False
    End If
  
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
End Sub

Private Sub cmdSair_Click()

    PainelControlo.Enabled = True
    ListPacientes_Ordena_SetUp
    PainelControlo.SetFocus
    PainelControlo.ListPacientes.SetFocus
    SQL = "select * from pacientes"
    Set rst = db.OpenRecordset(SQL)
    If Not (rst.BOF = True And rst.EOF = True) Then
        PainelControlo.cmdAlterar3.Enabled = True
        PainelControlo.cmdDel3.Enabled = True
    End If
    Unload Me
    
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

Public Function IsValidEmail(email As String) As Boolean

Dim myAt As Integer
Dim myAtLastPos As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim myDotAt As Integer
Dim myAtDot As Integer
Dim mySpace As Integer
IsValidEmail = True
mySpace = InStr(1, email, " ", vbTextCompare)
myAtLastPos = InStrRev(email, "@", , vbTextCompare)
myAt = InStr(1, email, "@", vbTextCompare)
myAtDot = InStr(1, email, "@.", vbTextCompare)
myDotAt = InStr(1, email, ".@", vbTextCompare)
myDot = InStr(myAt + 2, email, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
If myAtDot > 0 Or myDotAt > 0 Or myAtLastPos <> myAt Or mySpace > 0 Or myAt = 0 Or myDot = 0 Or myDotDot > 0 Or Right(email, 1) = "." Then IsValidEmail = False

Carac = Array("!", "#", "$", "%", "&", "*", "(", ")", "+", "=", "/", "\", "|", "?", "'", """", "{", "}", "[", "]", "ª", "º", ":", ",", ";", "§", "°", "<", ">")
For intVer = LBound(Carac) To UBound(Carac)
  If InStr(email, Carac(intVer)) > 0 Then
    IsValidEmail = False
   Exit Function
  End If
Next

End Function
