VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Simac - Usuários"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsuarios_T.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8685
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmdTipoU 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwUsuarios 
      Height          =   5655
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9975
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame fraDadosUsuario 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   4980
      Begin VB.CheckBox chkEmBranco 
         Appearance      =   0  'Flat
         Caption         =   "Em branco"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkInativo 
         Appearance      =   0  'Flat
         Caption         =   "Inativo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   3480
         Width           =   975
      End
      Begin VB.CheckBox chkAlterar 
         Appearance      =   0  'Flat
         Caption         =   "Alterar Senha"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "&Salvar"
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   1095
         TabIndex        =   14
         Top             =   4080
         Width           =   945
      End
      Begin VB.TextBox txtConfirmar 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   3000
         Width           =   2160
      End
      Begin VB.TextBox txtEMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   4695
      End
      Begin VB.TextBox txtDepto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   945
         Width           =   4695
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Uusuário:"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblConfirmar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Confirmar Senha"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   2760
         Width           =   1470
      End
      Begin VB.Label lblSenha 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label lblEMail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblDepto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Depto."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblUsuario 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Usuario / CB"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblNome 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   135
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private strSQL As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Tempo()
'Sleep 3000
End Sub

'Private Sub cmdAdicionar_Click()
'Dim i As Integer
'    If Len(Trim(cboModulos.Text)) = 0 Then
'        MsgBox "Selecione um módulo!", vbOKOnly + vbExclamation, "Suporte Simac"
'        cboModulos.SetFocus
'        Exit Sub
'    End If
'
'    For i = 1 To Me.lvwUsuarios.ListItems.Count
'        If Me.lvwUsuarios.ListItems(i).Selected = True Then
             
'            If cmdNovo.Enabled = True Then
'                Call suAdicionarModulo(lvwUsuarios.ListItems(i).Text, cboModulos.Text)                                     'gintNovoUsuarioID, cboModulos.Text)
'                Call suMostrarPermissoes(lvwUsuarios.ListItems(i).Text)
'            ElseIf cmdEditar.Enabled = True Then
'                Call suAdicionarModulo(lvwUsuarios.ListItems(i).Text, cboModulos.Text)
'                Call suMostrarPermissoes(lvwUsuarios.ListItems(i).Text)
'            End If
'        End If
'    Next
        
'End Sub

Private Sub cmdCancelar_Click()
    Call suLimparDados
 
    Call suHabilitarDados(False)
    cmdNovo.Enabled = True
    cmdEditar.Enabled = True
    If cmdSalvar.Caption = "&Atualizar" Then
        cmdSalvar.Caption = "&Salvar"
    End If
End Sub

Private Sub cmdEditar_Click()
    If Len(Trim(txtNome.Text)) > 0 And Len(Trim(txtUsuario.Text)) > 0 Then
        Call suHabilitarDados(True)
        cmdNovo.Enabled = False
        cmdSalvar.Caption = "&Atualizar"
    Else
        MsgBox "Nenhum usuário foi selecionado!", vbOKOnly + vbExclamation, "Suporte Simac"
        cmdNovo.Enabled = True
        cmdSalvar.Caption = "&Salvar"
    End If
End Sub

Private Sub cmdFechar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdNovo_Click()
    Call suLimparDados
    Call suHabilitarDados(True)
    cmdEditar.Enabled = False
    txtNome.SetFocus
End Sub



Private Sub cmdSalvar_Click()
Dim i As Integer

    If Len(Trim(txtNome.Text)) = 0 Then
        MsgBox "Digite o nome do usuário!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtNome.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtUsuario.Text)) = 0 Then
        MsgBox "Digite o login do usuário!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtUsuario.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtDepto.Text)) = 0 Then
        MsgBox "Digite o departamento do usuário!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtDepto.SetFocus
        Exit Sub
    End If
    
    If chkEmBranco.Value = 0 Then
        If Len(Trim(txtSenha.Text)) = 0 Then
            MsgBox "Digite uma senha!", vbOKOnly + vbExclamation, "Suporte Simac"
            txtSenha.SetFocus
            Exit Sub
        End If
        
        If Len(Trim(txtConfirmar.Text)) = 0 Then
            MsgBox "Digite a confirmação da senha!", vbOKOnly + vbExclamation, "Suporte Simac"
            txtConfirmar.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(Me.cmdTipoU.Text) = "" Then
        MsgBox "Por favor, selecione o tipo de usuário!", vbExclamation, "Suporte Simac"
        Me.cmdTipoU.SetFocus
        Exit Sub
    End If
     
    If Trim(txtSenha.Text) <> Trim(txtConfirmar.Text) Then
        MsgBox "Senhas não conferem!", vbOKOnly + vbExclamation, "Suporte Simac"
        txtSenha.Text = ""
        txtConfirmar.Text = ""
        txtSenha.SetFocus
        Exit Sub
    Else
        If cmdSalvar.Caption = "&Salvar" Then
            If fnCadastrarUsuario(Trim(txtNome.Text), Trim(txtUsuario.Text), CodDec(Trim(txtSenha.Text)), Trim(txtDepto.Text), Trim(txtEMail.Text), chkInativo, chkAlterar, Me.cmdTipoU) = True Then
                MsgBox "Usuário cadastrado com sucesso!", vbOKOnly + vbInformation, "Suporte Simac"
                Call suListarUsuarios
                Call suHabilitarDados(False)
                chkEmBranco.Value = 0
                Me.cmdTipoU.Clear
              
            End If
        Else
            For i = 1 To Me.lvwUsuarios.ListItems.Count
                If Me.lvwUsuarios.ListItems(i).Selected = True Then
                    If fnAtualizar(lvwUsuarios.ListItems(i).Text, Trim(txtDepto.Text), Trim(txtEMail.Text), CodDec(Trim(txtSenha.Text)), chkInativo, chkAlterar, Trim(Me.txtNome.Text), Trim(Me.txtUsuario.Text), Me.cmdTipoU) = True Then
                    End If
                End If
            Next
              
                MsgBox "Atualizações efetuadas com sucesso!", vbOKOnly + vbInformation, "Suporte Simac"
                Call suListarUsuarios
                Call suHabilitarDados(False)
                chkEmBranco.Value = 0
                Me.cmdTipoU.Clear
                End If
            End If
End Sub

Private Sub Form_Load()
Call Combobox
Me.txtEMail.Enabled = False


Me.lvwUsuarios.LabelEdit = lvwManual
 With Me.lvwUsuarios
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "Usuarios", 3500
        .LabelEdit = lvwManual
End With

Call suListarUsuarios
Call suHabilitarDados(False)
End Sub
Private Sub Combobox()
Me.cmdTipoU.AddItem "Usuário"
Me.cmdTipoU.AddItem "Administrador"
End Sub

Private Sub lvwUsuarios_Click()
Call Combobox
Dim i As Integer

    If lvwUsuarios.ListItems.Count > 0 Then
        For i = 1 To lvwUsuarios.ListItems.Count
            If lvwUsuarios.ListItems(i).Selected = True Then
                Call suLimparDados
                Call suHabilitarDados(False)
                Call suMostrarDados(lvwUsuarios.ListItems(i).Text)
                Me.cmdTipoU.Enabled = False
                cmdNovo.Enabled = True
                cmdEditar.Enabled = True
                cmdSalvar.Caption = "&Salvar"
            End If
        Next
    End If
    
End Sub



Private Sub suRemoverModulo(ByVal vPermissaoID As Integer, ByVal vUsuarioID As Integer)
Call ConectarBD

    strSQL = "DELETE FROM tb_Permissoes WHERE PermissaoID = " & vPermissaoID & " AND UsuarioID = " & vUsuarioID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
    Set cn = Nothing
End Sub

Private Sub suAdicionarModulo(ByVal vUsuarioID As Integer, ByVal vModulo As String)
Call ConectarBD

  

    strSQL = "INSERT INTO tb_Permissoes (Modulo,Permissao,UsuarioID) " & _
             "VALUES ('" & vModulo & "',0," & vUsuarioID & ")"
             
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Set rs = Nothing
    Set cn = Nothing
End Sub



Private Function fnAtualizarPermissao(ByVal vPermissao As Integer, ByVal vUsuarioID As Integer) As Boolean
Call ConectarBD
On Error GoTo Erro

MsgBox "permissao " & vPermissao
MsgBox "ususario " & vUsuarioID

    fnAtualizarPermissao = False
    
    strSQL = "UPDATE tb_Permissoes SET " & _
    " Permissao = '" & vPermissao & "' " & _
    " WHERE PermissaoID = " & vUsuarioID & " "
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    fnAtualizarPermissao = True
    Set rs = Nothing
    Exit Function
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
    
End Function

Private Function fnAtualizar(ByVal vUsuarioID As Integer, ByVal vDepto As String, ByVal vEMail As String, ByVal vSenha As String, ByRef vInativo As CheckBox, ByVal vAlterarSenha As CheckBox, ByVal vNome As String, ByVal vUsuario As String, ByVal vPermissao As String) As Boolean
Call ConectarBD
On Error GoTo Erro
    fnAtualizar = False
    Screen.MousePointer = 11
    Call Tempo
    
    strSQL = "UPDATE tb_Usuario SET " & _
    "Departamento = '" & vDepto & "'," & _
    "Nome = '" & vNome & "', " & _
    "Usuario = '" & vUsuario & "', " & _
    "EMail = '" & vEMail & "'," & _
    "Senha = '" & vSenha & "'," & _
    "Inativo = '" & vInativo & "'," & _
    "Permissao = '" & vPermissao & "', " & _
    "AlterarSenha = '" & vAlterarSenha & "' " & _
    "WHERE UsuarioID = " & vUsuarioID & " "
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Screen.MousePointer = 0
    fnAtualizar = True
    Set rs = Nothing
    Exit Function

Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
End Function

Private Function fnCadastrarUsuario(ByVal vNome As String, ByVal vUsuario As String, ByVal vSenha As String, ByVal vDepto As String, ByVal vEMail As String, ByRef vInativo As CheckBox, ByRef vAlterarSenha As CheckBox, ByVal vTipoUsuario As String) As Boolean
On Error GoTo Erro
Call ConectarBD


    fnCadastrarUsuario = False
    gintNovoUsuarioID = 0
    Screen.MousePointer = 11
    Call Tempo
    
    strSQL = "SELECT * FROM tb_Usuario WHERE Usuario = '" & vUsuario & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = True Then
        rs.AddNew
        
        rs!Nome = vNome & ""
        rs!Usuario = vUsuario & ""
        rs!Senha = vSenha & ""
        rs!Departamento = vDepto & ""
        rs!Email = vEMail & ""
        rs!Inativo = vInativo
        rs!AlterarSenha = vAlterarSenha
        rs!Permissao = vTipoUsuario
        
        rs.Update
        gintNovoUsuarioID = rs!usuarioID
       
        Screen.MousePointer = 0
        fnCadastrarUsuario = True
        Set rs = Nothing
        Exit Function
    End If
    
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbOKOnly + vbCritical, "Suporte Simac"
End Function

Private Sub suLimparDados()
    txtNome.Text = ""
    txtUsuario.Text = ""
    txtDepto.Text = ""
    txtEMail.Text = ""
    txtSenha.Text = ""
    txtConfirmar.Text = ""
    chkAlterar.Value = 0
    chkInativo.Value = 0
    Me.cmdTipoU.Clear
    Call Combobox
  

End Sub



Private Sub suHabilitarDados(ByVal vBoolean As Boolean)
    txtNome.Enabled = vBoolean
    txtUsuario.Enabled = vBoolean
    txtDepto.Enabled = vBoolean
    Me.cmdTipoU.Enabled = vBoolean
    'txtEMail.Enabled = vBoolean
    txtSenha.Enabled = vBoolean
    txtConfirmar.Enabled = vBoolean
    chkAlterar.Enabled = vBoolean
    chkInativo.Enabled = vBoolean
    cmdSalvar.Enabled = vBoolean
    cmdCancelar.Enabled = vBoolean
    chkEmBranco.Enabled = vBoolean
End Sub




Private Sub suMostrarDados(ByVal vUsuarioID As Integer)
Call ConectarBD
    strSQL = "SELECT * FROM vw_Usuarios WHERE UsuarioID = " & vUsuarioID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        txtNome.Text = rs!Nome & ""
        txtUsuario.Text = rs!Usuario & ""
        txtDepto.Text = rs!Departamento & ""
        txtEMail.Text = rs!Email & ""
        Me.chkInativo.Value = IIf(rs!Inativo = False, 0, 1)
        Me.chkAlterar.Value = IIf(rs!AlterarSenha = False, 0, 1)
        Me.cmdTipoU.Text = rs!Permissao
        
    End If
    
    rs.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub

Private Sub suListarUsuarios()
Dim lvwItem As ListItem
Call ConectarBD
    strSQL = "SELECT * FROM vw_Usuarios ORDER BY Nome"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Me.lvwUsuarios.ListItems.Clear
    
    If rs.EOF = False Then
        Do While Not rs.EOF
            Set lvwItem = Me.lvwUsuarios.ListItems.Add(, , rs!usuarioID)
            lvwItem.SubItems(1) = rs!Nome
            
            If rs!Inativo = True Then
                ' Colorir a linha de vermelho
                lvwItem.ForeColor = vbRed
                Dim i As Integer
                 ' Corrige o loop para percorrer os subitens corretamente
                For i = 1 To Me.lvwUsuarios.ColumnHeaders.Count - 1
                    If i <= lvwItem.ListSubItems.Count Then
                        lvwItem.ListSubItems(i).ForeColor = vbRed
                    End If
                Next i
            End If
            rs.MoveNext
        Loop
    Else
        'gexUsuarios.HoldFields
        'Set gexUsuarios.ADORecordset = rs
    End If
    
    Set rs = Nothing
    Set cn = Nothing
End Sub






