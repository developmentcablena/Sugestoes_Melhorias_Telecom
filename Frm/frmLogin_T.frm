VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Validar Usuário"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   15
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblSenha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
   Begin VB.Image imgSeguranca 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2280
      Picture         =   "frmLogin_T.frx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private blnAlterarSenha As Boolean
Private Declare Function ADsOpenObject Lib "activeds.dll" ( _
    ByVal lpszPath As String, _
    ByVal lpszUserName As String, _
    ByVal lpszPassword As String, _
    ByVal dwReserved As Long, _
    ByVal riid As Any, _
    ByRef ppObject As Any _
) As Long

Private Sub cmdCancelar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdValidar_Click()
vUsuario = Me.txtUsuario

Call GetUserFullName(vUsuario)
Call ValorDescricaoAD(vUsuario, sFullName)

PermissaoAD = False
 On Error GoTo ErrHandler
    Dim ldapObject As Object ' Usar Object para evitar problemas com interface
    Dim ldapConnection As Object ' Usar Object para a conexão LDAP
    Dim userDN As String
    Dim password As String
    Dim domain As String
    Dim adsPath As String
    Dim userObject As Object

    ' Pegando valores de login e senha
    domain = "LDAP://telsrv005/cn=Users,dc=cablenabr,dc=local" ' Ajuste para o seu domínio AD
    userDN = "CABLENABR\" & vUsuario ' Ajuste conforme o formato necessário
    password = txtSenha.Text
    
    ' Conectar ao AD usando as credenciais fornecidas
    Set ldapObject = GetObject("LDAP:")
    Set ldapConnection = ldapObject.OpenDSObject(domain, userDN, password, 0)
    
    
    If Not ldapConnection Is Nothing Then
        
        PermissaoAD = True
        Call Unload(Me)
        Call frmPrincipal.Show
        Exit Sub
    End If

ErrHandler:
    ' Em caso de erro, capturar o erro e mostrar mensagem ao usuário
    'MsgBox "Falha no login. Codigo de erro: " & Err.Number & ". Verifique suas credenciais.", vbExclamation, "Erro"
    If Len(Trim(Me.txtUsuario)) = 1 Then
        MsgBox "Digite o usuário!", vbOKOnly + vbExclamation, "Suporte Simac"
        Me.txtUsuario.SetFocus
        Exit Sub
    End If
    If fnValidarUsuario(Trim(Me.txtUsuario), CodDec(Trim(txtSenha.Text))) = True Then
        Call Unload(Me)
        Call frmPrincipal.Show
    ElseIf blnAlterarSenha = True Then
        Call Unload(Me)
        Call frmAlterarSenha.Show
    Else
        MsgBox "Usuário ou senha inválida!", vbOKOnly + vbCritical, "Suporte Simac"
        txtUsuario.Text = ""
        txtSenha.Text = ""
        Me.txtUsuario.SetFocus
        Exit Sub
    End If
End Sub
Private Sub ValorDescricaoAD(ByVal vUsuario As String, ByVal vNome As String)
Call GetUserDN(vUsuario)

Dim objUser As Object
Dim userObject As Object
Dim strUserName As String
Dim p As String
'p = "CN=Kleber Alves da Silva,OU=Users,OU=Producao,OU=Empresas,DC=cablenabr,DC=local"
' Definir o User Logon Name (UPN)
strUserName = "" & vUsuario & "@cablenabr.local" ' Exemplo de nome de login (UPN)

On Error Resume Next ' Ignorar erros

' Conectar ao AD usando GetObject
'MsgBox "DN teste : " & VCompletoDN  '********** para depurar validar o caminho do AD do usuario

 Set objUser = GetObject("LDAP://telsrv005/" & VCompletoDN)

' Verificar se o objeto foi encontrado
If Err.Number = 0 Then
    strDescription = objUser.Get("description")
    
    ' 'Exibir a descrição se encontrada
   ' If Len(strDescription) > 0 Then
    '    MsgBox "Descrição do usuário: " & strDescription
    'Else
    '    MsgBox "Descrição não encontrada."
    'End If
Else
   ' MsgBox "Erro ao conectar ao AD ou usuário não encontrado."
End If

On Error GoTo 0 ' Retorna ao tratamento padrão de erro

End Sub

Public Function GetUserDN(ByVal sUserName As String) As String
    On Error GoTo ErrorHandler

    ' Conexão com o LDAP
    Dim objConnection As Object
    Dim objCommand As Object
    Dim objRecordset As Object
    Dim ldapPath As String
    Dim sDomain As String
    Dim filter As String
    
    ' Defina o caminho base do domínio (ajuste conforme necessário)
   sDomain = "LDAP://telsrv005.cablenabr.local/DC=cablenabr,DC=local" ' Especifica o servidor explicitamente
   ldapPath = sDomain
    
    ' Criar objetos ADO
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    ' Configurar conexão LDAP
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    ' Configurar o comando de busca
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "<" & ldapPath & ">;(sAMAccountName=" & sUserName & ");distinguishedName;subtree"
    
    ' Executar a busca
    Set objRecordset = objCommand.Execute
    
    ' Verificar se encontrou o usuário
    If Not objRecordset.EOF Then
        VCompletoDN = objRecordset.Fields("distinguishedName").Value
        
    Else
    '    GetUserDN = "Usuário não encontrado"
    End If
    
    ' Fechar conexões
    objRecordset.Close
    objConnection.Close
    
    Set objRecordset = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Erro ao buscar o DN: " & Err.Description, vbCritical
End Function


Private Sub GetUserFullName(ByVal sUserName As String)
    Dim oRootDSE As Object
    Dim oConnection As Object
    Dim oCommand As Object
    Dim oRecordSet As Object
    Dim sDomain As String
    Dim sQuery As String
   
    ' Conectar ao AD
    Set oRootDSE = GetObject("LDAP://RootDSE")
    sDomain = "LDAP://telsrv005.cablenabr.local/DC=cablenabr,DC=local"   'oRootDSE.Get("defaultNamingContext")
    
    ' Configurar conexão e comando para consulta LDAP
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Provider = "ADsDSOObject"
    oConnection.Open "Active Directory Provider"
    
    Set oCommand = CreateObject("ADODB.Command")
    oCommand.ActiveConnection = oConnection
    
    ' Criar a consulta LDAP para buscar o usuário pelo sAMAccountName
    ' Se necessário, você pode alterar a parte do domínio ou a OU
    sQuery = "<" & sDomain & ">;(sAMAccountName=" & sUserName & ");cn;subtree"  '"<LDAP://" & sDomain & ">;(sAMAccountName=" & sUserName & ");cn;subtree"
    
    oCommand.CommandText = sQuery
    
    ' Executar a consulta
    Set oRecordSet = oCommand.Execute
    
    ' Verificar se encontrou resultados
    If Not oRecordSet.EOF Then
        sFullName = oRecordSet.Fields("cn").Value ' Obter o nome completo

    '    MsgBox "Nome completo do usuário: " & sFullName
    Else
    '    MsgBox "Usuário não encontrado no AD."
    End If
    
    ' Fechar conexão
    oRecordSet.Close
    oConnection.Close
End Sub



Private Sub Form_Load()
    Call ConectarBD
End Sub

Private Sub Form_Unload(cancel As Integer)
    cn.Close
    Set cn = Nothing
End Sub

Private Function fnValidarUsuario(ByVal vUsuario As String, ByVal vSenha As String) As Boolean
    fnValidarUsuario = False
    gintUsuarioID = 0
    gstrNome = ""
    gstrDepto = ""
    gstrEMail = ""
    gstrSenha = ""
    blnAlterarSenha = False
    gblnSupervisor = False
    
    strSQL = "SELECT * FROM vw_Usuarios WHERE Usuario = '" & vUsuario & "' AND Senha = '" & vSenha & "' AND Inativo = 0"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF = False Then
        gintUsuarioID = rs!usuarioID
        gstrNome = rs!Nome
        gstrDepto = rs!Departamento
        gstrEMail = rs!Email
        gstrSenha = rs!Senha
      
        
        If rs!AlterarSenha = True Then
            blnAlterarSenha = True
        Else
            fnValidarUsuario = True
        End If
    End If
    
  
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdValidar_Click
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtUsuario.Text)) > 0 Then
            txtSenha.SetFocus
        End If
    End If
End Sub
