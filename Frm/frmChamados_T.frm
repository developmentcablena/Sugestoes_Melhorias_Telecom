VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChamados 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte Simac - Verificar Sugestões"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15180
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChamados_T.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   15180
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwChamados 
      Height          =   8415
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   14843
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
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton cmdLimpar 
      Appearance      =   0  'Flat
      Caption         =   "&Atualizar"
      Height          =   300
      Left            =   5880
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdBaixarOS 
      Appearance      =   0  'Flat
      Caption         =   "&Finalizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAtender 
      Appearance      =   0  'Flat
      Caption         =   "&Atender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton cmdPesquisar 
      Appearance      =   0  'Flat
      Caption         =   "&Pesquisar"
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtOSID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   495
      Width           =   1575
   End
   Begin VB.CommandButton cmdAtendimento 
      Appearance      =   0  'Flat
      Caption         =   "&Atendimento"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   9480
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   375
      Left            =   13680
      TabIndex        =   3
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label lblOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Atendimento"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   255
      Width           =   1545
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmChamados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim i As Long

Private Sub Tempo()
Sleep 5000
End Sub

Private Sub cboStatus_Click()
Select Case Me.cboStatus
    Case Is = "Em Aberto"
        Me.cmdAtender.Enabled = True
        Me.cmdAtendimento.Enabled = False
        Me.cmdBaixarOS.Enabled = False
        Me.lvwChamados.ForeColor = vbBlack
        Me.cmdCancelar.Enabled = True
        Me.cmdCancelar.Visible = True
        gintOSID = 0
    Case Is = "Em Atendimento"
        Me.cmdAtender.Enabled = False
        Me.cmdAtendimento.Enabled = True
        Me.cmdBaixarOS.Enabled = True
        Me.lvwChamados.ForeColor = vbBlack
        Me.cmdCancelar.Enabled = False
        Me.cmdCancelar.Visible = False
        gintOSID = 0

    Case Is = "Finalizada"
        Me.cmdAtender.Enabled = False
        Me.cmdAtendimento.Enabled = False
        Me.cmdBaixarOS.Enabled = False
        Me.lvwChamados.ForeColor = vbBlue
        Me.cmdCancelar.Enabled = False
        Me.cmdCancelar.Visible = False
    Case Is = "Cancelada"
        Me.cmdAtender.Enabled = False
        Me.cmdAtendimento.Enabled = False
        Me.cmdBaixarOS.Enabled = False
        Me.lvwChamados.ForeColor = vbRed
        Me.cmdCancelar.Enabled = False
        Me.cmdCancelar.Visible = False
End Select

Call suListarChamados(Me.cboStatus.Text)
End Sub

Private Sub cmdAtender_Click()
Dim res As Integer
If gintOSID = 0 Then
    MsgBox "Favor selecione um Atendimento!", vbExclamation, "Suporte Simac"
    Exit Sub
Else
   res = MsgBox("Deseja atender o Atendimento N° " & Format(gintOSID, "0000") + " ?", vbYesNo + vbInformation, "Confirmação")
    If res = vbYes Then
        If fnAtualizarOS(gintOSID) = True Then
        MsgBox "N° " & Format(gintOSID, "0000") & " em Atendimento!", vbInformation, "Suporte Simac"
        Call suListarChamados(Me.cboStatus.Text)
        gintOSID = 0
    Else
        Exit Sub
    End If

End If
End If

End Sub

Private Sub cmdAtendimento_Click()
If gintOSID = 0 Then
    MsgBox "Favor selecione um Atendimento!", vbExclamation, "Suporte Simac"
    Exit Sub
Else
    frmMensagem.Show (vbModal)
    
End If

End Sub

Private Sub cmdBaixarOS_Click()
If gintOSID = 0 Then
    MsgBox "Favor selecione um Atendimento!", vbExclamation, "Suporte Simac"
    Exit Sub
ElseIf VElemento = "" Then
    MsgBox "Favor, realize o atendimento!", vbInformation, "Suporte Simac"
    Me.cmdAtendimento.SetFocus
    Exit Sub
Else
    Call frmFinalizarOS.Show(vbModal)
End If
End Sub

Private Sub cmdCancelar_Click()
Dim frm As New frmFinalizarOS
If gintOSID = 0 Then
    MsgBox "Favor selecione um Atendimento!", vbExclamation, "Suporte Simac"
    Exit Sub
End If

frm.Caption = "Reporte Tecnico - (Cancelar Atendimento)"
frm.cmdFinalizar.Caption = "&Cancelar"
frm.Label2.Caption = "Motivo:"
Call frm.Show(vbModal)


End Sub

Private Sub cmdFechar_Click()
Call Unload(Me)
End Sub
Private Sub cmdLimpar_Click()
Call Atualizar
End Sub
Private Sub Atualizar()
If Me.cboStatus = "" Then
    Exit Sub
Else
    Me.txtOSID = ""
    Call suListarChamados(Me.cboStatus)
End If
End Sub



Private Sub cmdPesquisar_Click()
If Len(Trim(Me.cboStatus.Text)) = 0 Then
    MsgBox "Selecione um Status!", vbOKOnly + vbExclamation, "Suporte Simac"
    Me.cboStatus.SetFocus
    Exit Sub
End If
If Len(Trim(Me.txtOSID.Text)) > 0 Then
    If IsNumeric(Me.txtOSID.Text) = True Then
        If fnPesquisarOS(Trim(Me.txtOSID.Text), Me.cboStatus.Text) = False Then
        
        End If
    Else
        MsgBox "Digite um valor numérico!", vbOKOnly + vbExclamation, "Suporte Simac"
    End If
Else
    MsgBox "Digite o código da ordem de serviço!", vbOKOnly + vbExclamation, "Supote Simac"
End If
End Sub

Private Function fnPesquisarOS(ByVal vOSID As Long, ByVal vStatus As String) As Boolean

fnPesquisarOS = False
Call ConectarBD
Select Case vStatus
    Case "Em Aberto"
        strSQL = "SELECT * FROM vw_Chamados WHERE ATID = " & vOSID & " AND Status = 0 "
    Case "Em Atendimento"
        strSQL = "SELECT * FROM vw_Chamados WHERE ATID =  " & vOSID & " AND Status = 1 ORDER BY OSID"
    Case "Finalizada"
        strSQL = "SELECT * FROM vw_Chamados WHERE ATID =  " & vOSID & " AND Status = 2 ORDER BY OSID"
    Case "Cancelada"
        strSQL = "SELECT * FROM vw_Chamados WHERE ATID =  " & vOSID & " AND Status = 3 ORDER BY OSID"
 End Select
 
 Set rs = New ADODB.Recordset
 rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
 
 Me.lvwChamados.ListItems.Clear

 If Not rs.EOF Then
    With Me.lvwChamados.ListItems.Add(, , Format(rs!ATID, "0000"))
        If Len(rs!Nome) > 0 Then
            .SubItems(1) = rs!Nome
        Else
            .SubItems(1) = sFullName
        End If
        .SubItems(2) = rs!Tipo
        
        If IsNull(rs!Elemento) Then
            .SubItems(3) = ""
        Else
            .SubItems(3) = rs!Elemento
        End If
        If Len(rs!Departamento) > 0 Then
            .SubItems(4) = rs!Departamento
        Else
            .SubItems(4) = strDescription
        End If
        .SubItems(5) = rs!DataCadastro
        .SubItems(6) = rs!Local
        .SubItems(7) = rs!DescricaoMelhorias
        
    
        If IsNull(rs!PropostaResultado) Then
            .SubItems(8) = ""
        Else
            .SubItems(8) = rs!PropostaResultado
        End If
        
        .SubItems(9) = rs!Atendente
        
        If IsNull(rs!ReporteTecnico) Then
            .SubItems(10) = ""
        Else
            .SubItems(10) = rs!ReporteTecnico
        End If
        
        If IsNull(rs!FinalizacaoOS) Then
            .SubItems(11) = ""
        Else
            .SubItems(11) = rs!FinalizacaoOS
        End If
        
        If IsNull(rs!DataBaixaAtual) Then
            .SubItems(12) = ""
        Else
            .SubItems(12) = rs!DataBaixaAtual
        End If
        If IsNull(rs!MotivoCancelamento) Then
            .SubItems(13) = ""
        Else
            .SubItems(13) = rs!MotivoCancelamento
        End If
        
    End With
    fnPesquisarOS = True
    
Else
    MsgBox "Atendimento não localizado!", vbOKOnly + vbExclamation, "Suporte Simac"
End If
rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
End Function


Private Sub Form_Load()

With Me.lvwChamados
    .ColumnHeaders.Add , , "N°AT", 710
    .ColumnHeaders.Add , , "Nome", 1800, 2
    .ColumnHeaders.Add , , "Tipo", 1300, 2
    .ColumnHeaders.Add , , "Elemento", 3000, 2
    .ColumnHeaders.Add , , "Depto.", 1500, 2
    .ColumnHeaders.Add , , "Data Cadastro", 2500, 2
    .ColumnHeaders.Add , , "Local", 1800, 2
    .ColumnHeaders.Add , , "Descrição da Situação Atual", 3000, 2
    .ColumnHeaders.Add , , "Proposta de Melhoria e Resultado Esperado", 4000, 2
    .ColumnHeaders.Add , , "Atendente", 1500, 2
    .ColumnHeaders.Add , , "Reporte Técnico", 4000, 2
    .ColumnHeaders.Add , , "RT - Atendimento Concluído", 4000, 2
    .ColumnHeaders.Add , , "Data Finalizada", 2500, 2
    .ColumnHeaders.Add , , "", 0
    .LabelEdit = lvwManual
End With
Call suListarStatus
End Sub

Private Sub suListarStatus()
Me.cboStatus.AddItem "Em Aberto"
Me.cboStatus.AddItem "Em Atendimento"
Me.cboStatus.AddItem "Finalizada"
Me.cboStatus.AddItem "Cancelada"
End Sub


Private Function fnAtualizarOS(ByVal vOSID As Integer)
On Error GoTo Erro
     
        Call ConectarBD
        fnAtualizarOS = False
        Screen.MousePointer = 11
        Call Tempo
        
        strSQL = "UPDATE tb_Atendimento SET Status = 1 WHERE ATID =  " & vOSID & ""
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
        
    
        Sleep 5000
        fnAtualizarOS = True
        Set rs = Nothing
        Screen.MousePointer = 0
        Exit Function
Erro:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    MsgBox "Erro: " & Err.Description, vbCritical, "Suporte Simac"
    End Function

Private Sub suListarChamados(ByVal vStatus As String)
Dim lvwItem As ListItem

    
    Call ConectarBD
    Select Case vStatus
        Case Is = "Em Aberto"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 0"
        Case Is = "Em Atendimento"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 1"
        Case Is = "Finalizada"
            strSQL = "SELECT * FROM vw_Chamados WHERE Status = 2"
        Case Is = "Cancelada"
            strSQL = "SELECT * FROM  vw_Chamados WHERE Status = 3"
    End Select
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSQL, cn, adOpenKeyset, adLockOptimistic
    
    Me.lvwChamados.ListItems.Clear
    If rs.EOF = False Then
        Do While Not rs.EOF
            Set lvwItem = Me.lvwChamados.ListItems.Add(, , Format(rs!ATID, "0000"))
                If Len(rs!Nome) > 0 Then
                    lvwItem.SubItems(1) = rs!Nome
                Else
                    'lvwItem.SubItems(1) = sFullName
                End If
                
                lvwItem.SubItems(2) = rs!Tipo
                If IsNull(rs!Elemento) Then
                    lvwItem.SubItems(3) = ""
                Else
                    lvwItem.SubItems(3) = rs!Elemento
                End If
                If Len(rs!Departamento) > 0 Then
                    lvwItem.SubItems(4) = rs!Departamento
                Else
                    'lvwItem.SubItems(4) = strDescription
                End If
                lvwItem.SubItems(5) = rs!DataCadastro
                lvwItem.SubItems(6) = rs!Local
                lvwItem.SubItems(7) = rs!DescricaoMelhorias
                
                If IsNull(rs!PropostaResultado) Then
                    rs!PropostaResultado = ""
                Else
                    lvwItem.SubItems(8) = rs!PropostaResultado
                End If
                
                lvwItem.SubItems(9) = rs!Atendente
                
                If IsNull(rs!ReporteTecnico) Then
                    lvwItem.SubItems(10) = ""
                Else
                    lvwItem.SubItems(10) = rs!ReporteTecnico
                End If
                If IsNull(rs!FinalizacaoOS) Then
                    lvwItem.SubItems(11) = ""
                Else
                    lvwItem.SubItems(11) = rs!FinalizacaoOS
                End If
                
                If IsNull(rs!DataBaixaAtual) Then
                    lvwItem.SubItems(12) = ""
                Else
                    lvwItem.SubItems(12) = rs!DataBaixaAtual
                End If
                
                If IsNull(rs!MotivoCancelamento) Then
                    lvwItem.SubItems(13) = ""
                Else
                    lvwItem.SubItems(13) = rs!MotivoCancelamento
                End If
                
                rs.MoveNext
            Loop
            
        Else
            MsgBox "Nenhum registro foi encontrato!", vbOKOnly + vbInformation, "Consulta"
        End If
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
End Sub

Private Sub lvwChamados_Click()
    ValorElemento = ""
    gintOSID = 0
    
    For i = 1 To Me.lvwChamados.ListItems.Count
        If Me.lvwChamados.ListItems(i).Selected = True Then
            gintOSID = Me.lvwChamados.ListItems(i).Text
            VElemento = Me.lvwChamados.ListItems(i).SubItems(3)
        End If
    Next
End Sub

Private Sub lvwChamados_DblClick()

If Me.cboStatus = "Cancelada" Then
    Dim frm As New frmSuporteSistemasV
        If Not Me.lvwChamados.selectedItem Is Nothing Then
            Dim selectedItem As ListItem
            Set selectedItem = lvwChamados.selectedItem
        
            ValorTipo = selectedItem.SubItems(2)
            ValorElemento = selectedItem.SubItems(3)
            ValorLocal = selectedItem.SubItems(6)
            ValorAtendente = selectedItem.SubItems(9)
            ValorDescricao = selectedItem.SubItems(7)
            ValorReporte = selectedItem.SubItems(10)
            ValorResultado = selectedItem.SubItems(8)
            ValorFinalizado = selectedItem.SubItems(13)
            frm.Label3.Caption = "Reporte Técnico - Atendimento Cancelado:"
            frm.txtAtendimentoConcluido.ForeColor = vbRed
            frm.Show (vbModal)
        End If
Else
If Not Me.lvwChamados.selectedItem Is Nothing Then
            Dim selectedItem2 As ListItem
            Set selectedItem2 = lvwChamados.selectedItem
        
        ValorTipo = selectedItem2.SubItems(2)
        ValorElemento = selectedItem2.SubItems(3)
        ValorLocal = selectedItem2.SubItems(6)
        ValorAtendente = selectedItem2.SubItems(9)
        ValorDescricao = selectedItem2.SubItems(7)
        ValorReporte = selectedItem2.SubItems(10)
        ValorResultado = selectedItem2.SubItems(8)
        ValorFinalizado = selectedItem2.SubItems(11)
        frm.txtAtendimentoConcluido.ForeColor = vbBlue
        frm.Show (vbModal)
End If
End If

End Sub
