Attribute VB_Name = "IMTSolverUI"
'==============================================================
' IMTSolver 2.0 - botoes da faixa de opcoes
'
' Cada botao do customUI14.xml chama uma destas rotinas. Elas so
' encaminham: a logica esta em IMTSolver (nucleo) e nas janelas.
'==============================================================
Option Explicit

Public Sub RibbonModelo(control As IRibbonControl)
    On Error GoTo Falhou
    If Not AbaValida() Then Exit Sub
    frmModelo.Abrir ActiveSheet
    Exit Sub
Falhou:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, IMT_NOME
End Sub

Public Sub RibbonResolver(control As IRibbonControl)
    On Error GoTo Falhou
    If Not AbaValida() Then Exit Sub
    If Not TemModelo(ActiveSheet) Then
        If MsgBox("Esta aba ainda nao tem modelo." & vbLf & vbLf & _
                  OndeHaModelo(ActiveSheet) & vbLf & vbLf & _
                  "Quer definir um modelo nesta aba agora?", _
                  vbQuestion + vbYesNo, IMT_NOME) = vbYes Then frmModelo.Abrir ActiveSheet
        Exit Sub
    End If
    IMTResolver ActiveSheet
    Exit Sub
Falhou:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, IMT_NOME
End Sub

Public Sub RibbonVerificar(control As IRibbonControl)
    On Error GoTo Falhou
    If Not AbaValida() Then Exit Sub
    IMTVerificar ActiveSheet
    Exit Sub
Falhou:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, IMT_NOME
End Sub

Public Sub RibbonCancelar(control As IRibbonControl)
    IMTCancelar
End Sub

Public Sub RibbonMostrar(control As IRibbonControl)
    If AbaValida() Then IMTMostrar ActiveSheet
End Sub

Public Sub RibbonEsconder(control As IRibbonControl)
    If AbaValida() Then IMTEsconder ActiveSheet
End Sub

Public Sub RibbonLimpar(control As IRibbonControl)
    If AbaValida() Then IMTLimparSolucao ActiveSheet
End Sub

Public Sub RibbonResposta(control As IRibbonControl)
    IMTRelatorioResposta
End Sub

Public Sub RibbonSensibilidade(control As IRibbonControl)
    IMTRelatorioSensibilidade
End Sub

Public Sub RibbonSolvers(control As IRibbonControl)
    Dim lista As Collection, p As Variant, locais As String, neos As String
    Set lista = ListarSolvers()
    If lista.Count = 0 Then
        MsgBox "Nao consegui listar. O motor esta em:" & vbLf & CaminhoMotor() & vbLf & vbLf & _
               "Confira se o imtsolver.exe esta nessa pasta.", vbExclamation, IMT_NOME
        Exit Sub
    End If
    ' uma linha por solver: nome, o que resolve, licenca. Em duas linhas
    ' com recuo, a lista ficava esparramada e dificil de comparar.
    Dim tipos As String
    For Each p In lista
        If UBound(p) >= 4 Then tipos = p(4) Else tipos = "PL, PLIM"
        If p(3) = "neos" Then
            neos = neos & vbLf & "   " & p(1) & "  -  " & tipos
        Else
            locais = locais & vbLf & "   " & p(1) & "  -  " & tipos & _
                     "  -  " & p(2)
        End If
    Next p
    ' iif-ok: os dois ramos sao seguros (revisado)
    MsgBox "Nesta maquina:" & locais & vbLf & vbLf & _
           IIf(neos <> "", "Pelo NEOS (gratis, precisa do seu e-mail na janela Modelo):" & neos, _
               "NEOS: sem acesso a internet agora.") & vbLf & vbLf & _
           String$(62, "-") & vbLf & _
           "PL = linear   PLIM = linear inteira mista   QP = quadratica" & vbLf & _
           "MINLP = nao-linear inteira" & vbLf & vbLf & _
           "O IMTSolver so monta modelos LINEARES: a leitura mede os" & vbLf & _
           "coeficientes das suas formulas, e isso exige linearidade." & vbLf & _
           "O que cada solver sabe fazer alem disso nao e usado aqui.", _
           vbInformation, IMT_NOME & " - solvers"
End Sub

Public Sub RibbonExemplos(control As IRibbonControl)
    Dim arq As String
    arq = ThisWorkbook.Path & "\IMTSolver_Exemplos.xlsx"
    If Dir(arq) = "" Then
        MsgBox "Nao achei " & arq, vbExclamation, IMT_NOME
    Else
        Workbooks.Open arq
    End If
End Sub

Public Sub RibbonSobre(control As IRibbonControl)
    Dim lista As Collection, p As Variant, locais As Long, remotos As Long
    Set lista = ListarSolvers()
    For Each p In lista
        If p(3) = "neos" Then remotos = remotos + 1 Else locais = locais + 1
    Next p

    MsgBox IMT_NOME & " " & IMT_VERSAO & vbLf & _
           "Otimizacao linear e inteira dentro do Excel." & vbLf & vbLf & _
           "Criado por Pedro da Silva Bezerra, aluno do 3o periodo do" & vbLf & _
           "Instituto Maua de Tecnologia, para a disciplina de" & vbLf & _
           "Pesquisa Operacional I, com o incentivo da professora Joyce." & vbLf & vbLf & _
           String$(58, "-") & vbLf & vbLf & _
           "Resolve com o solver que voce escolher: CBC, HiGHS e SCIP " & _
           "na sua maquina, e CPLEX, COPT e SCIP de graca pelo NEOS." & vbLf & vbLf & _
           "O modelo fica guardado na planilha no mesmo formato do Solver " & _
           "do Excel, entao um modelo feito la abre aqui, e vice-versa." & vbLf & vbLf & _
           "Solvers disponiveis agora: " & locais & " nesta maquina" & _
           IIf(remotos > 0, " e " & remotos & " pelo NEOS", " (NEOS sem acesso)") & vbLf & _
           "Motor: " & CaminhoMotor(), _
           vbInformation, "Sobre o " & IMT_NOME
End Sub

Private Function AbaValida() As Boolean
    If ActiveSheet Is Nothing Then
        MsgBox "Abra uma planilha primeiro.", vbInformation, IMT_NOME
    ElseIf TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Esta aba nao e uma planilha comum.", vbInformation, IMT_NOME
    Else
        AbaValida = True
    End If
End Function
