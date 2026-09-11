Attribute VB_Name = "IMTSolver"
'==============================================================
' IMTSolver 2.0 - nucleo
'
' O modelo mora nos nomes solver_* da planilha, a mesma convencao
' do Solver do Excel e do OpenSolver. Um modelo feito em qualquer
' um dos tres abre nos outros dois, e viaja dentro do arquivo.
'
' Fluxo de um Resolver:
'   LerModelo      nomes solver_*  ->  estrutura Modelo
'   Extrair        perturbacao: zera tudo, liga uma variavel por vez
'   EscreverLP     arquivo .lp padrao
'   Lancar         imtsolver.exe escondido, sem travar o Excel
'   IMTPoll        a cada segundo le o progresso e atualiza a janela
'   Concluir       devolve a solucao e monta os relatorios
'
' O Excel nunca fala com solver nenhum. So com o motor.
'==============================================================
Option Explicit

Public Const IMT_VERSAO As String = "2.0"
Public Const IMT_NOME As String = "IMTSolver"

' relacoes, na mesma numeracao do Solver do Excel
Public Const REL_LE As Long = 1
Public Const REL_EQ As Long = 2
Public Const REL_GE As Long = 3
Public Const REL_INT As Long = 4
Public Const REL_BIN As Long = 5

' tipo de objetivo, idem
Public Const TIPO_MAX As Long = 1
Public Const TIPO_MIN As Long = 2
Public Const TIPO_VALOR As Long = 3

Public Type Restricao
    lhs As Range
    rel As Long
    rhs As Range            ' Nothing quando e constante
    rhsConst As Double
    rhsEhConst As Boolean
End Type

Public Type Modelo
    ws As Worksheet
    obj As Range            ' Nothing = so viabilidade
    tipo As Long
    valorAlvo As Double
    vars As Range           ' pode ter varias areas
    naoNeg As Boolean
    n As Long
    r() As Restricao
End Type

'--- estado da execucao em andamento ---
Private gM As Modelo
Private gPasta As String
Private gArqLp As String, gArqSol As String, gArqProg As String
Private gArqCanc As String, gArqLog As String
Private gInicio As Double
Private gProximo As Date
Private gRodando As Boolean
Private gSemProgressoDesde As Double

' variaveis: uma celula por indice, e o mapa inverso
Private gVarCel() As Range
Private gVarEnd() As String
Private gVarOriginal() As Variant
Private gNV As Long

' linhas do .lp -> de onde vieram
Private gRowBloco() As Long
Private gRowCel() As Long
Private gRowEnd() As String
Private gRowRel() As Long
Private gNR As Long

Private gObjOriginal As Variant

' modo de calculo do Excel antes da extracao, para devolver depois
Private gCalcAnterior As XlCalculation
Private gEventosAnterior As Boolean
Private gTelaAnterior As Boolean

' coeficientes esparsos, na ordem em que sao medidos (por variavel)
Private gNzRow() As Long, gNzVar() As Long, gNzVal() As Double, gNz As Long
Private gCObj() As Double
Private gRow0() As Double
Private gObj0 As Double

' Os lados esquerdos guardados COMO VIERAM do Excel, sem converter.
' Assim, a cada variavel, basta perguntar "mudou?" - uma comparacao crua -
' em vez de converter as 10.050 linhas para so entao comparar. Como uma
' variavel mexe em ~7 linhas, quase toda a conversao era jogada fora.
Private gNB As Long                 ' quantos blocos de restricao
Private gBlocoLhs() As Range
Private gBlocoIni() As Long         ' linha global da 1a celula do bloco
Private gBlocoLin() As Long
Private gBlocoCol() As Long
Private gBlocoRhsFixo() As Boolean  ' o lado direito nao depende das variaveis
Private gBase() As Variant          ' matriz do bloco com tudo zerado
Private gAtual() As Variant         ' matriz do bloco na medicao corrente
' A comparacao crua so vale quando nenhum lado direito tem formula. Se
' algum tiver, ele muda junto e a diferenca do lado esquerdo sozinha nao
' e o coeficiente - ai voltamos ao caminho antigo, que le os dois lados.
Private gPodeRapido As Boolean

' Ate onde procurar o rotulo de uma celula. Sem limite, um modelo com
' milhares de variaveis no fim da planilha varria milhares de celulas por
' variavel - dezenas de milhoes de leituras so para montar o relatorio.
Private Const ROTULO_COLS As Long = 16
Private Const ROTULO_LINHAS As Long = 200

' achados da verificacao do modelo
Private gAchGrav() As String     ' "Erro" ou "Aviso"
Private gAchOnde() As String     ' endereco da celula
Private gAchOque() As String
Private gAchFaca() As String
Private gAchPintar As Object     ' enderecos a marcar na planilha
Private gNAch As Long
Private gAchJaDito As Object
' Verdadeiro durante o Verificar. As rotinas que normalmente
' PERGUNTAM alguma coisa viram registro de achado: uma caixa de
' dialogo no meio de um parecer trava tudo esperando um clique.
Private gVerificando As Boolean
Private gBloqueia As Boolean     ' erro que impede ate ler o modelo

' resultado da ultima execucao
Private gSol As Object          ' cabecalho da solucao
Private gValores() As Double
Private gTipos As Object        ' endereco -> Inteira/Binaria
Private gSensVar As Collection
Private gSensRestr As Collection
Private gTemSolucao As Boolean

'==============================================================
' Nomes solver_* : ler e gravar
'==============================================================
Private Function RefDe(ws As Worksheet, ByVal nome As String) As String
    On Error Resume Next
    RefDe = ws.Names(nome).RefersTo
End Function

Private Function FaixaDe(ws As Worksheet, ByVal nome As String) As Range
    On Error Resume Next
    Set FaixaDe = ws.Names(nome).RefersToRange
End Function

Private Function NumeroDe(ws As Worksheet, ByVal nome As String, ByVal padrao As Double) As Double
    Dim s As String
    s = RefDe(ws, nome)
    If Len(s) < 2 Then
        NumeroDe = padrao
    Else
        NumeroDe = Val(Mid$(s, 2))      ' RefersTo vem sempre em ingles: "=5.5"
    End If
End Function

Private Sub Gravar(ws As Worksheet, ByVal nome As String, ByVal refersTo As String)
    On Error Resume Next
    ws.Names(nome).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nome, RefersTo:=refersTo, Visible:=False
End Sub

Private Sub Apagar(ws As Worksheet, ByVal nome As String)
    On Error Resume Next
    ws.Names(nome).Delete
End Sub

Private Function RefFaixa(r As Range) As String
    RefFaixa = "=" & r.Address(True, True, xlA1, True)
End Function

Public Function TemModelo(ws As Worksheet) As Boolean
    TemModelo = Not (FaixaDe(ws, "solver_adj") Is Nothing)
End Function

' Cada aba guarda o seu proprio modelo, nos nomes solver_* dela. Quem
' clica numa aba sem modelo quase sempre so esta na aba errada: dizemos
' em quais abas desta pasta ha modelo.
Public Function OndeHaModelo(ws As Worksheet) As String
    Dim w As Worksheet, lista As String, n As Long
    For Each w In ws.Parent.Worksheets
        If Not w Is ws Then
            If TemModelo(w) Then
                n = n + 1
                If n <= 6 Then
                    If n > 1 Then lista = lista & ", "
                    lista = lista & "'" & w.Name & "'"
                End If
            End If
        End If
    Next w
    If n > 6 Then lista = lista & " e mais " & (n - 6)
    If n = 0 Then
        OndeHaModelo = "Nenhuma aba desta pasta tem modelo ainda. " & _
                       "Cada aba guarda o seu proprio modelo."
    ElseIf n = 1 Then
        OndeHaModelo = "O modelo desta pasta esta na aba " & lista & ". " & _
                       "Cada aba guarda o seu proprio modelo: va para ela e clique de novo."
    Else
        OndeHaModelo = "Nesta pasta ha modelo nas abas " & lista & ". " & _
                       "Cada aba guarda o seu proprio modelo: va para a aba certa e clique de novo."
    End If
End Function

Public Function LerModelo(ws As Worksheet, m As Modelo) As Boolean
    Dim i As Long, s As String
    Set m.ws = ws
    Set m.obj = FaixaDe(ws, "solver_opt")
    m.tipo = NumeroDe(ws, "solver_typ", TIPO_MIN)
    m.valorAlvo = NumeroDe(ws, "solver_val", 0)
    Set m.vars = FaixaDe(ws, "solver_adj")
    m.naoNeg = (NumeroDe(ws, "solver_neg", 1) = 1)
    m.n = NumeroDe(ws, "solver_num", 0)
    If m.n > 0 Then
        ReDim m.r(1 To m.n)
        For i = 1 To m.n
            Set m.r(i).lhs = FaixaDe(ws, "solver_lhs" & i)
            m.r(i).rel = NumeroDe(ws, "solver_rel" & i, REL_LE)
            Set m.r(i).rhs = FaixaDe(ws, "solver_rhs" & i)
            If m.r(i).rhs Is Nothing Then
                s = RefDe(ws, "solver_rhs" & i)
                m.r(i).rhsEhConst = True
                m.r(i).rhsConst = Val(Replace(Mid$(s, 2), """", ""))
            End If
        Next i
    End If
    LerModelo = Not (m.vars Is Nothing)
End Function

Public Sub GravarModelo(m As Modelo)
    Dim ws As Worksheet, i As Long
    Set ws = m.ws
    If m.obj Is Nothing Then Apagar ws, "solver_opt" Else Gravar ws, "solver_opt", RefFaixa(m.obj)
    Gravar ws, "solver_typ", "=" & m.tipo
    Gravar ws, "solver_val", "=" & Trim$(Str$(m.valorAlvo))
    Gravar ws, "solver_adj", RefFaixa(m.vars)
    Gravar ws, "solver_neg", "=" & IIf(m.naoNeg, 1, 2)
    Gravar ws, "solver_num", "=" & m.n
    ' os que o Solver do Excel espera encontrar, para abrir sem reclamar
    Gravar ws, "solver_eng", "=2"
    Gravar ws, "solver_ver", "=3"
    For i = 1 To m.n
        Gravar ws, "solver_lhs" & i, RefFaixa(m.r(i).lhs)
        Gravar ws, "solver_rel" & i, "=" & m.r(i).rel
        If m.r(i).rel = REL_INT Then
            Gravar ws, "solver_rhs" & i, "=""integer"""
        ElseIf m.r(i).rel = REL_BIN Then
            Gravar ws, "solver_rhs" & i, "=""binary"""
        ElseIf m.r(i).rhsEhConst Then
            Gravar ws, "solver_rhs" & i, "=" & Trim$(Str$(m.r(i).rhsConst))
        Else
            Gravar ws, "solver_rhs" & i, RefFaixa(m.r(i).rhs)
        End If
    Next i
    ' sobras de um modelo anterior maior
    i = m.n + 1
    Do While Len(RefDe(ws, "solver_lhs" & i)) > 0
        Apagar ws, "solver_lhs" & i: Apagar ws, "solver_rel" & i: Apagar ws, "solver_rhs" & i
        i = i + 1
    Loop
End Sub

Public Sub ApagarModelo(ws As Worksheet)
    ' ws.Names so traz os nomes com escopo desta aba
    Dim k As Long
    For k = ws.Names.Count To 1 Step -1
        If InStr(1, ws.Names(k).Name, "solver_", vbTextCompare) > 0 Then ws.Names(k).Delete
    Next k
End Sub

' relatorios avulsos, depois de uma execucao
Public Sub IMTRelatorioResposta()
    If Not gTemSolucao Then
        MsgBox "Ainda nao ha solucao nesta sessao. Resolva primeiro.", vbInformation, IMT_NOME
    Else
        RelatorioResposta
    End If
End Sub

Public Sub IMTRelatorioSensibilidade()
    If gSensVar Is Nothing Then
        MsgBox "Ainda nao ha solucao nesta sessao. Resolva primeiro.", vbInformation, IMT_NOME
    ElseIf gSensVar.Count = 0 Then
        MsgBox "A ultima execucao nao gerou sensibilidade. Marque a opcao no modelo e resolva de novo.", vbInformation, IMT_NOME
    Else
        RelatorioSensibilidade
    End If
End Sub

'==============================================================
' Opcoes do usuario (registro do Windows; nao vao para o arquivo)
'==============================================================
Public Function Opcao(ByVal chave As String, ByVal padrao As String) As String
    Opcao = GetSetting(IMT_NOME, "Opcoes", chave, padrao)
End Function

Public Sub GravarOpcao(ByVal chave As String, ByVal valor As String)
    SaveSetting IMT_NOME, "Opcoes", chave, valor
End Sub

Public Function CaminhoMotor() As String
    ' O motor agora e uma pasta (o exe mais as bibliotecas ao lado dele).
    ' Em arquivo unico ele levava 1,4 s so para se descompactar a cada
    ' chamada; em pasta comeca em 0,1 s. A forma antiga continua valendo
    ' para quem ja tinha instalado assim.
    Dim c As String, base As String
    c = Opcao("motor", "")
    If c <> "" Then
        CaminhoMotor = c
        Exit Function
    End If
    base = ThisWorkbook.Path
    c = base & "\imtsolver\imtsolver.exe"
    If Dir(c) = "" Then c = base & "\imtsolver.exe"
    ' quem instalou antes da mudanca de nome ainda tem o motor antigo
    If Dir(c) = "" Then c = base & "\byesolver\byesolver.exe"
    If Dir(c) = "" Then c = base & "\byesolver.exe"
    CaminhoMotor = c
End Function

Public Function ListarSolvers(Optional ByVal forcar As Boolean = False) As Collection
    ' Collection de arrays (nome, rotulo, licenca, origem).
    '
    ' A lista muda so quando alguem instala ou remove um solver, entao
    ' fica guardada. Perguntar ao motor custa um processo novo, e isso
    ' acontecia toda vez que a janela do modelo abria.
    Dim linha As Variant, p() As String
    Dim lista As New Collection, guardado As String

    If Not forcar Then
        guardado = Opcao("solvers", "")
        If Len(guardado) > 0 Then
            For Each linha In Split(guardado, vbLf)
                p = Split(CStr(linha), "|")
                ' 5 campos: nome, rotulo, licenca, origem, tipos. Uma lista
                ' guardada por versao antiga tem 4 e e descartada, senao o
                ' codigo que le p(4) quebra.
                If UBound(p) >= 4 Then lista.Add p
            Next linha
            If lista.Count > 0 Then
                Set ListarSolvers = lista
                Exit Function
            End If
        End If
    End If

    Set lista = PerguntarSolvers()
    If lista.Count > 0 Then GravarOpcao "solvers", TextoDosSolvers(lista)
    Set ListarSolvers = lista
End Function

Private Function PerguntarSolvers() As Collection
    Dim sh As Object, arq As String, f As Integer, linha As String, p() As String
    Dim lista As New Collection
    Set PerguntarSolvers = lista
    If Dir(CaminhoMotor()) = "" Then Exit Function
    arq = PastaTemp() & "\solvers.txt"
    On Error Resume Next
    Kill arq
    On Error GoTo 0
    Set sh = CreateObject("WScript.Shell")
    sh.Run """" & CaminhoMotor() & """ --list --out """ & arq & """", 0, True
    If Dir(arq) = "" Then Exit Function
    f = FreeFile
    Open arq For Input As #f
    Do While Not EOF(f)
        Line Input #f, linha
        p = Split(linha, "|")
        If UBound(p) >= 3 Then lista.Add p
    Loop
    Close #f
End Function

Private Function TextoDosSolvers(lista As Collection) As String
    Dim p As Variant, s As String
    For Each p In lista
        s = s & IIf(Len(s) > 0, vbLf, "") & Join(p, "|")
    Next p
    TextoDosSolvers = s
End Function

'==============================================================
' Resolver: ponto de entrada
'==============================================================
Public Sub IMTResolver(Optional ws As Worksheet = Nothing)
    Dim m As Modelo
    If ws Is Nothing Then Set ws = ActiveSheet
    If gRodando Then
        MsgBox "Ja tem uma otimizacao rodando. Espere terminar ou cancele.", vbExclamation, IMT_NOME
        Exit Sub
    End If
    If Not LerModelo(ws, m) Then
        MsgBox "Esta aba nao tem modelo." & vbLf & vbLf & OndeHaModelo(ws) & _
               vbLf & vbLf & "Para criar um modelo aqui: IMTSolver > Modelo.", _
               vbInformation, IMT_NOME
        Exit Sub
    End If
    If Dir(CaminhoMotor()) = "" Then
        MsgBox "Nao achei o motor em:" & vbLf & CaminhoMotor() & vbLf & vbLf & _
               "Ele precisa ficar na mesma pasta do suplemento, ou informe o " & _
               "caminho em Opcoes.", vbCritical, IMT_NOME
        Exit Sub
    End If

    gM = m
    gRodando = True
    On Error GoTo Falhou

    frmProgresso.Abrir Opcao("solver", "HiGHS")
    frmProgresso.Fase "Lendo o modelo da planilha"
    DoEvents

    If Not Extrair() Then GoTo Desistiu
    frmProgresso.Fase "Escrevendo o arquivo .lp"
    DoEvents
    PrepararPasta
    EscreverLP
    frmProgresso.Fase "Iniciando o solver"
    Lancar
    Exit Sub

Falhou:
    frmProgresso.Erro "Erro " & Err.Number & ": " & Err.Description
Desistiu:
    RestaurarCalculo
    gRodando = False
End Sub

'==============================================================
' Extracao por perturbacao
'==============================================================
Private Sub PrepararCalculo()
    gCalcAnterior = Application.Calculation
    gEventosAnterior = Application.EnableEvents
    gTelaAnterior = Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Private Sub RestaurarCalculo()
    On Error Resume Next
    Application.Calculation = gCalcAnterior
    Application.EnableEvents = gEventosAnterior
    Application.ScreenUpdating = gTelaAnterior
    If gCalcAnterior = 0 Then Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub AnotaNz(ByVal r As Long, ByVal v As Long, ByVal x As Double)
    If gNz = 0 Then
        ReDim gNzRow(1 To 4096): ReDim gNzVar(1 To 4096): ReDim gNzVal(1 To 4096)
    ElseIf gNz >= UBound(gNzRow) Then
        ReDim Preserve gNzRow(1 To 2 * gNz): ReDim Preserve gNzVar(1 To 2 * gNz)
        ReDim Preserve gNzVal(1 To 2 * gNz)
    End If
    gNz = gNz + 1
    gNzRow(gNz) = r: gNzVar(gNz) = v: gNzVal(gNz) = x
End Sub

'--------------------------------------------------------------
' Leitura rapida: comparar cru em vez de converter tudo
'--------------------------------------------------------------
Private Sub MontarBlocos()
    Dim i As Long, b As Long, j As Long, r As Range, hf As Variant
    gNB = 0
    For i = 1 To gM.n
        If gM.r(i).rel <= REL_GE Then gNB = gNB + 1
    Next i
    If gNB = 0 Then Exit Sub
    ReDim gBlocoLhs(1 To gNB): ReDim gBlocoIni(1 To gNB)
    ReDim gBlocoLin(1 To gNB): ReDim gBlocoCol(1 To gNB)
    ReDim gBlocoRhsFixo(1 To gNB)
    ReDim gBase(1 To gNB): ReDim gAtual(1 To gNB)
    b = 0: j = 0
    For i = 1 To gM.n
        If gM.r(i).rel <= REL_GE Then
            b = b + 1
            Set r = gM.r(i).lhs
            Set gBlocoLhs(b) = r
            gBlocoIni(b) = j + 1
            gBlocoLin(b) = r.Rows.Count
            gBlocoCol(b) = r.Columns.Count
            ' Se o lado direito nao tem formula, ele nao muda quando
            ' mexemos numa variavel: nem precisa ser relido.
            '
            ' Cuidado com HasFormula: numa faixa MISTA ele devolve Null, e
            ' "Not (Null = False)" tambem da Null, que o If trata como
            ' falso - a trava passava batido justamente no caso perigoso.
            gBlocoRhsFixo(b) = True
            If Not gM.r(i).rhsEhConst Then
                hf = gM.r(i).rhs.HasFormula
                If IsNull(hf) Then
                    gBlocoRhsFixo(b) = False        ' parte da faixa tem formula
                ElseIf hf Then
                    gBlocoRhsFixo(b) = False
                End If
            End If
            j = j + r.Cells.Count
        End If
    Next i
    gPodeRapido = True
    For b = 1 To gNB
        If Not gBlocoRhsFixo(b) Then gPodeRapido = False
        If gBlocoLhs(b).Areas.Count > 1 Then gPodeRapido = False
    Next b
End Sub

Private Sub LerBlocos(destino() As Variant)
    Dim b As Long
    For b = 1 To gNB
        destino(b) = FaixaComoMatriz(gBlocoLhs(b))
    Next b
End Sub

Private Function MedirVariavel(ByVal v As Long) As Boolean
    ' Anota os coeficientes da variavel v comparando o estado atual com a
    ' base. Devolve False se alguma celula deu erro - ai quem chama refaz
    ' pelo caminho lento, que sabe dizer qual celula foi.
    Dim b As Long, r As Long, c As Long, j As Long, nl As Long, nc As Long
    Dim d As Double, antes As Long
    Dim a As Variant, z As Variant

    antes = gNz                 ' para desfazer se parar no meio
    On Error GoTo Falhou
    For b = 1 To gNB
        a = gAtual(b): z = gBase(b)
        nl = gBlocoLin(b): nc = gBlocoCol(b)
        j = gBlocoIni(b) - 1
        For r = 1 To nl
            For c = 1 To nc
                j = j + 1
                If a(r, c) <> z(r, c) Then
                    ' so aqui vale a pena converter
                    d = Val0(a(r, c)) - Val0(z(r, c))
                    If Abs(d) > 1E-12 Then AnotaNz j, v, d
                End If
            Next c
        Next r
    Next b
    MedirVariavel = True
    Exit Function
Falhou:
    gNz = antes                 ' apaga o que ja tinha anotado desta variavel
    MedirVariavel = False
End Function

Private Function LerLinhas(valores() As Double) As String
    ' Le todas as linhas (lhs - rhs) do modelo para valores(1..gNR).
    ' Devolve "" se deu certo, ou o endereco da celula com erro.
    Dim i As Long, k As Long, j As Long
    Dim a As Variant, b As Variant, av As Variant, bv As Variant
    j = 0
    For i = 1 To gM.n
        If gM.r(i).rel <= REL_GE Then
            a = gM.r(i).lhs.Value2
            If Not gM.r(i).rhsEhConst Then b = gM.r(i).rhs.Value2
            For k = 1 To gM.r(i).lhs.Cells.Count
                j = j + 1
                av = Pega(a, k)
                If IsError(av) Then LerLinhas = gRowEnd(j): Exit Function
                If gM.r(i).rhsEhConst Then
                    bv = gM.r(i).rhsConst
                ElseIf gM.r(i).rhs.Cells.Count = 1 Then
                    bv = Pega(b, 1)
                Else
                    bv = Pega(b, k)
                End If
                If IsError(bv) Then LerLinhas = gM.r(i).rhs.Cells(k).Address(False, False): Exit Function
                valores(j) = CDbl(Val0(av)) - CDbl(Val0(bv))
            Next k
        End If
    Next i
    LerLinhas = ""
End Function

Private Function Pega(a As Variant, ByVal k As Long) As Variant
    ' k-esimo elemento de um Value2 (escalar ou matriz 2D, linha a linha)
    Dim nc As Long
    If Not IsArray(a) Then
        Pega = a
    Else
        nc = UBound(a, 2) - LBound(a, 2) + 1
        Pega = a(LBound(a, 1) + (k - 1) \ nc, LBound(a, 2) + (k - 1) Mod nc)
    End If
End Function

Private Function Val0(v As Variant) As Double
    If IsEmpty(v) Or IsNull(v) Then
        Val0 = 0
    ElseIf VarType(v) = vbString Then
        Val0 = Val(Replace(v, ",", "."))
    Else
        Val0 = CDbl(v)
    End If
End Function

Private Function LerObjetivo() As Variant
    If gM.obj Is Nothing Then
        LerObjetivo = 0
    Else
        LerObjetivo = gM.obj.Value2
    End If
End Function

Private Function Extrair() As Boolean
    Dim i As Long, k As Long, v As Long, j As Long
    Dim area As Range, cel As Range
    Dim linhaAtual() As Double, erro As String, ov As Variant
    Dim dic As Object

    '--- variaveis ---
    gNV = gM.vars.Cells.Count
    ReDim gVarCel(1 To gNV): ReDim gVarEnd(1 To gNV): ReDim gVarOriginal(1 To gNV)
    Set dic = CreateObject("Scripting.Dictionary")
    v = 0
    For Each area In gM.vars.Areas
        For Each cel In area.Cells
            v = v + 1
            Set gVarCel(v) = cel
            gVarEnd(v) = cel.Address(False, False)
            gVarOriginal(v) = cel.Value2
            dic(gVarEnd(v)) = v
            If cel.HasFormula Then
                MsgBox "A variavel " & gVarEnd(v) & " tem formula. Variaveis de " & _
                       "decisao precisam ser celulas de valor.", vbExclamation, IMT_NOME
                Exit Function
            End If
        Next cel
    Next area

    '--- linhas ---
    gNR = 0
    For i = 1 To gM.n
        If gM.r(i).rel <= REL_GE Then
            If Not gM.r(i).rhsEhConst Then
                If gM.r(i).rhs.Cells.Count <> 1 And gM.r(i).rhs.Cells.Count <> gM.r(i).lhs.Cells.Count Then
                    MsgBox "Restricao " & i & ": o lado direito (" & gM.r(i).rhs.Address(False, False) & _
                           ") precisa ter uma celula, ou o mesmo tamanho do lado esquerdo.", vbExclamation, IMT_NOME
                    Exit Function
                End If
            End If
            gNR = gNR + gM.r(i).lhs.Cells.Count
        End If
    Next i
    If gNR = 0 And gM.obj Is Nothing Then
        MsgBox "O modelo nao tem objetivo nem restricoes.", vbExclamation, IMT_NOME
        Exit Function
    End If
    ReDim gRowBloco(1 To IIf(gNR = 0, 1, gNR)): ReDim gRowCel(1 To IIf(gNR = 0, 1, gNR))
    ReDim gRowEnd(1 To IIf(gNR = 0, 1, gNR)): ReDim gRowRel(1 To IIf(gNR = 0, 1, gNR))
    j = 0
    For i = 1 To gM.n
        If gM.r(i).rel <= REL_GE Then
            For k = 1 To gM.r(i).lhs.Cells.Count
                j = j + 1
                gRowBloco(j) = i: gRowCel(j) = k: gRowRel(j) = gM.r(i).rel
                gRowEnd(j) = gM.r(i).lhs.Cells(k).Address(False, False)
            Next k
        End If
    Next i

    '--- medicoes ---
    ReDim gRow0(1 To IIf(gNR = 0, 1, gNR)): ReDim linhaAtual(1 To IIf(gNR = 0, 1, gNR))
    ReDim gCObj(1 To gNV)
    gNz = 0
    PrepararCalculo
    gObjOriginal = LerObjetivo()

    For v = 1 To gNV: gVarCel(v).Value2 = 0: Next v
    gM.ws.Calculate
    If gM.obj Is Nothing Then
        gObj0 = 0
    Else
        ov = LerObjetivo()
        If IsError(ov) Then
            Call Desfaz: MsgBox "O objetivo " & gM.obj.Address(False, False) & " da erro com as variaveis zeradas.", vbExclamation, IMT_NOME
            Exit Function
        End If
        gObj0 = Val0(ov)
    End If
    erro = LerLinhas(gRow0)
    If erro <> "" Then
        Call Desfaz: MsgBox "A restricao em " & erro & " da erro com as variaveis zeradas.", vbExclamation, IMT_NOME
        Exit Function
    End If

    MontarBlocos
    If gPodeRapido Then LerBlocos gBase

    Dim t0 As Double: t0 = Timer
    For v = 1 To gNV
        gVarCel(v).Value2 = 1
        gM.ws.Calculate
        If Not gM.obj Is Nothing Then
            ov = LerObjetivo()
            If IsError(ov) Then
                Call Desfaz: MsgBox "O objetivo da erro quando " & gVarEnd(v) & " = 1.", vbExclamation, IMT_NOME
                Exit Function
            End If
            gCObj(v) = Val0(ov) - gObj0
        End If

        If gPodeRapido Then
            LerBlocos gAtual
            If Not MedirVariavel(v) Then gPodeRapido = False   ' celula com erro
        End If
        If Not gPodeRapido Then
            erro = LerLinhas(linhaAtual)
            If erro <> "" Then
                Call Desfaz: MsgBox "A restricao em " & erro & " da erro quando " & gVarEnd(v) & " = 1.", vbExclamation, IMT_NOME
                Exit Function
            End If
            For j = 1 To gNR
                If Abs(linhaAtual(j) - gRow0(j)) > 1E-12 Then AnotaNz j, v, linhaAtual(j) - gRow0(j)
            Next j
        End If

        gVarCel(v).Value2 = 0
        If v Mod 50 = 0 Or v = gNV Then
            frmProgresso.Lendo v, gNV, Timer - t0
            If frmProgresso.Cancelado Then Desfaz: Exit Function
        End If
    Next v

    '--- teste de linearidade num segundo ponto ---
    If Not TestarLinearidade(linhaAtual) Then Desfaz: Exit Function

    Desfaz
    Extrair = True
End Function

Private Sub Desfaz()
    Dim v As Long
    On Error Resume Next
    For v = 1 To gNV: gVarCel(v).Value2 = gVarOriginal(v): Next v
    gM.ws.Calculate
    RestaurarCalculo
End Sub

Private Function TestarLinearidade(buffer() As Double) As Boolean
    ' Poe cada variavel num valor e confere se as linhas batem com
    ' constante + soma dos coeficientes. Se nao bate, o modelo nao e linear
    ' e os coeficientes medidos nao descrevem o que esta na planilha.
    '
    ' Os pesos sao SORTEADOS. Antes eram 1 + (v mod 3), e isso deixava um
    ' furo: duas variaveis com o mesmo resto recebiam o mesmo peso, entao
    ' um erro do tipo "+1 numa e -1 na outra" se cancelava e a prova
    ' passava sem acusar nada. Com sorteio isso deixa de acontecer.
    Dim v As Long, j As Long, k As Long, previsto() As Double, erro As String
    Dim ruins As String, nRuins As Long, peso() As Double, escala() As Double
    Dim t As Double
    If gNR = 0 Then TestarLinearidade = True: Exit Function
    ReDim previsto(1 To gNR): ReDim peso(1 To gNV): ReDim escala(1 To gNR)
    Randomize
    For v = 1 To gNV
        ' 1 a 97: sorteado, mas sem inflar os produtos a ponto de o
        ' cancelamento entre termos grandes virar ruido de arredondamento
        peso(v) = Int(Rnd * 97) + 1
        gVarCel(v).Value2 = peso(v)
    Next v
    For j = 1 To gNR
        previsto(j) = gRow0(j): escala(j) = Abs(gRow0(j))
    Next j
    For k = 1 To gNz
        t = gNzVal(k) * peso(gNzVar(k))
        previsto(gNzRow(k)) = previsto(gNzRow(k)) + t
        escala(gNzRow(k)) = escala(gNzRow(k)) + Abs(t)
    Next k
    gM.ws.Calculate
    erro = LerLinhas(buffer)
    If erro <> "" Then TestarLinearidade = True: Exit Function   ' erro aqui nao e conclusivo
    For j = 1 To gNR
        ' a folga acompanha o tamanho dos termos somados, nao so o
        ' resultado: numa linha que se cancela o resultado e ~0 mas o
        ' arredondamento acumulado nao e
        If Abs(buffer(j) - previsto(j)) > 0.000000001 * escala(j) + 0.000001 Then
            nRuins = nRuins + 1
            If nRuins <= 8 Then ruins = ruins & vbLf & "   " & gRowEnd(j)
        End If
    Next j
    ' o objetivo tambem
    Dim ov As Variant, prevObj As Double, escObj As Double
    If Not gM.obj Is Nothing Then
        prevObj = gObj0: escObj = Abs(gObj0)
        For v = 1 To gNV
            prevObj = prevObj + gCObj(v) * peso(v)
            escObj = escObj + Abs(gCObj(v) * peso(v))
        Next v
        ov = LerObjetivo()
        If Not IsError(ov) Then
            If Abs(Val0(ov) - prevObj) > 0.000000001 * escObj + 0.000001 Then
                nRuins = nRuins + 1
                ruins = ruins & vbLf & "   " & gM.obj.Address(False, False) & " (objetivo)"
            End If
        End If
    End If
    If nRuins > 0 And gVerificando Then
        Achado "Aviso", "", _
               "O modelo nao parece linear: " & nRuins & " formula(s) nao " & _
               "batem com coeficientes constantes." & ruins, _
               "Procure SE, MAXIMO, MINIMO, ABS, ARRED ou variavel " & _
               "multiplicando variavel. O resultado de um solver linear nao " & _
               "vale para um modelo desses."
        TestarLinearidade = True
        Exit Function
    End If
    If nRuins > 0 Then
        If MsgBox("O modelo NAO parece linear. " & nRuins & " formula(s) nao batem com " & _
                  "coeficientes constantes:" & ruins & IIf(nRuins > 8, vbLf & "   ...", "") & vbLf & vbLf & _
                  "Procure SE, MAXIMO, MINIMO, ABS, ARRED ou variavel multiplicando " & _
                  "variavel. O resultado de um solver linear nao vale para esse modelo." & vbLf & vbLf & _
                  "Quer continuar mesmo assim?", vbExclamation + vbYesNo + vbDefaultButton2, IMT_NOME) = vbNo Then
            Exit Function
        End If
    End If
    TestarLinearidade = True
End Function

'==============================================================
' Arquivo .lp
'==============================================================
Private Function Num(ByVal d As Double) As String
    Num = Trim$(Str$(d))                ' Str$ usa ponto em qualquer idioma
    If Left$(Num, 1) = "." Then Num = "0" & Num
    If Left$(Num, 2) = "-." Then Num = "-0" & Mid$(Num, 2)
End Function

Private Function Termo(ByVal coef As Double, ByVal v As Long) As String
    If coef < 0 Then
        Termo = " - " & Num(-coef) & " x" & v
    Else
        Termo = " + " & Num(coef) & " x" & v
    End If
End Function

Private Sub PrepararPasta()
    gPasta = PastaTemp() & "\" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir gPasta
    gArqLp = gPasta & "\modelo.lp"
    gArqSol = gPasta & "\solucao.txt"
    gArqProg = gPasta & "\progresso.txt"
    gArqCanc = gPasta & "\cancelar.flag"
    gArqLog = gPasta & "\solver.log"
End Sub

Public Function PastaTemp() As String
    PastaTemp = Environ$("TEMP") & "\IMTSolver"
    If Dir(PastaTemp, vbDirectory) = "" Then MkDir PastaTemp
End Function

Private Sub EscreverLP()
    Dim f As Integer, v As Long, j As Long, k As Long, i As Long
    Dim rowStart() As Long, ordVar() As Long, ordVal() As Double, pos() As Long
    Dim ehInt() As Boolean, ehBin() As Boolean, dic As Object, cel As Range
    Dim linha As String, nTermos As Long, rhs As Double, sinal As String

    ' CSR: reordena os coeficientes por linha
    ReDim rowStart(1 To gNR + 2)
    For k = 1 To gNz: rowStart(gNzRow(k) + 1) = rowStart(gNzRow(k) + 1) + 1: Next k
    rowStart(1) = 1
    For j = 1 To gNR: rowStart(j + 1) = rowStart(j + 1) + rowStart(j): Next j
    ReDim pos(1 To gNR + 1)
    For j = 1 To gNR + 1: pos(j) = rowStart(j): Next j
    ReDim ordVar(1 To IIf(gNz = 0, 1, gNz)): ReDim ordVal(1 To IIf(gNz = 0, 1, gNz))
    For k = 1 To gNz
        ordVar(pos(gNzRow(k))) = gNzVar(k): ordVal(pos(gNzRow(k))) = gNzVal(k)
        pos(gNzRow(k)) = pos(gNzRow(k)) + 1
    Next k

    ' inteiras e binarias, pelo endereco
    ReDim ehInt(1 To gNV): ReDim ehBin(1 To gNV)
    Set dic = CreateObject("Scripting.Dictionary")
    For v = 1 To gNV: dic(gVarEnd(v)) = v: Next v
    For i = 1 To gM.n
        If gM.r(i).rel = REL_INT Or gM.r(i).rel = REL_BIN Then
            For Each cel In gM.r(i).lhs.Cells
                If dic.Exists(cel.Address(False, False)) Then
                    If gM.r(i).rel = REL_INT Then ehInt(dic(cel.Address(False, False))) = True
                    If gM.r(i).rel = REL_BIN Then ehBin(dic(cel.Address(False, False))) = True
                End If
            Next cel
        End If
    Next i

    f = FreeFile
    Open gArqLp For Output As #f
    Print #f, "\ " & IMT_NOME & " " & IMT_VERSAO & " - " & gM.ws.Parent.Name & " / " & gM.ws.Name

    ' objetivo
    If gM.tipo = TIPO_MAX Then Print #f, "Maximize" Else Print #f, "Minimize"
    linha = "obj:": nTermos = 0
    If gM.tipo <> TIPO_VALOR And Not gM.obj Is Nothing Then
        For v = 1 To gNV
            If gCObj(v) <> 0 Then
                linha = linha & Termo(gCObj(v), v): nTermos = nTermos + 1
                If nTermos Mod 8 = 0 Then Print #f, linha: linha = ""
            End If
        Next v
    End If
    If nTermos = 0 And linha = "obj:" Then linha = "obj: 0 x1"
    If linha <> "" Then Print #f, linha

    Print #f, "Subject To"
    For j = 1 To gNR
        linha = "c" & j & ":": nTermos = 0
        For k = rowStart(j) To rowStart(j + 1) - 1
            linha = linha & Termo(ordVal(k), ordVar(k)): nTermos = nTermos + 1
            If nTermos Mod 8 = 0 Then Print #f, linha: linha = ""
        Next k
        If nTermos = 0 Then linha = linha & " 0 x1"
        Select Case gRowRel(j)
            Case REL_LE: sinal = " <= "
            Case REL_GE: sinal = " >= "
            Case Else: sinal = " = "
        End Select
        Print #f, linha & sinal & Num(-gRow0(j))
    Next j
    If gM.tipo = TIPO_VALOR And Not gM.obj Is Nothing Then
        linha = "alvo:": nTermos = 0
        For v = 1 To gNV
            If gCObj(v) <> 0 Then
                linha = linha & Termo(gCObj(v), v): nTermos = nTermos + 1
                If nTermos Mod 8 = 0 Then Print #f, linha: linha = ""
            End If
        Next v
        If nTermos = 0 Then linha = linha & " 0 x1"
        Print #f, linha & " = " & Num(gM.valorAlvo - gObj0)
    End If

    ' limites: sem "nao-negativas", tudo que nao e binaria fica livre
    Print #f, "Bounds"
    For v = 1 To gNV
        If ehBin(v) Then
            Print #f, "0 <= x" & v & " <= 1"
        ElseIf Not gM.naoNeg Then
            Print #f, "x" & v & " free"
        End If
    Next v
    nTermos = 0
    For v = 1 To gNV
        If ehInt(v) And Not ehBin(v) Then
            If nTermos = 0 Then Print #f, "Generals"
            Print #f, "x" & v: nTermos = nTermos + 1
        End If
    Next v
    nTermos = 0
    For v = 1 To gNV
        If ehBin(v) Then
            If nTermos = 0 Then Print #f, "Binaries"
            Print #f, "x" & v: nTermos = nTermos + 1
        End If
    Next v
    Print #f, "End"
    Close #f
End Sub

'==============================================================
' Lancar o motor e acompanhar
'==============================================================
Private Sub Lancar()
    Dim cmd As String, sh As Object, solver As String, email As String
    solver = Opcao("solver", "HiGHS")
    cmd = """" & CaminhoMotor() & """ --lp """ & gArqLp & """ --solver " & solver & _
          " --out """ & gArqSol & """ --progress """ & gArqProg & """" & _
          " --cancel """ & gArqCanc & """ --log """ & gArqLog & """"
    If Val(Opcao("tempo", "300")) > 0 Then cmd = cmd & " --timelimit " & Val(Opcao("tempo", "300"))
    cmd = cmd & " --gap " & Num(Val(Replace(Opcao("gap", "0"), ",", ".")))
    If Opcao("sens", "1") = "1" Then cmd = cmd & " --sens"
    email = Opcao("email", "")
    If email <> "" Then cmd = cmd & " --email " & email
    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd, 0, False                     ' escondido, sem esperar
    gInicio = Timer
    gSemProgressoDesde = Timer
    frmProgresso.Iniciou solver, gArqLog
    Agendar
End Sub

Private Sub Agendar()
    gProximo = Now + TimeSerial(0, 0, 1)
    Application.OnTime gProximo, "'" & ThisWorkbook.Name & "'!IMTPoll"
End Sub

Public Sub IMTPoll()
    Dim d As Object, fase As String
    If Not gRodando Then Exit Sub
    On Error GoTo Falhou
    Set d = LerKv(gArqProg)
    If d Is Nothing Then
        If Timer - gSemProgressoDesde > 20 Then
            frmProgresso.Erro "O motor nao respondeu em 20 segundos. Veja se o antivirus bloqueou o imtsolver.exe."
            gRodando = False
            Exit Sub
        End If
        Agendar
        Exit Sub
    End If
    gSemProgressoDesde = Timer
    fase = d("fase")
    frmProgresso.Atualizar d, Timer - gInicio
    If fase = "concluido" Or fase = "cancelado" Or fase = "erro" Then
        gRodando = False
        Concluir
    Else
        Agendar
    End If
    Exit Sub
Falhou:
    gRodando = False
    frmProgresso.Erro "Erro " & Err.Number & " acompanhando o solver: " & Err.Description
End Sub

Public Sub IMTCancelar()
    Dim f As Integer
    If Not gRodando Then Exit Sub
    f = FreeFile
    Open gArqCanc For Output As #f
    Print #f, "cancelar"
    Close #f
    frmProgresso.Fase "Pedindo ao solver para parar..."
End Sub

Public Function LerKv(ByVal arq As String) As Object
    Dim f As Integer, linha As String, p As Long, d As Object
    If Dir(arq) = "" Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo Fim
    f = FreeFile
    Open arq For Input As #f
    Do While Not EOF(f)
        Line Input #f, linha
        If Left$(linha, 3) = "---" Then Exit Do
        p = InStr(linha, "=")
        If p > 0 Then d(Left$(linha, p - 1)) = Mid$(linha, p + 1)
    Loop
    Close #f
    Set LerKv = d
    Exit Function
Fim:
    On Error Resume Next
    Close #f
    Set LerKv = Nothing
End Function

Public Function Rabo(ByVal arq As String, Optional bytes As Long = 6000) As String
    ' ultimas linhas de um arquivo de log, sem carregar tudo
    Dim f As Integer, tam As Long, s As String
    If Dir(arq) = "" Then Exit Function
    On Error Resume Next
    f = FreeFile
    Open arq For Binary Access Read As #f
    tam = LOF(f)
    If tam = 0 Then Close #f: Exit Function
    If tam > bytes Then Seek #f, tam - bytes + 1: s = Space$(bytes) Else s = Space$(tam)
    Get #f, , s
    Close #f
    Rabo = Replace(Replace(s, vbCrLf, vbLf), vbLf, vbCrLf)
End Function

'==============================================================
' Conclusao: solucao de volta, relatorios
'==============================================================
Private Sub Concluir()
    Dim f As Integer, linha As String, secao As String, p() As String, v As Long, dic As Object
    Set gSol = LerKv(gArqSol)
    gTemSolucao = False
    Set gSensVar = New Collection: Set gSensRestr = New Collection
    If gSol Is Nothing Then
        frmProgresso.Erro "O motor terminou sem gravar a solucao. Log:" & vbLf & Rabo(gArqLog, 1500)
        Exit Sub
    End If
    ReDim gValores(1 To gNV)
    f = FreeFile
    Open gArqSol For Input As #f
    secao = ""
    Do While Not EOF(f)
        Line Input #f, linha
        If Left$(linha, 3) = "---" Then
            secao = Mid$(linha, 4)
            If secao = "" Then secao = "vars"
        ElseIf secao = "vars" Then
            p = Split(linha, " ")
            If UBound(p) >= 1 Then
                If Left$(p(0), 1) = "x" Then
                    v = Val(Mid$(p(0), 2))
                    If v >= 1 And v <= gNV Then gValores(v) = Val(p(1)): gTemSolucao = True
                End If
            End If
        ElseIf secao = "sens_var" Then
            gSensVar.Add Split(linha, " ")
        ElseIf secao = "sens_restr" Then
            gSensRestr.Add Split(linha, " ")
        End If
    Loop
    Close #f
    frmProgresso.Terminou gSol, gTemSolucao, gSensVar.Count > 0
End Sub

' chamado pela janela quando o usuario decide o que fazer com o resultado
Public Sub IMTAplicar(ByVal manter As Boolean, ByVal relResposta As Boolean, ByVal relSens As Boolean)
    Dim v As Long
    On Error GoTo Falhou
    PrepararCalculo
    If manter And gTemSolucao Then
        For v = 1 To gNV: gVarCel(v).Value2 = gValores(v): Next v
    Else
        For v = 1 To gNV: gVarCel(v).Value2 = gVarOriginal(v): Next v
    End If
    gM.ws.Calculate
    RestaurarCalculo
    If relResposta And gTemSolucao Then RelatorioResposta
    If relSens And gSensVar.Count > 0 Then RelatorioSensibilidade
    gM.ws.Activate
    Exit Sub
Falhou:
    RestaurarCalculo
    MsgBox "Erro " & Err.Number & " aplicando a solucao: " & Err.Description, vbCritical, IMT_NOME
End Sub

Private Function NovaAba(ByVal prefixo As String) As Worksheet
    Dim n As Long, nome As String, ws As Worksheet
    n = 1
    Do
        nome = prefixo & " " & n
        Set ws = Nothing
        On Error Resume Next
        Set ws = gM.ws.Parent.Worksheets(nome)
        On Error GoTo 0
        If ws Is Nothing Then Exit Do
        n = n + 1
    Loop
    Set ws = gM.ws.Parent.Worksheets.Add(After:=gM.ws.Parent.Worksheets(gM.ws.Parent.Worksheets.Count))
    ws.Name = nome
    Set NovaAba = ws
End Function

Public Function Rotulo(cel As Range) As String
    ' Como o Solver do Excel: texto mais proximo a esquerda na linha,
    ' mais o texto mais proximo acima na coluna.
    '
    ' Le cada faixa de uma vez so. Celula a celula, um modelo com milhares
    ' de variaveis fazia milhoes de leituras aqui.
    Dim esq As String, cima As String, k As Long, lim As Long
    Dim ws As Worksheet, a As Variant
    Set ws = cel.Worksheet

    lim = cel.Column - ROTULO_COLS
    If lim < 1 Then lim = 1
    If cel.Column > 1 Then
        a = FaixaComoMatriz(ws.Range(ws.Cells(cel.Row, lim), _
                                     ws.Cells(cel.Row, cel.Column - 1)))
        For k = UBound(a, 2) To LBound(a, 2) Step -1
            If VarType(a(LBound(a, 1), k)) = vbString Then
                If Len(a(LBound(a, 1), k)) > 0 Then
                    esq = a(LBound(a, 1), k): Exit For
                End If
            End If
        Next k
    End If

    lim = cel.Row - ROTULO_LINHAS
    If lim < 1 Then lim = 1
    If cel.Row > 1 Then
        a = FaixaComoMatriz(ws.Range(ws.Cells(lim, cel.Column), _
                                     ws.Cells(cel.Row - 1, cel.Column)))
        For k = UBound(a, 1) To LBound(a, 1) Step -1
            If VarType(a(k, LBound(a, 2))) = vbString Then
                If Len(a(k, LBound(a, 2))) > 0 Then
                    cima = a(k, LBound(a, 2)): Exit For
                End If
            End If
        Next k
    End If

    Rotulo = Trim$(esq & " " & cima)
End Function

Private Function FaixaComoMatriz(r As Range) As Variant
    ' Value2 de uma celula unica vem escalar; devolvemos sempre 2D.
    Dim a(1 To 1, 1 To 1) As Variant
    If r.Cells.Count = 1 Then
        a(1, 1) = r.Value2
        FaixaComoMatriz = a
    Else
        FaixaComoMatriz = r.Value2
    End If
End Function

Private Sub Cabecalho(ws As Worksheet, ByVal titulo As String)
    ws.Range("A1").Value = IMT_NOME & " " & IMT_VERSAO & " - " & titulo
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Planilha: " & gM.ws.Parent.Name & " [" & gM.ws.Name & "]"
    ' iif-ok: os dois ramos sao seguros (revisado)
    ws.Range("A3").Value = "Solver: " & gSol("solver") & "   Status: " & StatusBonito(gSol("status")) & _
                           "   Tempo do solver: " & TempoBonito(gSol("segundos")) & _
                           IIf(Len(gSol("gap")) > 0, "   Gap: " & Format$(Val(gSol("gap")), "0.00%"), "")
    ws.Range("A4").Value = "Gerado em " & Format$(Now, "dd/mm/yyyy hh:nn")
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:H").ColumnWidth = 16
End Sub

Private Sub Tabela(ws As Worksheet, ByVal linha As Long, ByVal titulo As String, cab As Variant)
    Dim k As Long
    ws.Cells(linha, 1).Value = titulo
    ws.Cells(linha, 1).Font.Bold = True
    For k = 0 To UBound(cab)
        With ws.Cells(linha + 1, k + 1)
            .Value = cab(k)
            .Font.Bold = True
            .Interior.Color = RGB(68, 84, 106)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
    Next k
End Sub

Public Function TempoBonito(s As Variant) As String
    ' O motor grava os segundos com 4 casas. Mostramos na unidade que
    ' faz sentido, para nao esconder um solve de milissegundos num "0 s".
    Dim v As Double, m As Long
    If IsEmpty(s) Then TempoBonito = "-": Exit Function
    If Len(CStr(s)) = 0 Then TempoBonito = "-": Exit Function
    v = Val(s)                       ' Val usa ponto decimal em qualquer idioma
    If v < 0.001 Then
        TempoBonito = Format$(v * 1000, "0.000") & " ms"
    ElseIf v < 1 Then
        TempoBonito = Format$(v * 1000, "#,##0.0") & " ms"
    ElseIf v < 60 Then
        TempoBonito = Format$(v, "#,##0.00") & " s"
    ElseIf v < 3600 Then
        m = Int(v / 60)
        TempoBonito = m & " min " & Format$(v - 60 * m, "0") & " s"
    Else
        m = Int(v / 3600)
        TempoBonito = m & " h " & Format$((v - 3600# * m) / 60, "0") & " min"
    End If
End Function

Public Function StatusBonito(ByVal s As String) As String
    Select Case s
        Case "Optimal": StatusBonito = "Otimo encontrado"
        Case "TimeLimit": StatusBonito = "Limite de tempo (melhor solucao encontrada)"
        Case "Infeasible": StatusBonito = "Inviavel: nao existe solucao que atenda as restricoes"
        Case "Unbounded": StatusBonito = "Ilimitado: o objetivo pode melhorar sem fim"
        Case "Cancelled": StatusBonito = "Cancelado pelo usuario"
        Case "Erro": StatusBonito = "Erro"
        Case Else: StatusBonito = s
    End Select
End Function

Private Function Milhar(v As Variant) As Variant
    If Val(v) >= 1E+29 Then Milhar = "1E+30" Else Milhar = Val(v)
End Function

Private Sub RelatorioResposta()
    Dim ws As Worksheet, L As Long, v As Long, j As Long, i As Long, k As Long
    Dim lhsV As Double, rhsV As Double, folga As Double, cel As Range
    Set ws = NovaAba("Resposta")
    Cabecalho ws, "Relatorio de Resposta"

    L = 6
    If Not gM.obj Is Nothing Then
        ' iif-ok: os dois ramos sao seguros (revisado)
        Tabela ws, L, "Celula do Objetivo (" & IIf(gM.tipo = TIPO_MAX, "Max", IIf(gM.tipo = TIPO_MIN, "Min", "Valor de")) & ")", _
               Array("Celula", "Nome", "Valor Original", "Valor Final")
        ws.Cells(L + 2, 1).Value = gM.obj.Address(False, False)
        ws.Cells(L + 2, 2).Value = Rotulo(gM.obj)
        ws.Cells(L + 2, 3).Value = gObjOriginal
        ws.Cells(L + 2, 4).Value = gM.obj.Value2
        L = L + 4
    End If

    ' Tabelas montadas em memoria e gravadas de uma vez. Celula a celula,
    ' o modelo grande fazia ~90 mil escritas separadas na planilha.
    Dim dados As Variant
    Tabela ws, L, "Celulas Variaveis", Array("Celula", "Nome", "Valor Original", "Valor Final", "Tipo")
    ReDim dados(1 To gNV, 1 To 5)
    For v = 1 To gNV
        dados(v, 1) = gVarEnd(v)
        dados(v, 2) = Rotulo(gVarCel(v))
        dados(v, 3) = gVarOriginal(v)
        dados(v, 4) = gValores(v)
        dados(v, 5) = TipoDaVar(v)
    Next v
    ws.Range(ws.Cells(L + 2, 1), ws.Cells(L + 1 + gNV, 5)).Value2 = dados
    L = L + 3 + gNV

    Tabela ws, L, "Restricoes", Array("Celula", "Nome", "Valor da Celula", "Formula", "Status", "Folga")
    ReDim dados(1 To IIf(gNR = 0, 1, gNR), 1 To 6)
    For j = 1 To gNR
        i = gRowBloco(j)
        Set cel = gM.r(i).lhs.Cells(gRowCel(j))
        lhsV = Val0(cel.Value2)
        If gM.r(i).rhsEhConst Then
            rhsV = gM.r(i).rhsConst
        ElseIf gM.r(i).rhs.Cells.Count = 1 Then
            rhsV = Val0(gM.r(i).rhs.Value2)
        Else
            rhsV = Val0(gM.r(i).rhs.Cells(gRowCel(j)).Value2)
        End If
        folga = Abs(lhsV - rhsV)
        dados(j, 1) = gRowEnd(j)
        dados(j, 2) = Rotulo(cel)
        dados(j, 3) = lhsV
        dados(j, 4) = gRowEnd(j) & Sinal(gRowRel(j)) & TextoLadoDireito(i, gRowCel(j))
        dados(j, 5) = IIf(folga < 0.000001, "Ativa", "Nao ativa")
        dados(j, 6) = folga
    Next j
    If gNR > 0 Then
        ws.Range(ws.Cells(L + 2, 1), ws.Cells(L + 1 + gNR, 6)).Value2 = dados
    End If
    ws.Activate
    ws.Range("A1").Select
End Sub

Private Function TextoLadoDireito(ByVal i As Long, ByVal k As Long) As String
    ' Com IIf isto quebrava: o VBA avalia TODOS os ramos, e quando o lado
    ' direito e uma constante o objeto rhs vale Nothing.
    If gM.r(i).rhsEhConst Then
        TextoLadoDireito = Num(gM.r(i).rhsConst)
    ElseIf gM.r(i).rhs.Cells.Count = 1 Then
        TextoLadoDireito = gM.r(i).rhs.Address(False, False)
    Else
        TextoLadoDireito = gM.r(i).rhs.Cells(k).Address(False, False)
    End If
End Function

Private Function Sinal(ByVal rel As Long) As String
    Select Case rel
        Case REL_LE: Sinal = "<="
        Case REL_GE: Sinal = ">="
        Case Else: Sinal = "="
    End Select
End Function

Private Sub MontarTipos()
    ' Uma passada so por todas as faixas int/bin. Antes era busca linear
    ' por variavel: com 5.724 variaveis e 4.464 celulas declaradas davam
    ' 25 milhoes de comparacoes de endereco.
    Dim i As Long, cel As Range
    Set gTipos = CreateObject("Scripting.Dictionary")
    For i = 1 To gM.n
        If gM.r(i).rel = REL_INT Or gM.r(i).rel = REL_BIN Then
            For Each cel In gM.r(i).lhs.Cells
                gTipos(cel.Address(False, False)) = _
                    IIf(gM.r(i).rel = REL_BIN, "Binaria", "Inteira")
            Next cel
        End If
    Next i
End Sub

Private Function TipoDaVar(ByVal v As Long) As String
    If gTipos Is Nothing Then MontarTipos
    If gTipos.Exists(gVarEnd(v)) Then
        TipoDaVar = gTipos(gVarEnd(v))
    Else
        TipoDaVar = "Continua"
    End If
End Function

Private Sub RelatorioSensibilidade()
    Dim ws As Worksheet, L As Long, k As Long, p As Variant, v As Long, j As Long
    Set ws = NovaAba("Sensibilidade")
    Cabecalho ws, "Relatorio de Sensibilidade"
    If gSol("sensibilidade") = "fixando_inteiras" Then
        ws.Range("A5").Value = "O modelo tem variaveis inteiras. A analise abaixo e da programacao linear que sobra " & _
                               "quando cada inteira e fixada no valor da solucao; as inteiras nao aparecem."
        ws.Range("A5").Font.Italic = True
    End If

    L = 7
    Tabela ws, L, "Celulas Variaveis", Array("Celula", "Nome", "Valor Final", "Custo Reduzido", _
                                              "Coeficiente Objetivo", "Aumento Permitido", "Reducao Permitida")
    Dim dados As Variant
    ' iif-ok: os dois ramos sao seguros (revisado)
    ReDim dados(1 To IIf(gSensVar.Count = 0, 1, gSensVar.Count), 1 To 7)
    k = 0
    For Each p In gSensVar
        v = Val(Mid$(p(0), 2))
        If v >= 1 And v <= gNV And UBound(p) >= 5 Then
            k = k + 1
            dados(k, 1) = gVarEnd(v)
            dados(k, 2) = Rotulo(gVarCel(v))
            dados(k, 3) = Val(p(1))
            dados(k, 4) = Val(p(2))
            dados(k, 5) = Val(p(3))
            dados(k, 6) = Milhar(p(4))
            dados(k, 7) = Milhar(p(5))
        End If
    Next p
    If k > 0 Then ws.Range(ws.Cells(L + 2, 1), ws.Cells(L + 1 + k, 7)).Value2 = dados
    L = L + 3 + k

    Tabela ws, L, "Restricoes", Array("Celula", "Nome", "Valor Final", "Preco Sombra", _
                                        "Lado Direito", "Aumento Permitido", "Reducao Permitida")
    ' iif-ok: os dois ramos sao seguros (revisado)
    ReDim dados(1 To IIf(gSensRestr.Count = 0, 1, gSensRestr.Count), 1 To 7)
    k = 0
    For Each p In gSensRestr
        If Left$(p(0), 1) = "c" And UBound(p) >= 5 Then
            j = Val(Mid$(p(0), 2))
            If j >= 1 And j <= gNR Then
                k = k + 1
                dados(k, 1) = gRowEnd(j)
                dados(k, 2) = Rotulo(gM.r(gRowBloco(j)).lhs.Cells(gRowCel(j)))
                ' no .lp a linha e (lhs - rhs) menos a constante; aqui volta ao valor da celula
                dados(k, 3) = Val(p(1)) + gRow0(j) + LadoDireito(j)
                dados(k, 4) = Val(p(2))
                dados(k, 5) = LadoDireito(j)
                dados(k, 6) = Milhar(p(4))
                dados(k, 7) = Milhar(p(5))
            End If
        End If
    Next p
    If k > 0 Then ws.Range(ws.Cells(L + 2, 1), ws.Cells(L + 1 + k, 7)).Value2 = dados
    ws.Range("A" & (L + 3 + k)).Value = "Preco sombra: quanto o objetivo muda por unidade a mais no lado direito. " & _
        "Aumento/reducao permitidos: ate onde o lado direito pode ir sem que o preco sombra mude."
    ws.Range("A" & (L + 3 + k)).Font.Italic = True
    ws.Activate
    ws.Range("A1").Select
End Sub

Private Function LadoDireito(ByVal j As Long) As Double
    Dim i As Long
    i = gRowBloco(j)
    If gM.r(i).rhsEhConst Then
        LadoDireito = gM.r(i).rhsConst
    ElseIf gM.r(i).rhs.Cells.Count = 1 Then
        LadoDireito = Val0(gM.r(i).rhs.Value2)
    Else
        LadoDireito = Val0(gM.r(i).rhs.Cells(gRowCel(j)).Value2)
    End If
    ' o lado direito de "lhs - rhs rel 0" no .lp foi -row0; aqui e o da planilha
End Function

'==============================================================
' Verificar modelo
'
' Aponta o que parece errado, com o endereco da celula, e pinta as
' suspeitas na planilha. NUNCA impede de resolver: e opiniao, nao
' porteiro. O Resolver tem as travas dele, para o que impede de
' montar o modelo; aqui a conversa e mais ampla.
'
' Duas fases:
'   estrutura  - instantanea, so olha como o modelo foi declarado
'   conteudo   - le os coeficientes (a mesma leitura do Resolver) e
'                encontra restricao vazia, variavel orfa, escala ruim
'==============================================================

Private Sub Bloqueia(ByVal onde As String, ByVal oque As String, _
                     ByVal faca As String)
    ' erro que impede a propria leitura do modelo
    gBloqueia = True
    Achado "Erro", onde, oque, faca
End Sub

Private Sub Achado(ByVal grav As String, ByVal onde As String, _
                   ByVal oque As String, ByVal faca As String)
    ' Uma celula pode cair em mais de uma faixa do modelo, e o mesmo
    ' achado saia repetido. Guardamos o que ja foi dito.
    Dim chave As String
    chave = onde & "|" & oque
    If gAchJaDito Is Nothing Then Set gAchJaDito = CreateObject("Scripting.Dictionary")
    If gAchJaDito.Exists(chave) Then Exit Sub
    gAchJaDito(chave) = True
    gNAch = gNAch + 1
    ReDim Preserve gAchGrav(1 To gNAch): ReDim Preserve gAchOnde(1 To gNAch)
    ReDim Preserve gAchOque(1 To gNAch): ReDim Preserve gAchFaca(1 To gNAch)
    gAchGrav(gNAch) = grav: gAchOnde(gNAch) = onde
    gAchOque(gNAch) = oque: gAchFaca(gNAch) = faca
    If Len(onde) > 0 Then gAchPintar(onde) = True
End Sub

Public Sub IMTVerificar(Optional ws As Worksheet = Nothing)
    Dim m As Modelo, i As Long, k As Long, v As Long, j As Long
    Dim cel As Range, area As Range, dic As Object, hf As Variant
    Dim n As Long, resumo As String, tudo As Range

    If ws Is Nothing Then Set ws = ActiveSheet
    If Not LerModelo(ws, m) Then
        MsgBox "Esta aba nao tem modelo salvo." & vbLf & vbLf & _
               OndeHaModelo(ws) & vbLf & vbLf & _
               "Para criar um modelo aqui: IMTSolver > Modelo, defina objetivo, " & _
               "variaveis e restricoes, e clique em Salvar.", _
               vbInformation, IMT_NOME
        Exit Sub
    End If

    gNAch = 0: gBloqueia = False: gVerificando = True
    Set gAchJaDito = CreateObject("Scripting.Dictionary")
    Set gAchPintar = CreateObject("Scripting.Dictionary")
    Set dic = CreateObject("Scripting.Dictionary")
    On Error GoTo Falhou

    '--- objetivo ---
    If m.obj Is Nothing Then
        Achado "Aviso", "", "Sem celula de objetivo.", _
               "Sem objetivo ele so procura uma solucao viavel. " & _
               "Se era isso que voce queria, ignore."
    ElseIf m.obj.Cells.Count > 1 Then
        Achado "Erro", m.obj.Address(False, False), _
               "O objetivo aponta para " & m.obj.Cells.Count & " celulas.", _
               "O objetivo tem de ser UMA celula so."
    ElseIf Not m.obj.HasFormula Then
        Achado "Erro", m.obj.Address(False, False), _
               "A celula do objetivo nao tem formula.", _
               "Ela deveria calcular o total a otimizar, algo como " & _
               "=SOMARPRODUTO(custos; variaveis)."
    End If

    '--- variaveis ---
    n = m.vars.Cells.Count
    For Each area In m.vars.Areas
        hf = area.HasFormula
        If IsNull(hf) Or hf Then
            For Each cel In area.Cells
                If cel.HasFormula Then
                    Bloqueia cel.Address(False, False), _
                           "Variavel de decisao com formula.", _
                           "Celula de decisao guarda um valor; quem calcula " & _
                           "e o solver. Apague a formula."
                End If
            Next cel
        End If
        For Each cel In area.Cells
            dic(cel.Address(False, False)) = True
        Next cel
    Next area
    If n > 0 Then
        If Not m.obj Is Nothing Then
            If Not Application.Intersect(m.vars, m.obj) Is Nothing Then
                Achado "Erro", m.obj.Address(False, False), _
                       "A celula do objetivo esta dentro da faixa de variaveis.", _
                       "O objetivo e calculado A PARTIR das variaveis; nao " & _
                       "pode ser uma delas."
            End If
        End If
    End If

    '--- restricoes ---
    Dim visto As Object, ender As String
    Set visto = CreateObject("Scripting.Dictionary")
    For i = 1 To m.n
        If m.r(i).lhs Is Nothing Then
            Achado "Erro", "", "Restricao " & i & " tem lado esquerdo invalido.", _
                   "Abra a janela Modelo e refaca essa linha."
            GoTo Proxima
        End If
        ender = m.r(i).lhs.Address(False, False)
        If visto.Exists(ender) Then
            Achado "Aviso", ender, "Restricao repetida: " & ender & _
                   " ja aparece na linha " & visto(ender) & ".", _
                   "Nao muda o resultado, so pesa. Apague a duplicada."
        Else
            visto(ender) = i
        End If

        If m.r(i).rel <= REL_GE Then
            ' lado esquerdo tem de ser formula
            hf = m.r(i).lhs.HasFormula
            If IsNull(hf) Or (hf = False) Then
                For Each cel In m.r(i).lhs.Cells
                    If Not cel.HasFormula Then
                        Achado "Erro", cel.Address(False, False), _
                               "Lado esquerdo sem formula: e um numero fixo.", _
                               "Uma restricao compara uma FORMULA das variaveis " & _
                               "com um limite. Assim ela nao restringe nada."
                        If gNAch > 40 Then Exit For
                    End If
                Next cel
            End If
            ' tamanho do lado direito
            If Not m.r(i).rhsEhConst Then
                If m.r(i).rhs.Cells.Count <> 1 And _
                   m.r(i).rhs.Cells.Count <> m.r(i).lhs.Cells.Count Then
                    Bloqueia m.r(i).rhs.Address(False, False), _
                           "Lado direito com " & m.r(i).rhs.Cells.Count & _
                           " celulas contra " & m.r(i).lhs.Cells.Count & _
                           " do esquerdo.", _
                           "Use uma celula so (vale para todas) ou o mesmo " & _
                           "tamanho do lado esquerdo."
                End If
                If Not Application.Intersect(m.r(i).rhs, m.vars) Is Nothing Then
                    Achado "Aviso", m.r(i).rhs.Address(False, False), _
                           "O lado direito invade a faixa de variaveis.", _
                           "Costuma ser engano ao escolher a faixa."
                End If
            End If
            If Not Application.Intersect(m.r(i).lhs, m.vars) Is Nothing Then
                Achado "Aviso", m.r(i).lhs.Address(False, False), _
                       "O lado esquerdo invade a faixa de variaveis.", _
                       "A celula seria variavel e restricao ao mesmo tempo."
            End If
        Else
            ' int / bin fora da faixa de variaveis nao tem efeito
            For Each cel In m.r(i).lhs.Cells
                If Not dic.Exists(cel.Address(False, False)) Then
                    Dim marca As String
                    If m.r(i).rel = REL_INT Then marca = "int" Else marca = "bin"
                    Achado "Erro", cel.Address(False, False), _
                           "Declarada " & marca & _
                           " mas nao e uma variavel do modelo.", _
                           "int e bin so valem para celulas que estao na " & _
                           "faixa de variaveis. Fora dela, sao ignoradas."
                    If gNAch > 40 Then Exit For
                End If
            Next cel
        End If
Proxima:
    Next i

    '--- celulas com erro em qualquer faixa do modelo ---
    Set tudo = m.vars
    If Not m.obj Is Nothing Then Set tudo = Application.Union(tudo, m.obj)
    For i = 1 To m.n
        If Not m.r(i).lhs Is Nothing Then Set tudo = Application.Union(tudo, m.r(i).lhs)
        If m.r(i).rel <= REL_GE And Not m.r(i).rhsEhConst Then
            Set tudo = Application.Union(tudo, m.r(i).rhs)
        End If
    Next i
    ProcurarErros tudo

    '--- fase 2: os coeficientes ---
    ' Le os coeficientes sempre. Perguntar antes so acrescentava um
    ' clique: a leitura e rapida, nao muda a planilha, e e dela que saem
    ' os achados que importam.
    Dim leu As Boolean
    If gBloqueia Then
        Achado "Nota", "", "A leitura dos coeficientes foi pulada.", _
               "Conserte os erros acima e verifique de novo: ai eu procuro " & _
               "restricao sem variavel, variavel sem uso e problema de escala."
    Else
        leu = VerificarCoeficientes(m, resumo)
    End If

    gVerificando = False
    MostrarAchados m, leu, resumo
    Exit Sub

Falhou:
    gVerificando = False
    MsgBox "Erro " & Err.Number & " ao verificar: " & Err.Description, _
           vbCritical, IMT_NOME
End Sub

Private Sub ProcurarErros(faixa As Range)
    ' Uma celula com #VALOR! ou #REF! dentro do modelo derruba a leitura.
    ' Le em bloco: celula a celula, num modelo grande, demoraria.
    Dim a As Range, mat As Variant, r As Long, c As Long, cel As Range
    Dim achou As Long
    On Error Resume Next
    For Each a In faixa.Areas
        Set cel = a.SpecialCells(xlCellTypeFormulas, 16)   ' 16 = so erros
        If Not cel Is Nothing Then
            Dim e As Range
            For Each e In cel.Cells
                achou = achou + 1
                If achou <= 10 Then
                    Bloqueia e.Address(False, False), _
                           "Celula com erro: " & CStr(e.Text), _
                           "Conserte a formula. Com erro aqui, o modelo nao " & _
                           "pode ser lido."
                End If
            Next e
        End If
        Set cel = Nothing
    Next a
    If achou > 10 Then
        Achado "Erro", "", "e mais " & (achou - 10) & " celulas com erro.", _
               "Conserte as formulas antes de resolver."
    End If
End Sub

Private Function VerificarCoeficientes(m As Modelo, ByRef resumo As String) As Boolean
    ' Reaproveita a leitura do Resolver e olha a matriz que sai dela.
    Dim k As Long, v As Long, j As Long
    Dim usadaEmRestr() As Boolean, temNaLinha() As Long
    Dim menor As Double, maior As Double, x As Double
    Dim orfas As Long, vazias As Long

    gM = m
    Set gTipos = Nothing
    frmProgresso.Abrir "", "Verificando o modelo"
    frmProgresso.Fase "Lendo o modelo para verificar"
    If Not Extrair() Then
        frmProgresso.Hide
        Exit Function
    End If
    frmProgresso.Hide

    ReDim usadaEmRestr(1 To gNV)
    ReDim temNaLinha(1 To IIf(gNR = 0, 1, gNR))
    menor = 1E+300: maior = 0
    For k = 1 To gNz
        usadaEmRestr(gNzVar(k)) = True
        temNaLinha(gNzRow(k)) = temNaLinha(gNzRow(k)) + 1
        x = Abs(gNzVal(k))
        If x > 0 Then
            If x < menor Then menor = x
            If x > maior Then maior = x
        End If
    Next k

    ' restricao que nao depende de variavel nenhuma
    For j = 1 To gNR
        If temNaLinha(j) = 0 Then
            vazias = vazias + 1
            If vazias <= 15 Then
                Achado "Erro", gRowEnd(j), _
                       "Esta restricao nao depende de nenhuma variavel.", _
                       "O valor dela nunca muda. Ou nao restringe nada, ou " & _
                       "torna o modelo inviavel sozinha - e ai voce procuraria " & _
                       "o motivo por horas."
            End If
        End If
    Next j
    If vazias > 15 Then
        Achado "Erro", "", "e mais " & (vazias - 15) & " restricoes sem variavel.", _
               "Confira se a faixa do lado esquerdo esta certa."
    End If

    ' variavel que nao aparece em lugar nenhum
    For v = 1 To gNV
        If Not usadaEmRestr(v) Then
            If gCObj(v) = 0 Then
                orfas = orfas + 1
                If orfas <= 15 Then
                    Achado "Aviso", gVarEnd(v), _
                           "Esta variavel nao aparece no objetivo nem em " & _
                           "restricao alguma.", _
                           "Ela nao faz nada: o solver pode devolver qualquer " & _
                           "valor. Ou faltou usa-la, ou ela nao devia estar " & _
                           "na faixa."
                End If
            End If
        End If
    Next v
    If orfas > 15 Then
        Achado "Aviso", "", "e mais " & (orfas - 15) & " variaveis sem uso.", _
               "Reveja a faixa de variaveis."
    End If

    ' objetivo que nao depende de variavel
    If Not gM.obj Is Nothing Then
        Dim temObj As Boolean
        For v = 1 To gNV
            If gCObj(v) <> 0 Then temObj = True: Exit For
        Next v
        If Not temObj Then
            Achado "Aviso", gM.obj.Address(False, False), _
                   "O objetivo nao depende de nenhuma variavel.", _
                   "Ele vai valer o mesmo em qualquer solucao. O solver so " & _
                   "vai procurar uma resposta viavel."
        End If
    End If

    ' escala dos coeficientes
    If gNz > 0 And menor < 1E+299 Then
        If maior / menor > 1000000# Then
            Achado "Aviso", "", _
                   "Coeficientes em escalas muito diferentes: de " & _
                   Format$(menor, "0.0####E+00") & " a " & _
                   Format$(maior, "0.0####E+00") & " (razao " & _
                   Format$(maior / menor, "#,##0") & ").", _
                   "O solver trabalha com 15 digitos; essa faixa gasta boa " & _
                   "parte deles. Deixa a relaxacao fraca (demora a fechar) e " & _
                   "arrisca aceitar como viavel algo que nao e. Se houver " & _
                   "Big M, use o menor valor que ainda funcione."
        End If
    End If

    Dim nInt As Long, nBin As Long
    MontarTipos
    For v = 1 To gNV
        If gTipos.Exists(gVarEnd(v)) Then
            If gTipos(gVarEnd(v)) = "Binaria" Then nBin = nBin + 1 Else nInt = nInt + 1
        End If
    Next v
    resumo = Format$(gNV, "#,##0") & " variaveis (" & Format$(nInt, "#,##0") & _
             " inteiras, " & Format$(nBin, "#,##0") & " binarias)" & vbLf & _
             Format$(gNR, "#,##0") & " restricoes, " & Format$(gNz, "#,##0") & _
             " coeficientes"
    If gNV > 0 And gNR > 0 Then
        resumo = resumo & ", densidade " & _
                 Format$(gNz / (CDbl(gNV) * gNR), "0.000%")
    End If
    If gNz > 0 And menor < 1E+299 Then
        resumo = resumo & vbLf & "coeficientes de " & _
                 Format$(menor, "0.0####E+00") & " a " & Format$(maior, "0.0####E+00")
    End If
    VerificarCoeficientes = True
End Function

Private Sub MostrarAchados(m As Modelo, ByVal leu As Boolean, ByVal resumo As String)
    Dim i As Long, nErro As Long, nAviso As Long, ws As Worksheet
    Dim dados As Variant, msg As String

    Dim nNota As Long
    For i = 1 To gNAch
        Select Case gAchGrav(i)
            Case "Erro": nErro = nErro + 1
            Case "Aviso": nAviso = nAviso + 1
            Case Else: nNota = nNota + 1
        End Select
    Next i

    If gNAch = 0 Then
        msg = "Nenhum problema encontrado."
        If Len(resumo) > 0 Then msg = msg & vbLf & vbLf & resumo
        If Not leu Then msg = msg & vbLf & vbLf & _
            "(so a estrutura foi conferida; os coeficientes nao)"
        MsgBox msg, vbInformation, IMT_NOME & " - verificacao"
        Exit Sub
    End If

    Set ws = NovaAbaEm(m.ws.Parent, "Verificacao")
    ws.Range("A1").Value = IMT_NOME & " " & IMT_VERSAO & " - verificacao do modelo"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Aba: " & m.ws.Name & "     " & _
                           nErro & " erro(s), " & nAviso & " aviso(s)"
    ws.Range("A3").Value = "Gerado em " & Format$(Now, "dd/mm/yyyy hh:nn")
    If Len(resumo) > 0 Then ws.Range("A4").Value = Replace(resumo, vbLf, "     ")
    ws.Range("A5").Value = "Isto e um parecer: nada aqui impede voce de resolver o modelo. " & _
        "As celulas citadas foram marcadas de vermelho na aba do modelo - o botao Esconder tira as marcas."
    ws.Range("A5").Font.Italic = True

    Dim cab As Variant
    cab = Array("Gravidade", "Celula", "O que parece errado", "O que fazer")
    For i = 0 To 3
        With ws.Cells(7, i + 1)
            .Value = cab(i)
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 84, 106)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    ReDim dados(1 To gNAch, 1 To 4)
    For i = 1 To gNAch
        dados(i, 1) = gAchGrav(i)
        dados(i, 2) = gAchOnde(i)
        dados(i, 3) = gAchOque(i)
        dados(i, 4) = gAchFaca(i)
    Next i
    ws.Range(ws.Cells(8, 1), ws.Cells(7 + gNAch, 4)).Value2 = dados
    For i = 1 To gNAch
        If gAchGrav(i) = "Erro" Then
            ws.Cells(7 + i, 1).Font.Color = RGB(192, 0, 0)
        Else
            ws.Cells(7 + i, 1).Font.Color = RGB(191, 143, 0)
        End If
    Next i
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 52
    ws.Columns("D").ColumnWidth = 62
    ws.Range("C:D").WrapText = True
    ws.Rows("8:" & (7 + gNAch)).VerticalAlignment = xlTop
    ws.Activate
    ws.Range("A1").Select

    ' marcar e reversivel (o botao Esconder limpa), entao nao perguntamos
    If gAchPintar.Count > 0 Then PintarAchados m.ws
End Sub

Private Sub PintarAchados(ws As Worksheet)
    Dim ch As Variant, r As Range
    On Error Resume Next
    For Each ch In gAchPintar.Keys
        Set r = Nothing
        Set r = ws.Range(CStr(ch))
        If Not r Is Nothing Then Moldura ws, r, RGB(200, 0, 0), ""
    Next ch
End Sub

Private Function NovaAbaEm(wb As Workbook, ByVal prefixo As String) As Worksheet
    Dim n As Long, nome As String, ws As Worksheet
    n = 1
    Do
        nome = prefixo & " " & n
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(nome)
        On Error GoTo 0
        If ws Is Nothing Then Exit Do
        n = n + 1
    Loop
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = nome
    Set NovaAbaEm = ws
End Function

'==============================================================
' Mostrar o modelo na planilha (molduras, sem mexer nas celulas)
'==============================================================
Public Sub IMTMostrar(Optional ws As Worksheet = Nothing)
    Dim m As Modelo, i As Long, sh As Shape, rotulo As String
    If ws Is Nothing Then Set ws = ActiveSheet
    IMTEsconder ws
    If Not LerModelo(ws, m) Then
        MsgBox "Esta aba nao tem modelo." & vbLf & vbLf & OndeHaModelo(ws), _
               vbInformation, IMT_NOME
        Exit Sub
    End If
    ' iif-ok: os dois ramos sao seguros (revisado)
    If Not m.obj Is Nothing Then Moldura ws, m.obj, RGB(0, 150, 60), IIf(m.tipo = TIPO_MAX, "max", IIf(m.tipo = TIPO_MIN, "min", "= " & m.valorAlvo))
    Dim a As Range
    For Each a In m.vars.Areas
        Moldura ws, a, RGB(0, 90, 200), ""
    Next a
    For i = 1 To m.n
        Select Case m.r(i).rel
            Case REL_INT: Moldura ws, m.r(i).lhs, RGB(120, 60, 180), "int"
            Case REL_BIN: Moldura ws, m.r(i).lhs, RGB(120, 60, 180), "bin"
            Case Else
                If m.r(i).rhsEhConst Then
                    rotulo = " " & Num(m.r(i).rhsConst)
                Else
                    rotulo = " " & m.r(i).rhs.Address(False, False)
                End If
                Moldura ws, m.r(i).lhs, RGB(230, 120, 0), Sinal(m.r(i).rel) & rotulo
                If Not m.r(i).rhsEhConst Then Moldura ws, m.r(i).rhs, RGB(230, 120, 0), ""
        End Select
    Next i
End Sub

Private Sub Moldura(ws As Worksheet, r As Range, ByVal cor As Long, ByVal texto As String)
    Dim sh As Shape, a As Range, t As Shape
    For Each a In r.Areas
        Set sh = ws.Shapes.AddShape(msoShapeRectangle, a.Left, a.Top, a.Width, a.Height)
        sh.Name = "imt_" & Replace(a.Address(False, False), ":", "_") & "_" & ws.Shapes.Count
        sh.Fill.Visible = msoFalse
        sh.Line.ForeColor.RGB = cor
        sh.Line.Weight = 2
        sh.Placement = xlMove
        If texto <> "" Then
            Set t = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, a.Left + a.Width + 2, a.Top, 60, 14)
            t.Name = "imt_txt_" & ws.Shapes.Count
            t.TextFrame.Characters.Text = texto
            t.TextFrame.Characters.Font.Size = 8
            t.TextFrame.Characters.Font.Bold = True
            t.TextFrame.Characters.Font.Color = cor
            t.TextFrame.MarginLeft = 1: t.TextFrame.MarginTop = 0
            t.TextFrame.AutoSize = True
            t.Fill.Visible = msoFalse: t.Line.Visible = msoFalse
        End If
    Next a
End Sub

Public Sub IMTEsconder(Optional ws As Worksheet = Nothing)
    Dim k As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    For k = ws.Shapes.Count To 1 Step -1
        ' "bye_" e o prefixo antigo: molduras desenhadas antes da troca de nome
        If Left$(ws.Shapes(k).Name, 4) = "imt_" Or _
           Left$(ws.Shapes(k).Name, 4) = "bye_" Then ws.Shapes(k).Delete
    Next k
End Sub

Public Sub IMTLimparSolucao(Optional ws As Worksheet = Nothing)
    Dim m As Modelo
    If ws Is Nothing Then Set ws = ActiveSheet
    If Not LerModelo(ws, m) Then
        ' antes saia calado: o botao parecia nao funcionar
        MsgBox "Esta aba nao tem modelo, entao nao ha variaveis para zerar." & _
               vbLf & vbLf & OndeHaModelo(ws), vbInformation, IMT_NOME
        Exit Sub
    End If
    If MsgBox("Zerar as celulas de variavel em " & m.vars.Address(False, False) & "?", _
              vbQuestion + vbYesNo, IMT_NOME) = vbYes Then m.vars.Value2 = 0
End Sub
