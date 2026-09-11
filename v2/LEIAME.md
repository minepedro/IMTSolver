# IMTSolver 2.0

Otimização linear e inteira no Excel, com o solver que você escolher.

> Criado por **Pedro da Silva Bezerra**, aluno do 3º período do
> **Instituto Mauá de Tecnologia**, para a disciplina de
> **Pesquisa Operacional I**, com o incentivo da **professora Joyce**.

- **Janela de modelo** sem ActiveX: funciona em qualquer Excel recente
- **Modelo guardado nos nomes `solver_*`** — o mesmo formato do Solver do
  Excel e do OpenSolver. Um modelo feito lá abre aqui, e vice-versa
- **Execução assíncrona**: janela de progresso com fase, cronômetro, melhor
  solução / limite / gap ao vivo, log do solver e botão Cancelar. O Excel
  nunca congela
- **Solvers**: CBC, HiGHS e SCIP na sua máquina (vêm no pacote); Gurobi,
  CPLEX e GLPK se você tiver licença; **CPLEX, COPT e SCIP de graça pelo
  NEOS** (só precisa informar um e-mail). O botão *Solvers* diz, de cada
  um, se resolve linear, inteiro ou não-linear
- **Relatórios de resposta e de sensibilidade** no formato do Solver do
  Excel: preço-sombra, custo reduzido, aumento e redução permitidos
- **Teste de linearidade** automático: mede o modelo com pesos aleatórios
  e avisa se ele tem `SE`, `MÁXIMO` ou variável vezes variável — coisas
  que nenhum solver de PL resolve
- **Verificar modelo**: aponta as células que parecem erradas e explica o
  porquê, sem impedir você de resolver
- Sem limite de variáveis

---

## Instalar (usuário final)

1. Descompacte a pasta.
2. Feche o Excel.
3. Duplo clique em **`Instalar.bat`**.

Ele copia o suplemento e o motor para a pasta de suplementos do Excel,
tira o bloqueio "arquivo da internet" e registra o suplemento. Não pede
administrador. Ao abrir o Excel, a aba **IMTSolver** aparece.

Para remover: `Desinstalar.bat`.

---

## Usar

1. Monte o modelo numa aba: **células de valor** para as decisões, uma
   **fórmula** para o objetivo, e fórmulas para o lado esquerdo de cada
   restrição.
2. **IMTSolver → Modelo**: aponte o objetivo, as variáveis e as
   restrições. Os botões `...` abrem o seletor de faixa do Excel.
   *Sugerir pelo objetivo* preenche as variáveis com as células de valor
   das quais o objetivo depende.
3. Escolha o solver, o tempo máximo e o gap. Para os do NEOS, informe um
   e-mail.
4. **Resolver**. Acompanhe na janela. Ao terminar, escolha manter ou
   restaurar os valores e quais relatórios gerar.

**Um modelo por aba.** O modelo fica gravado na própria aba, nos nomes
`solver_*` — o mesmo lugar onde o Solver do Excel guarda o dele. Cada aba
de uma pasta pode ter o seu, e todos os botões (Resolver, Verificar,
Mostrar, Zerar variáveis) agem sobre **a aba que está aberta**. Se
aparecer "esta aba não tem modelo", você provavelmente está na aba
errada: a própria mensagem diz em quais abas da pasta há modelo.

**Verificar** na faixa procura problemas no modelo desta aba: variável com
fórmula, restrição que não depende de variável nenhuma, célula com erro,
coeficientes em escalas muito diferentes. Ele escreve uma aba de
verificação e pinta de vermelho as células citadas — *Esconder* tira as
marcas. É um parecer: nada ali impede de resolver.

**Exemplos** na faixa abre uma planilha com três modelos prontos (inteiro,
mix de produção, transporte) e a resposta esperada em cada um.

### Sinais das restrições

| sinal | significado |
|---|---|
| `<=` `=` `>=` | lado esquerdo (faixa de fórmulas) contra lado direito (faixa ou número) |
| `int` | as células do lado esquerdo são inteiras |
| `bin` | as células do lado esquerdo são 0 ou 1 |

Se o lado direito for uma faixa, ela precisa ter uma célula (vale para
todas) ou o mesmo tamanho do lado esquerdo. O lado direito pode ter
fórmula que dependa das variáveis — o modelo trata `lhs − rhs`.

---

## Como funciona por dentro

**Leitura do modelo.** Por perturbação: zera todas as variáveis e lê as
fórmulas (constante), depois liga uma variável por vez em 1 e mede o
quanto cada fórmula mudou (coeficiente). Em seguida testa num segundo
ponto (variáveis em 1, 2, 3) se `constante + Σ coef·x` reproduz as
fórmulas; se não reproduz, o modelo não é linear e ele avisa.

**Motor.** O VBA escreve um `.lp` padrão e chama `imtsolver.exe`
escondido. O motor grava um arquivo de progresso a cada meio segundo e
obedece a um arquivo de cancelamento. O Excel lê o progresso uma vez por
segundo com `Application.OnTime`, sem travar.

O motor é uma **pasta**, não um arquivo único. Empacotado como arquivo
único ele se descompactava inteiro na memória a cada chamada — 1,42 s
antes da primeira instrução rodar. Em pasta, começa em 0,1 s. Zipado dá
no mesmo (79 MB); só ocupa mais espaço depois de instalado.

A lista de solvers fica guardada entre as sessões, porque só muda quando
alguém instala ou remove um. O botão **Atualizar lista de solvers**, na
janela do modelo, refaz a consulta.

Os solvers do NEOS aparecem na lista sem consulta à internet — são um
serviço remoto, não algo instalado aqui. Se a rede estiver fora, o erro
aparece na hora de resolver, dizendo isso.

**NEOS.** Pela API XML-RPC oficial (`neos-server.org:3333`). O motor pede o
modelo de submissão ao próprio servidor e preenche os campos, então não
quebra quando o NEOS muda o formulário — foi isso que derrubou o
OpenSolver.

**Sensibilidade.** Sempre pelo HiGHS (duais e *ranging*), mesmo que o
solver escolhido tenha sido outro. Em modelos inteiros, fixa as inteiras
no valor da solução e analisa a PL que sobra; o relatório diz isso.

---

## Arquivos

| arquivo | o que é |
|---|---|
| `IMTSolver.xlam` | o suplemento (módulos + janelas + faixa) |
| `imtsolver\` | o motor: o executável e as bibliotecas de CBC, HiGHS e SCIP |
| `IMTSolver_Exemplos.xlsx` | três modelos prontos |
| `Instalar.bat` / `Desinstalar.bat` | instalação de um clique |
| `IMTSolver.bas`, `IMTSolverUI.bas` | fonte dos módulos |
| `frmModelo.txt`, `frmProgresso.txt` | fonte das janelas |
| `customUI14.xml` | a faixa de opções |
| `injetar_faixa.py` | põe a faixa dentro do `.xlam` |

---

## Montar o `.xlam` a partir dos fontes (só quem desenvolve)

Um `.xlam` tem uma parte binária que só o Excel gera, então a montagem
é feita nele, uma vez:

1. Excel → pasta de trabalho **em branco** → **Alt+F11**
2. **Arquivo → Importar** `IMTSolver.bas`, depois `IMTSolverUI.bas`
3. **Inserir → UserForm**. No painel Propriedades, `(Name)` =
   `frmProgresso`. **F7** abre o código: apague o que tiver e cole
   `frmProgresso.txt` inteiro
4. Repita: **Inserir → UserForm**, `(Name)` = `frmModelo`, cole
   `frmModelo.txt`
5. Clique em `VBAProject` no painel esquerdo e, em Propriedades,
   `(Name)` = `IMTSolver`
6. **Depurar → Compilar IMTSolver** — tem que ficar em silêncio
7. Volte ao Excel: **Arquivo → Salvar como** → **Suplemento do Excel
   (\*.xlam)** → nome `IMTSolver`, **nesta pasta** (o Excel tenta pular
   para a pasta AddIns; volte para cá)
8. `python injetar_faixa.py IMTSolver.xlam` — põe a aba na faixa

Pronto para distribuir: a pasta inteira, zipada.

---

## Limitações desta versão

- Só modelos lineares e lineares inteiros mistos. Não trata não-linear.
- A leitura por perturbação recalcula a planilha uma vez por variável: em
  planilhas pesadas, com milhares de variáveis, pode levar minutos. A
  janela mostra o andamento.
- A sensibilidade de um modelo inteiro é a da PL com as inteiras fixadas —
  é o que dá para fazer; preço-sombra não é definido em programação
  inteira.
- No NEOS, o tempo na fila depende do movimento do servidor.
- O motor instalado ocupa cerca de 240 MB em disco (é o preço de começar
  em 0,1 s em vez de 1,4 s).

---

## Velocidade

Medido nesta máquina, três execuções cada:

| | arquivo único | pasta |
|---|---|---|
| iniciar o motor | 1,42 s | **0,10 s** |
| listar os solvers | 2,25 s | **0,17 s** |
| resolver um modelo pequeno, ponta a ponta | 1,49 s | **0,15 s** |

Com a lista de solvers guardada, abrir a janela **Modelo** não chama o
motor nenhuma vez.
