# IMTSolver

Otimização linear e inteira dentro do Excel, com o solver que você escolher.

> Criado por **Pedro da Silva Bezerra**, aluno do 3º período do
> **Instituto Mauá de Tecnologia**, para a disciplina de
> **Pesquisa Operacional I**, com o incentivo da **professora Joyce**.
>
> © 2026 Pedro da Silva Bezerra — licença MIT. Quem copiar ou modificar
> este código precisa manter este aviso de autoria.

Você monta o modelo numa aba do Excel — células de valor para as
decisões, fórmulas para o objetivo e para as restrições — e o IMTSolver
lê os coeficientes direto das fórmulas, chama o solver e devolve a
resposta na planilha, com relatório de resposta e de sensibilidade no
formato do Solver do Excel.

## Baixar e instalar

1. Vá em **[Releases](../../releases)** e baixe o `IMTSolver_2.0_Setup_leve.exe`.
2. Feche o Excel e rode o instalador. Não precisa de administrador.
3. Se o Windows mostrar *"O Windows protegeu o seu PC"*: **Mais informações →
   Executar assim mesmo**. O aviso aparece porque o instalador não tem
   assinatura digital paga.
4. Abra o Excel: a aba **IMTSolver** aparece na faixa de opções.

Funciona no Excel para Windows (Office 2010 em diante, incluindo o 365).

## O que ele faz

- **Janela de modelo**: objetivo, variáveis e restrições. O modelo fica
  salvo **na própria aba**, nos nomes `solver_*` — o mesmo formato do
  Solver do Excel e do OpenSolver. Cada aba pode ter o seu.
- **Vários solvers**: CBC e HiGHS na sua máquina; CPLEX, COPT e SCIP de
  graça pelo [NEOS Server](https://neos-server.org) (só pede um e-mail).
- **Acompanhamento ao vivo**: melhor solução, limite, gap e tempo, sem
  travar o Excel, com botão Cancelar.
- **Relatórios** de resposta e de sensibilidade (preço-sombra, custo
  reduzido, faixas permitidas).
- **Verificar modelo**: aponta células suspeitas — variável com fórmula,
  restrição que não depende de variável nenhuma, coeficientes em escalas
  muito diferentes — sem impedir de resolver.
- **Teste de linearidade**: avisa se o modelo tem `SE`, `MÁXIMO` ou
  variável vezes variável, que um solver de programação linear não resolve.

O manual completo está em [`v2/LEIAME.md`](v2/LEIAME.md).

## Como está organizado

| caminho | o que é |
|---|---|
| `v2/IMTSolver.bas`, `v2/IMTSolverUI.bas` | o suplemento em VBA: leitura do modelo, relatórios, faixa de opções |
| `v2/frmModelo.txt`, `v2/frmProgresso.txt` | código das janelas (os controles são criados em tempo de execução) |
| `v2/customUI14.xml` | a faixa de opções |
| `v2/IMTSolver.xlam` | o suplemento já montado |
| `byesolver.py` | o motor em Python: recebe o `.lp`, chama o solver, grava o progresso e a solução (nome histórico: o projeto se chamava ByeSolver) |
| `imtsolver_pasta.spec` | empacota o motor como `imtsolver.exe` com o PyInstaller |
| `v2/IMTSolver.iss` | o instalador (Inno Setup) |

## Montar a partir do código

```bash
pip install highspy pulp pyinstaller
pyinstaller imtsolver_pasta.spec --distpath dist_imt
```

Copie `dist_imt/imtsolver` para `v2/imtsolver_leve` e compile o
instalador com o Inno Setup:

```bash
ISCC.exe /DLEVE v2/IMTSolver.iss
```

## Componentes de terceiros

O IMTSolver chama solvers de outros projetos, cada um com a sua licença:
[HiGHS](https://highs.dev) (MIT), [CBC](https://github.com/coin-or/Cbc)
(EPL-2.0, via [PuLP](https://github.com/coin-or/pulp)) e, pela internet,
os solvers do [NEOS Server](https://neos-server.org). A licença MIT
acima vale para o código do IMTSolver.
