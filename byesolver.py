# -*- coding: utf-8 -*-
"""
ByeSolver 2.0 - motor de otimizacao para o suplemento de Excel.

Recebe um modelo em formato .lp, resolve com o solver escolhido (local ou
no NEOS) e devolve a solucao num arquivo texto simples. Enquanto roda,
mantem um arquivo de progresso que o Excel le a cada segundo, e obedece
a um arquivo de cancelamento que o Excel cria quando o usuario desiste.

O suplemento nunca fala com solver nenhum: escreve o .lp e chama este
programa. Acrescentar um solver nao mexe em nada do lado do Excel.

Uso:
    byesolver --list
        solvers disponiveis, um por linha:
        NOME|rotulo|licenca|origem|tipos
        (origem = local ou neos; tipos = classes de problema)

    byesolver --lp modelo.lp --solver HiGHS --out solucao.txt
              [--timelimit 300] [--gap 0.0] [--log solver.log]
              [--progress progresso.txt] [--cancel cancelar.flag]
              [--sens] [--email voce@exemplo.com]

Arquivo de progresso (reescrito a cada meio segundo, chave=valor):
    fase        iniciando | lendo_modelo | resolvendo | neos_enviando |
                neos_fila | neos_executando | sensibilidade | gravando |
                concluido | cancelado | erro
    solver, segundos, incumbente, limite, gap, nos, mensagem, pid

Arquivo de solucao:
    status=Optimal            Optimal | Infeasible | Unbounded | TimeLimit |
                              Cancelled | Unknown | Erro
    objetivo=90.0
    limite=90.0
    gap=0.0
    segundos=12.3456        (segundos, 4 casas)
    solver=HiGHS
    variaveis=360
    nos=17
    sensibilidade=nao         nao | sim | fixando_inteiras
    mensagem=
    ---
    nome valor                (uma variavel por linha)
    ---sens_var
    nome valor custo_reduzido coef_objetivo aumento_permitido reducao_permitida
    ---sens_restr
    nome valor_final preco_sombra lado_direito aumento_permitido reducao_permitida sinal
"""

import argparse
import io
import os
import re
import sys
import threading
import time

VERSAO = "2.0"

# Falhas de rede, para separar "a internet caiu" de um defeito do motor.
# Nao da para usar OSError inteiro: ele tambem pega arquivo que nao abre,
# e ai a mensagem culparia a internet por um problema no disco.
import socket as _socket
import ssl as _ssl
import xmlrpc.client as _xmlrpc
REDE = (TimeoutError, ConnectionError, _socket.gaierror, _ssl.SSLError,
        _xmlrpc.ProtocolError)
INF_EXCEL = 1e30      # o Solver do Excel mostra 1E+30 onde e "sem limite"
SEM_VALOR = 1e19      # acima disso, o solver esta dizendo "nao tenho"


# ==================================================================
# Progresso e cancelamento
# ==================================================================
def fmt(v):
    """Numero -> texto, sem locale, inf como o Excel mostra."""
    if v is None or v == "":
        return ""
    if isinstance(v, float):
        if v != v:                       # NaN
            return ""
        if abs(v) >= SEM_VALOR:
            return "1E+30" if v > 0 else "-1E+30"
        if abs(v) < 1e-9:
            v = 0.0
        return f"{v:.10g}"
    return str(v)


def num(v):
    """Valor que o solver deu -> float ou None se ele nao tinha."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    if v != v or abs(v) >= SEM_VALOR:
        return None
    return v


class Progresso:
    """Arquivo chave=valor que o Excel le enquanto o solver roda.

    Escrito de forma atomica (temporario + rename) para o Excel nunca
    ler um arquivo pela metade.
    """

    def __init__(self, caminho, solver):
        self.caminho = caminho
        self.t0 = time.perf_counter()
        self.lock = threading.Lock()
        self.ultima_escrita = 0.0
        self.d = {"fase": "iniciando", "solver": solver, "segundos": 0.0,
                  "incumbente": "", "limite": "", "gap": "", "nos": "",
                  "mensagem": "", "pid": os.getpid(), "versao": VERSAO}
        self.set()

    def set(self, **kw):
        """Atualiza e grava agora."""
        with self.lock:
            self.d.update(kw)
            self._gravar()

    def tick(self, **kw):
        """Atualiza; grava no maximo duas vezes por segundo."""
        with self.lock:
            self.d.update(kw)
            if time.time() - self.ultima_escrita >= 0.5:
                self._gravar()

    def _gravar(self):
        self.d["segundos"] = round(time.perf_counter() - self.t0, 2)
        self.ultima_escrita = time.time()
        if not self.caminho:
            return
        tmp = self.caminho + ".tmp"
        try:
            with io.open(tmp, "w", encoding="utf-8") as f:
                for k, v in self.d.items():
                    f.write(f"{k}={fmt(v)}\n")
            os.replace(tmp, self.caminho)
        except OSError:
            pass


class Cancelador:
    """O Excel cria este arquivo quando o usuario clica em Cancelar."""

    def __init__(self, caminho):
        self.caminho = caminho
        self.ultima = 0.0
        self.sim = False
        if caminho and os.path.exists(caminho):
            os.remove(caminho)              # sobra de uma execucao anterior

    def pedido(self):
        if self.sim:
            return True
        if not self.caminho:
            return False
        agora = time.time()
        if agora - self.ultima < 0.3:       # nao martelar o disco
            return False
        self.ultima = agora
        self.sim = os.path.exists(self.caminho)
        return self.sim


class ErroUsuario(Exception):
    """Erro que e culpa da configuracao, nao do programa."""


# ==================================================================
# Catalogo de solvers
# ==================================================================
CATALOGO = []


def registrar(nome, rotulo, licenca, origem="local", tipos="PL e PLIM"):
    """tipos: que classes de problema o solver resolve.

    PL   = programacao linear (variaveis continuas)
    PLIM = linear inteira mista
    QP   = quadratica    SOCP/conica    MINLP = nao-linear inteira

    Vale um aviso: o ByeSolver so MONTA modelos lineares, porque a leitura
    por perturbacao exige linearidade. O que cada solver sabe fazer alem
    disso fica sem uso aqui - mas e informacao honesta na hora de escolher.
    """
    def wrap(cls):
        CATALOGO.append({"nome": nome, "rotulo": rotulo, "licenca": licenca,
                         "origem": origem, "tipos": tipos, "cls": cls})
        return cls
    return wrap


class Resultado:
    def __init__(self, status="Unknown", obj=None, limite=None, gap=None,
                 valores=None, nos=None, mensagem=""):
        self.status = status
        self.obj = obj
        self.limite = limite
        self.gap = gap
        self.valores = valores or {}
        self.nos = nos
        self.mensagem = mensagem


class Base:
    """Interface: disponivel() diz se da para usar; resolver() resolve."""

    @staticmethod
    def disponivel():
        raise NotImplementedError

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        raise NotImplementedError


def gap_relativo(inc, lim):
    if inc is None or lim is None:
        return None
    if abs(inc) < 1e-9:
        return abs(inc - lim)
    return abs(inc - lim) / abs(inc)


def nomes_do_lp(caminho):
    """Nomes das variaveis declaradas no .lp.

    O CBC, com printingOptions all, lista restricoes E variaveis no mesmo
    arquivo de solucao. Para devolver so variaveis, precisamos saber quais
    nomes sao variaveis.

    Nao basta ler a secao Bounds: o formato .lp omite de la as variaveis
    com o limite padrao (0 a infinito). Entao varremos o arquivo inteiro e
    ficamos com todo identificador que NAO seja nome de restricao (esses
    vem seguidos de dois-pontos) nem palavra reservada do formato.
    """
    RESERVADAS = {
        "maximize", "minimize", "max", "min", "subject", "to", "st", "s.t.",
        "such", "that", "bounds", "bound", "generals", "general", "gen",
        "binaries", "binary", "bin", "integers", "integer", "int", "free",
        "end", "inf", "infinity", "sec", "sos", "semi-continuous",
    }
    texto = io.open(caminho, encoding="utf-8", errors="replace").read()
    # comentario do formato lp: uma barra invertida ate o fim da linha
    texto = re.sub(r"\\.*", "", texto)
    nomes = []
    vistos = set()
    # o (?<![\d.]) evita ler o expoente de 1.5E-05 como se fosse um nome
    for m in re.finditer(r"(?<![\d.])([A-Za-z_][A-Za-z0-9_.\[\]\-]*)\s*(:?)",
                         texto):
        ident, doispontos = m.group(1), m.group(2)
        if doispontos == ":" or ident.lower() in RESERVADAS:
            continue
        if ident not in vistos:
            vistos.add(ident)
            nomes.append(ident)
    return nomes


def exe_empacotado(nome):
    """Executavel que veio junto do byesolver.exe, se houver."""
    base = getattr(sys, "_MEIPASS", None)
    cands = []
    if base:
        cands.append(os.path.join(base, nome))
    cands.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])),
                              nome))
    for c in cands:
        if os.path.exists(c):
            return c
    return None


# ------------------------------------------------------------------
# CBC: processo externo; progresso lendo o log linha a linha
# ------------------------------------------------------------------
@registrar("CBC", "CBC (COIN-OR)", "livre", tipos="PL, PLIM")
class SolverCBC(Base):
    @staticmethod
    def _exe():
        e = exe_empacotado("cbc.exe")
        if e:
            return e
        try:
            from pulp.apis import PULP_CBC_CMD
            p = PULP_CBC_CMD().path
            return p if os.path.exists(p) else None
        except Exception:
            return None

    @staticmethod
    def disponivel():
        return SolverCBC._exe() is not None

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        import queue
        import subprocess

        saida = lp + ".cbcsol"
        if os.path.exists(saida):
            os.remove(saida)
        cmd = [SolverCBC._exe(), lp]
        if tempo:
            cmd += ["-seconds", str(int(tempo))]
        if gap is not None:
            cmd += ["-ratio", str(gap)]
        # printingOptions all: lista TODAS as variaveis, inclusive as
        # zeradas. Sem isso nao da para distinguir "zero" de "nao veio".
        cmd += ["-printingOptions", "all", "-solve", "-solution", saida]

        flags = 0
        if os.name == "nt":
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, text=True,
                                encoding="utf-8", errors="replace",
                                bufsize=1, creationflags=flags)
        fila = queue.Queue()

        def leitor():
            for linha in proc.stdout:
                fila.put(linha)
            fila.put(None)

        threading.Thread(target=leitor, daemon=True).start()

        flog = io.open(log, "w", encoding="utf-8") if log else None
        inc = lim = None
        nos = None
        cancelado = False
        re_nos = re.compile(r"Cbc0010I After (\d+) nodes, \d+ on tree, "
                            r"(\S+) best solution, best possible (\S+)")
        # 0004/0012: solucao inteira na arvore; 0038: achada pela heuristica
        # (feasibility pump), que costuma vir bem antes do primeiro no
        re_sol = re.compile(r"Cbc00(?:04|12)I Integer solution of (\S+) found"
                            r"|Cbc0038I Solution found of (\S+)")
        acabou = False
        while not acabou:
            try:
                linha = fila.get(timeout=0.4)
            except queue.Empty:
                linha = ""
            if linha is None:
                acabou = True
                continue
            if linha:
                if flog:
                    flog.write(linha)
                    flog.flush()
                m = re_nos.search(linha)
                if m:
                    nos = int(m.group(1))
                    inc = num(m.group(2))
                    lim = num(m.group(3))
                    prog.tick(incumbente=inc, limite=lim,
                              gap=gap_relativo(inc, lim), nos=nos)
                m = re_sol.search(linha)
                if m:
                    inc = num(m.group(1) or m.group(2))
                    prog.tick(incumbente=inc, gap=gap_relativo(inc, lim))
            if not cancelado and canc.pedido():
                cancelado = True
                proc.terminate()
        proc.wait()
        if flog:
            flog.close()

        res = Resultado(nos=nos, limite=lim)
        if cancelado:
            res.status = "Cancelled"
            res.obj = inc
            return res
        if not os.path.exists(saida):
            res.status = "Erro"
            res.mensagem = "o CBC nao gerou arquivo de solucao"
            return res
        valores = {}
        for i, linha in enumerate(io.open(saida, encoding="utf-8",
                                          errors="replace")):
            if i == 0:
                cab = linha.strip()
                low = cab.lower()
                if "infeasible" in low:
                    res.status = "Infeasible"
                elif "unbounded" in low:
                    res.status = "Unbounded"
                elif "stopped on time" in low:
                    res.status = "TimeLimit"
                elif "optimal" in low:
                    res.status = "Optimal"
                m = re.search(r"objective value\s+([-\d.eE+]+)", cab)
                if m:
                    res.obj = num(m.group(1))
                continue
            p = linha.split()
            if len(p) >= 3:
                try:
                    valores[p[1]] = float(p[2])
                except ValueError:
                    pass
        declaradas = nomes_do_lp(lp)
        res.valores = {k: valores.get(k, 0.0) for k in declaradas}
        if res.status == "Optimal":
            res.limite = res.obj
            res.gap = 0.0
        else:
            res.gap = gap_relativo(res.obj, res.limite)
        return res


# ------------------------------------------------------------------
# HiGHS: dentro do processo; progresso e cancelamento por callback
# ------------------------------------------------------------------
@registrar("HiGHS", "HiGHS", "livre", tipos="PL, PLIM, QP")
class SolverHiGHS(Base):
    @staticmethod
    def disponivel():
        try:
            import highspy  # noqa
            return True
        except Exception:
            return False

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        import highspy
        h = highspy.Highs()
        h.setOptionValue("output_flag", bool(log))
        h.setOptionValue("log_to_console", False)
        if log:
            h.setOptionValue("log_file", log)
        if tempo:
            h.setOptionValue("time_limit", float(tempo))
        if gap is not None:
            h.setOptionValue("mip_rel_gap", float(gap))
        h.readModel(lp)

        T = highspy.cb.HighsCallbackType
        T_MIP = int(T.kCallbackMipInterrupt)
        T_MELHOR = int(T.kCallbackMipImprovingSolution)
        T_SIMPLEX = int(T.kCallbackSimplexInterrupt)
        T_IPM = int(T.kCallbackIpmInterrupt)

        def cb(tipo, msg, saida, entrada, user):
            t = int(tipo)
            if t in (T_MIP, T_MELHOR):
                prog.tick(incumbente=num(saida.mip_primal_bound),
                          limite=num(saida.mip_dual_bound),
                          gap=num(saida.mip_gap),
                          nos=saida.mip_node_count)
            elif t in (T_SIMPLEX, T_IPM):
                prog.tick(nos=saida.simplex_iteration_count)
            if canc.pedido():
                entrada.user_interrupt = True

        h.setCallback(cb, None)
        for t in (T.kCallbackMipInterrupt, T.kCallbackMipImprovingSolution,
                  T.kCallbackSimplexInterrupt, T.kCallbackIpmInterrupt):
            h.startCallback(t)
        h.run()

        st = h.modelStatusToString(h.getModelStatus())
        status = {"Optimal": "Optimal", "Infeasible": "Infeasible",
                  "Unbounded": "Unbounded",
                  "Time limit reached": "TimeLimit",
                  "Interrupt": "Cancelled"}.get(st, st)
        if canc.sim and status not in ("Optimal",):
            status = "Cancelled"
        info = h.getInfo()
        nomes = list(h.getLp().col_names_)
        sol = h.getSolution()
        valores = {}
        if sol.value_valid:
            valores = {nomes[i]: sol.col_value[i] for i in range(len(nomes))}
        obj = num(info.objective_function_value) if valores else None
        lim = num(getattr(info, "mip_dual_bound", None))
        g = num(getattr(info, "mip_gap", None))
        if status == "Optimal":
            lim, g = obj, 0.0
        nos = getattr(info, "mip_node_count", None)
        if nos is not None and nos < 0:     # PL pura: HiGHS devolve -1
            nos = None
        return Resultado(status, obj, lim, g, valores, nos)


# ------------------------------------------------------------------
# SCIP: dentro do processo; progresso e cancelamento por evento
# ------------------------------------------------------------------
@registrar("SCIP", "SCIP", "academico", tipos="PL, PLIM, MINLP")
class SolverSCIP(Base):
    @staticmethod
    def disponivel():
        try:
            import pyscipopt  # noqa
            return True
        except Exception:
            return False

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        from pyscipopt import Model, Eventhdlr, SCIP_EVENTTYPE

        m = Model()
        if log:
            m.setLogfile(log)
        m.hideOutput(True)          # o arquivo de log continua sendo gravado
        m.readProblem(lp)
        if tempo:
            m.setParam("limits/time", float(tempo))
        if gap is not None:
            m.setParam("limits/gap", float(gap))

        class Olheiro(Eventhdlr):
            def eventinit(self):
                self.model.catchEvent(SCIP_EVENTTYPE.NODESOLVED, self)
                self.model.catchEvent(SCIP_EVENTTYPE.BESTSOLFOUND, self)

            def eventexit(self):
                self.model.dropEvent(SCIP_EVENTTYPE.NODESOLVED, self)
                self.model.dropEvent(SCIP_EVENTTYPE.BESTSOLFOUND, self)

            def eventexec(self, event):
                mm = self.model
                prog.tick(incumbente=num(mm.getPrimalbound()),
                          limite=num(mm.getDualbound()),
                          gap=num(mm.getGap()), nos=mm.getNNodes())
                if canc.pedido():
                    mm.interruptSolve()

        m.includeEventhdlr(Olheiro(), "byesolver", "progresso e cancelamento")
        m.optimize()

        st = m.getStatus()
        status = {"optimal": "Optimal", "infeasible": "Infeasible",
                  "unbounded": "Unbounded", "timelimit": "TimeLimit",
                  "userinterrupt": "Cancelled"}.get(st, st)
        res = Resultado(status, nos=m.getNNodes())
        if m.getNSols() > 0:
            s = m.getBestSol()
            res.valores = {v.name: m.getSolVal(s, v) for v in m.getVars()}
            res.obj = num(m.getObjVal())
            res.limite = num(m.getDualbound())
            res.gap = num(m.getGap())
            if status == "Optimal":
                res.limite, res.gap = res.obj, 0.0
        return res


# ------------------------------------------------------------------
# Comerciais locais, se a pessoa tiver licenca
# ------------------------------------------------------------------
@registrar("GUROBI", "Gurobi (local)", "licenca", tipos="PL, PLIM, QP, conica")
class SolverGurobi(Base):
    @staticmethod
    def disponivel():
        try:
            import gurobipy  # noqa
            return True
        except Exception:
            return False

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        import gurobipy as gp
        mod = gp.read(lp)
        mod.Params.OutputFlag = 1 if log else 0
        if log:
            mod.Params.LogFile = log
        if tempo:
            mod.Params.TimeLimit = float(tempo)
        if gap is not None:
            mod.Params.MIPGap = float(gap)

        def cb(model, where):
            if where == gp.GRB.Callback.MIP:
                inc = num(model.cbGet(gp.GRB.Callback.MIP_OBJBST))
                lim = num(model.cbGet(gp.GRB.Callback.MIP_OBJBND))
                prog.tick(incumbente=inc, limite=lim,
                          gap=gap_relativo(inc, lim),
                          nos=int(model.cbGet(gp.GRB.Callback.MIP_NODCNT)))
            if canc.pedido():
                model.terminate()

        mod.optimize(cb)
        status = {2: "Optimal", 3: "Infeasible", 5: "Unbounded",
                  9: "TimeLimit", 11: "Cancelled"}.get(
            mod.Status, f"codigo {mod.Status}")
        res = Resultado(status, nos=int(mod.NodeCount))
        if mod.SolCount > 0:
            res.valores = {v.VarName: v.X for v in mod.getVars()}
            res.obj = num(mod.ObjVal)
            res.limite = num(mod.ObjBound)
            res.gap = num(mod.MIPGap)
        return res


@registrar("CPLEX", "IBM CPLEX (local)", "licenca", tipos="PL, PLIM, QP, conica")
class SolverCPLEX(Base):
    @staticmethod
    def disponivel():
        try:
            import cplex  # noqa
            return True
        except Exception:
            return False

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        import cplex
        c = cplex.Cplex(lp)
        if not log:
            c.set_log_stream(None)
            c.set_results_stream(None)
            c.set_warning_stream(None)
            c.set_error_stream(None)
        else:
            f = io.open(log, "w", encoding="utf-8")
            c.set_log_stream(f)
            c.set_results_stream(f)
        if tempo:
            c.parameters.timelimit.set(float(tempo))
        if gap is not None:
            c.parameters.mip.tolerances.mipgap.set(float(gap))
        c.solve()
        st = c.solution.get_status_string().lower()
        status = "Optimal" if "optimal" in st else (
            "Infeasible" if "infeasible" in st else (
                "TimeLimit" if "time" in st else st))
        valores = dict(zip(c.variables.get_names(), c.solution.get_values()))
        return Resultado(status, num(c.solution.get_objective_value()),
                         valores=valores)


@registrar("GLPK", "GLPK", "livre", tipos="PL, PLIM")
class SolverGLPK(Base):
    @staticmethod
    def _exe():
        from shutil import which
        return which("glpsol")

    @staticmethod
    def disponivel():
        return SolverGLPK._exe() is not None

    @staticmethod
    def resolver(lp, tempo, gap, log, prog, canc, email=None):
        import subprocess
        saida = lp + ".glpksol"
        cmd = [SolverGLPK._exe(), "--lp", lp, "-o", saida]
        if tempo:
            cmd += ["--tmlim", str(int(tempo))]
        if gap is not None:
            cmd += ["--mipgap", str(gap)]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
        # Popen em vez de run: com run() o botao Cancelar nao fazia nada,
        # porque nada olhava o pedido ate o glpsol terminar sozinho.
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, text=True,
                                encoding="utf-8", errors="replace",
                                creationflags=flags)
        cancelado = False
        while proc.poll() is None:
            time.sleep(0.3)
            if canc.pedido():
                cancelado = True
                proc.terminate()
                break
        try:
            texto_saida = proc.communicate(timeout=10)[0] or ""
        except Exception:
            proc.kill()
            texto_saida = ""
        if log:
            io.open(log, "w", encoding="utf-8", errors="replace").write(texto_saida)
        if cancelado:
            return Resultado("Cancelled", mensagem="cancelado pelo usuario")
        res = Resultado()
        if os.path.exists(saida):
            texto = io.open(saida, encoding="utf-8", errors="replace").read()
            if "OPTIMAL" in texto:
                res.status = "Optimal"
            elif "INFEASIBLE" in texto or "NO PRIMAL" in texto:
                res.status = "Infeasible"
            m = re.search(r"Objective:\s+\S+\s*=\s*([-\d.eE+]+)", texto)
            if m:
                res.obj = num(m.group(1))
            for linha in texto.splitlines():
                p = linha.split()
                if len(p) >= 4 and p[0].isdigit():
                    try:
                        res.valores[p[1]] = float(p[3])
                    except ValueError:
                        pass
        return res


# ==================================================================
# NEOS: solvers comerciais de graca, pela API oficial (XML-RPC)
# ==================================================================
class NEOS:
    URL = "https://neos-server.org:3333"
    _alcancavel = None

    @classmethod
    def proxy(cls, timeout=30):
        # O tempo limite vale para as conexoes abertas dali em diante; por
        # isso cada etapa que precisa de outro prazo pede um proxy novo.
        import socket
        import xmlrpc.client
        socket.setdefaulttimeout(timeout)
        return xmlrpc.client.ServerProxy(cls.URL)

    @classmethod
    def chamar(cls, metodo, *args, timeout=30, tentativas=4, prog=None,
               oque="responder"):
        """Chama o NEOS tolerando rede lenta ou servidor ocupado.

        Cada tentativa abre uma conexao nova, com o seu proprio prazo. So
        serve para chamadas que podem ser repetidas sem efeito colateral -
        pedir o modelo, baixar o resultado. Enviar o trabalho NAO passa
        por aqui: repetir o envio criaria um segundo trabalho no NEOS.
        """
        import xmlrpc.client
        erro = None
        for k in range(tentativas):
            try:
                return getattr(cls.proxy(timeout), metodo)(*args)
            except xmlrpc.client.Fault:
                raise           # o NEOS respondeu com um erro: repetir nao muda
            except (OSError, xmlrpc.client.ProtocolError) as e:
                erro = e        # tempo esgotado, conexao caida, SSL
                if k + 1 < tentativas:
                    if prog:
                        prog.set(mensagem=f"o NEOS demorou para {oque}; "
                                          f"tentando de novo ({k + 2} de "
                                          f"{tentativas})")
                    time.sleep(3 * (k + 1))
        raise erro

    @classmethod
    def alcancavel(cls):
        if cls._alcancavel is None:
            try:
                import socket
                import xmlrpc.client
                socket.setdefaulttimeout(4)
                cls._alcancavel = "alive" in str(
                    xmlrpc.client.ServerProxy(cls.URL).ping()).lower()
            except Exception:
                cls._alcancavel = False
        return cls._alcancavel


# o que cada solver do NEOS aceita como opcoes (campo -> texto)
def _opcoes_neos(solver, tempo, gap):
    t = int(tempo) if tempo else None
    if solver == "CPLEX":
        o = []
        if t:
            o.append(f"set timelimit {t}")
        if gap is not None:
            o.append(f"set mip tolerances mipgap {gap}")
        return {"options": "\n".join(o),
                "post": "display solution objective\n"
                        "display solution variables -"}
    if solver == "Gurobi":
        o = []
        if t:
            o.append(f"TimeLimit {t}")
        if gap is not None:
            o.append(f"MIPGap {gap}")
        return {"param": "\n".join(o)}
    if solver == "scip":
        o = []
        if t:
            o.append(f"limits/time = {t}")
        if gap is not None:
            o.append(f"limits/gap = {gap}")
        return {"options": "\n".join(o)}
    if solver == "MOSEK":
        o = []
        if t:
            o.append(f"MSK_DPAR_MIO_MAX_TIME {t}")
            o.append(f"MSK_DPAR_OPTIMIZER_MAX_TIME {t}")
        if gap is not None:
            o.append(f"MSK_DPAR_MIO_TOL_REL_GAP {gap}")
        return {"options": "\n".join(o)}
    if solver == "HiGHS":
        o = []
        if t:
            o.append(f"time_limit = {t}")
        if gap is not None:
            o.append(f"mip_rel_gap = {gap}")
        return {"options": "\n".join(o)}
    if solver == "COPT":
        o = []
        if t:
            o.append(f"TimeLimit {t}")
        if gap is not None:
            o.append(f"RelGap {gap}")
        return {"options": "\n".join(o), "param": "\n".join(o)}
    return {}


def _ler_solucao_neos(texto, declaradas):
    """Tenta cada formato de solucao que o NEOS devolve. Primeiro que
    render valores, vence. Devolve (valores, objetivo)."""
    obj = None
    # o separador e opcional: o .sol do COPT escreve "# Objective value 27",
    # sem dois-pontos, enquanto o CPLEX escreve "Objective value:  9.0"
    for m in re.finditer(r"objective(?:\s+value)?\s*[:=]?\s+"
                         r"([-+]?\d[\d.]*(?:[eE][-+]?\d+)?)\b",
                         texto, re.IGNORECASE):
        obj = num(m.group(1))
        if obj is not None:
            break
    conjunto = set(declaradas)

    # 1) Gurobi .sol / scip / generico:  nome  valor
    valores = {}
    for m in re.finditer(r"^\s*([A-Za-z_][\w.\[\]\-]*)\s+([-\d.eE+]+)"
                         r"(?:\s+\(obj:[^)]*\))?\s*$", texto, re.M):
        if m.group(1) in conjunto:
            valores[m.group(1)] = float(m.group(2))
    if valores:
        return valores, obj

    # 2) CPLEX em XML:  <variable name="x" index="0" value="1"/>
    for m in re.finditer(r'<variable\s+name="([^"]+)"[^>]*value="([^"]+)"',
                         texto):
        if m.group(1) in conjunto:
            valores[m.group(1)] = float(m.group(2))
    if valores:
        m = re.search(r'objectiveValue="([^"]+)"', texto)
        if m:
            obj = num(m.group(1))
        return valores, obj

    # 3) MOSEK:  indice nome status valor ...
    for m in re.finditer(r"^\s*\d+\s+([A-Za-z_][\w.\[\]\-]*)\s+\S+\s+"
                         r"([-\d.eE+]+)", texto, re.M):
        if m.group(1) in conjunto:
            valores[m.group(1)] = float(m.group(2))
    return valores, obj


# Frases de status, por solver. Palavra solta nao serve: o NEOS devolve o
# log inteiro, incluindo o eco dos parametros que mandamos - e um
# "set timelimit 60" fazia um otimo ser lido como limite de tempo.
_NEOS_INVIAVEL = (
    "integer infeasible", "problem is infeasible", "model is infeasible",
    "no solution exists", "status : infeasible", "status: infeasible",
    "[infeasible]", "primal infeasible", "problem is primal infeasible",
)
_NEOS_ILIMITADO = (
    "is unbounded", "problem is unbounded", "[unbounded]",
    "status : unbounded", "status: unbounded", "problem is dual infeasible",
)
_NEOS_LIMITE = (
    "time limit exceeded", "aborted, time limit", "time limit reached",
    "stopped on time limit", "terminated. time limit",
    "[time limit reached]", "due to time limit", "solution limit exceeded",
)
_NEOS_OTIMO = (
    "- optimal:", "integer optimal", "optimal solution found",
    "optimal solution isolated", "status : optimal", "status: optimal",
    "[optimal solution found]", "model status : optimal",
    "optimal objective", "problem is solved [optimal",
    "optimal - ", "solution status: optimal",
)


def _arquivo_de_solucao_neos(neos, job, senha):
    """O texto do .sol que vem dentro de solver-output.zip, ou ''.

    O NEOS so entrega alguns nomes de arquivo: results.txt, ampl.sol,
    solver-output.zip, job.in, job.out, job.results. Pedir 'soln.sol'
    direto e recusado.
    """
    import zipfile
    for nome in ("solver-output.zip", "ampl.sol"):
        try:
            r = NEOS.chamar("getOutputFile", job, senha, nome, timeout=120,
                            tentativas=3)
            d = r.data if hasattr(r, "data") else bytes(str(r), "utf-8")
        except Exception:
            continue
        if not d or b"does not exist" in d[:60]:
            continue
        if nome.endswith(".zip"):
            try:
                z = zipfile.ZipFile(io.BytesIO(d))
            except Exception:
                continue
            for m in z.namelist():
                if m.lower().endswith((".sol", ".txt")):
                    return z.read(m).decode("utf-8", "replace")
        else:
            return d.decode("utf-8", "replace")
    return ""


def _status_neos(texto, cancelado):
    if cancelado:
        return "Cancelled"
    t = texto.lower()
    for grupo, st in ((_NEOS_INVIAVEL, "Infeasible"),
                      (_NEOS_ILIMITADO, "Unbounded"),
                      (_NEOS_LIMITE, "TimeLimit"),
                      (_NEOS_OTIMO, "Optimal")):
        for p in grupo:
            if p in t:
                return st
    return "Unknown"


def fabrica_neos(solver_neos, categoria):
    class SolverNEOS(Base):
        @staticmethod
        def disponivel():
            # Um solver do NEOS esta sempre "disponivel" como opcao: e um
            # servico remoto, nao algo instalado aqui. Nao damos ping ao
            # listar porque isso custava quase um segundo toda vez que a
            # janela do modelo abria. Se a internet estiver fora, o erro
            # aparece na hora de resolver, com uma mensagem clara.
            return True

        @staticmethod
        def resolver(lp, tempo, gap, log, prog, canc, email=None):
            if not email or "@" not in email:
                raise ErroUsuario(
                    "O NEOS exige um e-mail valido. Preencha na janela "
                    "do modelo, no campo de e-mail.")
            prog.set(fase="neos_enviando", mensagem="pedindo o modelo ao NEOS")
            try:
                modelo = NEOS.chamar("getSolverTemplate", categoria,
                                     solver_neos, "LP", tentativas=3,
                                     prog=prog)
            except Exception as e:
                raise ErroUsuario(
                    "Nao consegui falar com o NEOS. Verifique a internet, ou "
                    "escolha um solver desta maquina, como o CBC ou o HiGHS.")
            if "unknown" in modelo.lower()[:200]:
                raise ErroUsuario(f"NEOS nao aceita .lp para {solver_neos}")

            texto_lp = io.open(lp, encoding="utf-8").read()
            preencher = {"email": email, "LP": texto_lp, "comments": "",
                         "wantsol": "yes", "wantlog": "yes",
                         "wantbas": "", "wantmst": "", "wantint": ""}
            preencher.update(_opcoes_neos(solver_neos, tempo, gap))

            def campo(m):
                nome = m.group(1)
                if nome not in preencher:
                    # category, solver, inputMethod e afins: o NEOS precisa
                    # deles exatamente como mandou. Apagar dava
                    # "'category' tag not found in xml file".
                    return m.group(0)
                v = preencher[nome]
                if v == "":
                    return ""                       # campo opcional, fora
                if nome == "email":
                    return f"<email>{v}</email>"
                return f"<{nome}><![CDATA[{v}]]></{nome}>"

            xml = re.sub(r"<(\w+)>(?:<!\[CDATA\[)?[^<]*?(?:\]\]>)?</\1>",
                         campo, modelo)
            # o que sobrou com o texto de exemplo do formulario sai fora
            xml = re.sub(r"<(\w+)>(?:<!\[CDATA\[)?[^<]*?Insert Value Here"
                         r"[^<]*?(?:\]\]>)?</\1>", "", xml)

            prog.set(fase="neos_enviando", mensagem="enviando o modelo")
            try:
                # Um modelo grande leva tempo para subir: 30 s nao bastava.
                # Sem nova tentativa - repetir criaria um trabalho duplicado.
                job, senha = NEOS.proxy(180).submitJob(xml)
            except Exception:
                raise ErroUsuario(
                    "O envio ao NEOS nao terminou. A internet pode estar "
                    "lenta ou o NEOS ocupado. Tente de novo em alguns "
                    "minutos.")
            if job == 0:
                raise ErroUsuario(f"NEOS recusou o trabalho: {senha}")
            prog.set(fase="neos_fila", mensagem=f"trabalho {job} na fila")

            neos = NEOS.proxy(30)
            flog = io.open(log, "w", encoding="utf-8") if log else None
            offset = 0
            cancelado = False
            ultimo_poll = 0.0
            estado = "Waiting"
            aviso_rede = False
            while estado not in ("Done", "Killed"):
                if canc.pedido() and not cancelado:
                    cancelado = True
                    try:
                        neos.killJob(job, senha)
                    except Exception:
                        pass
                    break
                time.sleep(0.3)
                if time.time() - ultimo_poll < 2.0:
                    continue
                ultimo_poll = time.time()
                try:
                    estado = neos.getJobStatus(job, senha)
                    pedaco, offset = neos.getIntermediateResults(
                        job, senha, offset)
                    novo = pedaco.data.decode("utf-8", "replace") \
                        if hasattr(pedaco, "data") else str(pedaco)
                except Exception as e:
                    # Uma consulta que demora nao derruba nada: o trabalho
                    # segue rodando no NEOS e a proxima consulta tenta de
                    # novo. Antes a mensagem crua em ingles ("The read
                    # operation timed out") ficava presa na tela ate o NEOS
                    # mandar texto novo - parecia que tinha falhado.
                    aviso_rede = True
                    prog.tick(mensagem="o NEOS demorou para responder; o "
                                       "trabalho continua la, tentando de novo")
                    if flog:
                        # a janela mostra este log nas execucoes do NEOS
                        flog.write("\n[IMTSolver: o NEOS demorou para "
                                   "responder; tentando de novo]\n")
                        flog.flush()
                    neos = NEOS.proxy(30)   # a conexao antiga pode ter travado
                    continue
                if aviso_rede and not novo.strip():
                    prog.tick(mensagem=f"trabalho {job}: conexao retomada")
                aviso_rede = False
                if novo.strip():
                    if flog:
                        flog.write(novo)
                        flog.flush()
                    ultima = [l for l in novo.splitlines() if l.strip()]
                    if ultima:
                        prog.tick(mensagem=ultima[-1].strip()[:120])
                fase = "neos_fila" if estado == "Waiting" else "neos_executando"
                prog.tick(fase=fase)

            if cancelado:
                if flog:
                    flog.close()
                return Resultado("Cancelled", mensagem="cancelado pelo usuario")

            prog.set(fase="neos_executando", mensagem="baixando o resultado")
            # O resultado de um modelo grande e grande (a LookVision imprime
            # 5.724 variaveis): prazo maior e nova tentativa. Perder isto
            # jogava fora um trabalho de 25 minutos ja terminado no NEOS.
            try:
                final = NEOS.chamar("getFinalResults", job, senha,
                                    timeout=120, tentativas=5, prog=prog,
                                    oque="entregar o resultado")
            except Exception:
                # Cinco tentativas falharam. Antes isto chegava a tela como
                # "TimeoutError: The read operation timed out".
                if flog:
                    flog.close()
                raise ErroUsuario(
                    "O NEOS terminou o trabalho, mas a conexao falhou 5 vezes "
                    "ao trazer o resultado. Verifique a internet e resolva "
                    "de novo.")
            texto = final.data.decode("utf-8", "replace") \
                if hasattr(final, "data") else str(final)
            if flog:
                flog.write("\n\n===== resultado final =====\n")
                flog.write(texto)
                flog.close()

            declaradas = nomes_do_lp(lp)
            valores, obj = _ler_solucao_neos(texto, declaradas)

            if not valores:
                # Alguns solvers (COPT, por exemplo) nao imprimem a solucao:
                # gravam num arquivo que o NEOS entrega separado, dentro de
                # solver-output.zip. So os nomes dessa lista sao permitidos.
                prog.set(fase="neos_executando",
                         mensagem="buscando o arquivo de solucao")
                extra = _arquivo_de_solucao_neos(neos, job, senha)
                if extra:
                    if flog:
                        flog = io.open(log, "a", encoding="utf-8")
                        flog.write("\n\n===== solver-output.zip =====\n")
                        flog.write(extra[:20000])
                        flog.close()
                        flog = None
                    valores, obj2 = _ler_solucao_neos(extra, declaradas)
                    if obj is None:
                        obj = obj2
            res = Resultado(_status_neos(texto, False), obj)
            if valores:
                res.valores = {k: valores.get(k, 0.0) for k in declaradas}
                if res.status == "Optimal":
                    res.limite, res.gap = obj, 0.0
            else:
                res.mensagem = ("NEOS terminou mas nao achei a solucao no "
                                "resultado; veja o log")
                if res.status == "Optimal":
                    res.status = "Unknown"
            if res.status == "Unknown" and res.valores:
                res.mensagem = ("o NEOS devolveu a solucao mas nao um status "
                                "que eu reconheca; confira no log")
            return res

    return SolverNEOS


# Testados ao vivo no NEOS em 08/09/2026. Ficaram de fora:
#   Gurobi - o NEOS recusa: "this solver is not allowed to be used via
#            this interface" (so pela pagina web deles)
#   MOSEK  - o script do proprio NEOS quebra: NameError: name 're' is
#            not defined
#   HiGHS  - resolve mas nao devolve a solucao por nenhum caminho, e
#            nao faz falta: o HiGHS ja roda aqui na sua maquina
for _nome, _rotulo, _sneos, _tipos in (
        ("NEOS-CPLEX", "IBM CPLEX (NEOS)", "CPLEX", "PL, PLIM, QP, conica"),
        ("NEOS-COPT", "COPT (NEOS)", "COPT", "PL, PLIM, QP, conica"),
        ("NEOS-SCIP", "SCIP (NEOS)", "scip", "PL, PLIM, MINLP")):
    registrar(_nome, _rotulo, "gratis, pede e-mail", "neos",
              tipos=_tipos)(fabrica_neos(_sneos, "milp"))


# ==================================================================
# Sensibilidade: precos-sombra, custos reduzidos e faixas (HiGHS)
# ==================================================================
def sensibilidade(lp, valores, prog):
    """Relatorio no molde do Solver do Excel.

    So faz sentido em PL continua. Se o modelo tem variaveis inteiras,
    fixamos cada uma no valor da solucao e analisamos a PL que sobra -
    e dizemos isso no cabecalho. Devolve (modo, linhas_var, linhas_restr)
    ou None se nao deu.
    """
    import highspy
    h = highspy.Highs()
    h.setOptionValue("output_flag", False)
    h.readModel(lp)
    L = h.getLp()
    n, m = L.num_col_, L.num_row_
    nomes = list(L.col_names_)
    rnomes = list(L.row_names_)
    CONT = highspy.HighsVarType.kContinuous
    # numa PL pura o HiGHS deixa integrality_ vazio
    integ = list(L.integrality_) or [CONT] * n

    fixadas = set()
    for j in range(n):
        if integ[j] != CONT:
            v = valores.get(nomes[j])
            if v is None:
                return None
            v = float(round(v))
            h.changeColBounds(j, v, v)
            h.changeColIntegrality(j, CONT)
            fixadas.add(j)
    modo = "fixando_inteiras" if fixadas else "sim"

    h.run()
    if h.modelStatusToString(h.getModelStatus()) != "Optimal":
        return None
    sol = h.getSolution()
    if not sol.dual_valid:
        return None
    st, rg = h.getRanging()
    tem_faixas = bool(getattr(rg, "valid", False))

    L = h.getLp()
    custo = list(L.col_cost_)
    rlo, rup = list(L.row_lower_), list(L.row_upper_)
    inf = float("inf")

    def lim(x):
        x = float(x)
        return x if abs(x) < SEM_VALOR else (INF_EXCEL if x > 0 else -INF_EXCEL)

    lin_var = []
    for j in range(n):
        if j in fixadas:
            continue
        c = custo[j]
        if tem_faixas:
            up = lim(rg.col_cost_up.value_[j])
            dn = lim(rg.col_cost_dn.value_[j])
            aum = INF_EXCEL if up >= INF_EXCEL else up - c
            red = INF_EXCEL if dn <= -INF_EXCEL else c - dn
        else:
            aum = red = ""
        lin_var.append((nomes[j], sol.col_value[j], sol.col_dual[j], c,
                        aum, red))

    lin_restr = []
    for i in range(m):
        lo, up_ = rlo[i], rup[i]
        ativ = sol.row_value[i]
        if lo == up_:
            sinal, rhs = "=", lo
        elif lo <= -inf:
            sinal, rhs = "<=", up_
        elif up_ >= inf:
            sinal, rhs = ">=", lo
        else:
            sinal = "faixa"
            rhs = up_ if abs(ativ - up_) <= abs(ativ - lo) else lo
        if tem_faixas:
            ru = lim(rg.row_bound_up.value_[i])
            rd = lim(rg.row_bound_dn.value_[i])
            aum = INF_EXCEL if ru >= INF_EXCEL else ru - rhs
            red = INF_EXCEL if rd <= -INF_EXCEL else rhs - rd
        else:
            aum = red = ""
        lin_restr.append((rnomes[i], ativ, sol.row_dual[i], rhs, aum, red,
                          sinal))
    return modo, lin_var, lin_restr


# ==================================================================
# Linha de comando
# ==================================================================
def listar(saida=None):
    linhas = [f"{s['nome']}|{s['rotulo']}|{s['licenca']}|{s['origem']}"
              f"|{s['tipos']}"
              for s in CATALOGO if s["cls"].disponivel()]
    texto = "\n".join(linhas) + "\n"
    if saida:                     # o Excel le daqui; nao tem console
        io.open(saida, "w", encoding="utf-8").write(texto)
    print(texto, end="")
    return 0


def gravar_solucao(caminho, res, seg, solver, sens):
    with io.open(caminho, "w", encoding="utf-8") as f:
        f.write(f"status={res.status}\n")
        f.write(f"objetivo={fmt(res.obj)}\n")
        f.write(f"limite={fmt(res.limite)}\n")
        f.write(f"gap={fmt(res.gap)}\n")
        f.write(f"segundos={seg:.4f}\n")
        f.write(f"solver={solver}\n")
        f.write(f"variaveis={len(res.valores)}\n")
        f.write(f"nos={fmt(res.nos)}\n")
        f.write(f"sensibilidade={sens[0] if sens else 'nao'}\n")
        f.write(f"mensagem={res.mensagem}\n")
        f.write("---\n")
        for nome, v in res.valores.items():
            f.write(f"{nome} {fmt(float(v))}\n")
        if sens:
            _, lv, lr = sens
            f.write("---sens_var\n")
            for t in lv:
                f.write(" ".join(fmt(float(x)) if x != "" else "" for x in t[1:])
                        .join([t[0] + " ", "\n"]))
            f.write("---sens_restr\n")
            for t in lr:
                nums = " ".join(fmt(float(x)) if x != "" else "" for x in t[1:6])
                f.write(f"{t[0]} {nums} {t[6]}\n")


def resolver(args):
    escolhido = None
    for s in CATALOGO:
        if s["nome"].upper() == args.solver.upper():
            escolhido = s
            break
    prog = Progresso(args.progress, args.solver)
    canc = Cancelador(args.cancel)

    def falha(codigo, msg):
        prog.set(fase="erro", mensagem=msg)
        print("ERRO:", msg, file=sys.stderr)
        gravar_solucao(args.out, Resultado("Erro", mensagem=msg), 0,
                       args.solver, None)
        return codigo

    if escolhido is None:
        return falha(2, f"solver '{args.solver}' nao existe no catalogo")
    if not escolhido["cls"].disponivel():
        return falha(3, f"{escolhido['rotulo']} nao esta disponivel "
                        "nesta maquina")
    if not os.path.exists(args.lp):
        return falha(4, f"nao achei o modelo {args.lp}")

    prog.set(fase="lendo_modelo")
    n_vars = len(nomes_do_lp(args.lp))
    prog.set(fase="resolvendo", mensagem=f"{n_vars} variaveis")

    t0 = time.perf_counter()
    try:
        res = escolhido["cls"].resolver(args.lp, args.timelimit, args.gap,
                                        args.log, prog, canc, args.email)
    except ErroUsuario as e:
        return falha(5, str(e))
    except REDE:
        # Rede: tempo esgotado, conexao caida, sem internet. O nome tecnico
        # em ingles ("TimeoutError: The read operation timed out") nao diz
        # nada a quem esta usando. O detalhe fica no stderr.
        import traceback
        traceback.print_exc()
        return falha(7, "A conexao com a internet falhou durante a "
                        "resolucao. Verifique a internet e tente de novo.")
    except Exception as e:
        import traceback
        traceback.print_exc()
        return falha(7, f"{type(e).__name__}: {e}")
    seg = time.perf_counter() - t0

    sens = None
    if args.sens and res.valores and res.status in ("Optimal", "TimeLimit"):
        prog.set(fase="sensibilidade", incumbente=res.obj, limite=res.limite,
                 gap=res.gap, nos=res.nos)
        try:
            sens = sensibilidade(args.lp, res.valores, prog)
        except Exception as e:
            res.mensagem = (res.mensagem + " | " if res.mensagem else "") + \
                f"sensibilidade falhou: {e}"

    prog.set(fase="gravando")
    gravar_solucao(args.out, res, seg, escolhido["nome"], sens)

    fase = {"Cancelled": "cancelado", "Erro": "erro"}.get(res.status,
                                                          "concluido")
    prog.set(fase=fase, incumbente=res.obj, limite=res.limite, gap=res.gap,
             nos=res.nos, mensagem=res.mensagem or
             f"{res.status} | objetivo {fmt(res.obj)}")
    print(f"{res.status} | objetivo={fmt(res.obj)} | {seg:.2f}s | "
          f"{len(res.valores)} variaveis"
          + (f" | sensibilidade={sens[0]}" if sens else ""))
    if res.status == "Cancelled":
        return 6
    return 0 if res.status in ("Optimal", "TimeLimit") else 1


def main():
    p = argparse.ArgumentParser(prog="byesolver",
                                description="Motor de otimizacao do ByeSolver")
    p.add_argument("--list", action="store_true",
                   help="lista os solvers disponiveis")
    p.add_argument("--version", action="store_true")
    p.add_argument("--lp", help="arquivo .lp com o modelo")
    p.add_argument("--solver", default="CBC")
    p.add_argument("--out", default=None,
                   help="arquivo de saida (solucao, ou a lista com --list)")
    p.add_argument("--timelimit", type=float, default=None)
    p.add_argument("--gap", type=float, default=None)
    p.add_argument("--log", default=None)
    p.add_argument("--progress", default=None,
                   help="arquivo de progresso, reescrito enquanto roda")
    p.add_argument("--cancel", default=None,
                   help="se este arquivo aparecer, o solver para")
    p.add_argument("--sens", action="store_true",
                   help="acrescenta o relatorio de sensibilidade")
    p.add_argument("--email", default=None, help="exigido pelo NEOS")
    a = p.parse_args()
    if a.version:
        print(VERSAO)
        return 0
    if a.list:
        return listar(a.out)
    if not a.lp:
        p.print_help()
        return 1
    if not a.out:
        a.out = "solucao.txt"
    return resolver(a)


if __name__ == "__main__":
    sys.exit(main())
