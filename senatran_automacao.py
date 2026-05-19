import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import json
import os
import sys
import logging
from datetime import datetime, date
from collections import defaultdict
import subprocess

# ── Dependências opcionais (instaladas na primeira execução) ──────────────────
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-q"])

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        NoSuchElementException, TimeoutException, WebDriverException
    )
    from selenium.webdriver.chrome.service import Service
except ImportError:
    print("Instalando Selenium...")
    install("selenium")
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        NoSuchElementException, TimeoutException, WebDriverException
    )
    from selenium.webdriver.chrome.service import Service

try:
    import pandas as pd
except ImportError:
    install("pandas")
    install("openpyxl")
    import pandas as pd

try:
    from webdriver_manager.chrome import ChromeDriverManager
    WEBDRIVER_MANAGER = True
except ImportError:
    install("webdriver-manager")
    from webdriver_manager.chrome import ChromeDriverManager
    WEBDRIVER_MANAGER = True

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

CONFIG_FILE = "senatran_config.json"
LOG_FILE    = "senatran_log.txt"
DATA_FILE   = "senatran_dados.json"

HORARIO_INICIO  = "08:00"
HORARIO_FIM     = "18:00"
INTERVALO_SEG   = 300          # 5 minutos

# Seletores CSS/XPath — ajuste conforme o site real do Senatran
SELETORES = {
    "tabela_mensagens":  "//div[@_ngcontent-whm-c156=''][@class='card-body']",
    # Container principal da caixa de mensagens
 
    "linhas_tabela":     "//div[@class='autuacao border ng-star-inserted']",
    # Cada mensagem é um div.autuacao
 
    "celulas":           "//div[@class='autuacao border ng-star-inserted']//div[contains(@class, 'col-')]",
    # Colunas dentro de cada mensagem (título, data, conteúdo)
 
    "titulo_mensagem":   "//div[@class='autuacao border ng-star-inserted']//span[@class='title']",
    # Título/tipo da mensagem
 
    "data_mensagem":     "//div[@class='autuacao border ng-star-inserted']//div[@_ngcontent-whm-c156=''][contains(@class, 'col-md-4')]",
    # Data e hora da mensagem
 
    "conteudo_mensagem": "//div[@class='autuacao border ng-star-inserted']//p[@_ngcontent-whm-c156='']",
    # Conteúdo/descrição da mensagem
 
    "total_registros":   "//div[@_ngcontent-whm-c139=''][contains(text(), 'de')]",
    # Elemento com total "1-100 de 636 itens"
 
    "proxima_pagina":    "//button[@id='btn-next-page']",
    # Botão próxima página
 
    "pagina_anterior":   "//button[@id='btn-last-page']",
    # Botão página anterior
 
    "seletor_itens":     "//ng-select[contains(@class, 'ng-select')]",
    # Seletor de quantidade de itens por página
}

# ─────────────────────────────────────────────────────────────────────────────
# LOGGER
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger("Senatran")

# ─────────────────────────────────────────────────────────────────────────────
# MOTOR DE COLETA
# ─────────────────────────────────────────────────────────────────────────────

class SenatranColetor:
    """Conecta a um Chrome já aberto (debuggingPort) e raspa a caixa de entrada."""

    def __init__(self, porta_debug: int = 9222):
        self.porta_debug   = porta_debug
        self.driver        = None
        self.conectado     = False
        self.multas_vistas = set()   # IDs já coletados nesta sessão

    # ── Conexão ───────────────────────────────────────────────────────────────

    def conectar(self) -> bool:
        try:
            opts = Options()
            opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.porta_debug}")
            opts.add_argument("--disable-blink-features=AutomationControlled")

            if WEBDRIVER_MANAGER:
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=opts)
            else:
                self.driver = webdriver.Chrome(options=opts)

            # Testa se o driver responde
            _ = self.driver.title
            self.conectado = True
            logger.info(f"✅ Conectado ao Chrome na porta {self.porta_debug} | Página: {self.driver.title}")
            return True

        except Exception as e:
            logger.error(f"❌ Falha ao conectar: {e}")
            self.conectado = False
            return False

    def desconectar(self):
        """Fecha somente a conexão do driver (NÃO fecha o Chrome do usuário)."""
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.conectado = False
        logger.info("🔌 Driver desconectado (Chrome continua aberto).")

    # ── Coleta ────────────────────────────────────────────────────────────────

    def coletar_multas(self) -> list[dict]:
        """
        Raspa todas as linhas da tabela de multas na página atual.
        Retorna lista de dicts com os dados de cada infração nova.
        """
        if not self.conectado or not self.driver:
            logger.warning("⚠️  Não conectado. Abortando coleta.")
            return []

        novas = []
        try:
            wait = WebDriverWait(self.driver, 10)

            # ── tenta localizar a tabela ──────────────────────────────────────
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, SELETORES["tabela_mensagens"])))
            except TimeoutException:
                logger.warning("⚠️  Tabela não encontrada na página atual.")
                return []

            pagina = 1
            while True:
                linhas = self.driver.find_elements(By.CSS_SELECTOR, SELETORES["linhas_tabela"])
                logger.info(f"   Página {pagina}: {len(linhas)} linha(s) encontrada(s).")

                for linha in linhas:
                    celulas = linha.find_elements(By.CSS_SELECTOR, SELETORES["celulas"])
                    if not celulas:
                        continue

                    textos = [c.text.strip() for c in celulas]
                    id_multa = self._gerar_id(textos)

                    if id_multa in self.multas_vistas:
                        continue  # já processada

                    multa = self._mapear_multa(textos, id_multa)
                    novas.append(multa)
                    self.multas_vistas.add(id_multa)

                # ── paginação ─────────────────────────────────────────────────
                try:
                    btn_prox = self.driver.find_element(By.CSS_SELECTOR, SELETORES["proxima_pagina"])
                    if btn_prox.is_displayed() and btn_prox.is_enabled():
                        btn_prox.click()
                        time.sleep(1.5)
                        pagina += 1
                    else:
                        break
                except NoSuchElementException:
                    break

        except WebDriverException as e:
            logger.error(f"❌ Erro Selenium: {e}")

        logger.info(f"✅ {len(novas)} infração(ões) nova(s) coletada(s).")
        return novas

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _gerar_id(textos: list) -> str:
        return "|".join(textos[:4])   # usa as 4 primeiras colunas como chave

    @staticmethod
    def _mapear_multa(textos: list, id_multa: str) -> dict:
        """
        Mapeia colunas para campos nomeados.
        AJUSTE os índices conforme as colunas reais do Senatran.
        """
        def safe(lst, i, default="—"):
            return lst[i] if i < len(lst) else default

        return {
            "id":            id_multa,
            "coletado_em":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data_infracao": safe(textos, 0),
            "placa":         safe(textos, 1),
            "veiculo":       safe(textos, 2),
            "infrator":      safe(textos, 3),
            "descricao":     safe(textos, 4),
            "local":         safe(textos, 5),
            "valor":         safe(textos, 6),
            "status":        safe(textos, 7),
            "raw":           textos,            # colunas brutas para diagnóstico
        }

# ─────────────────────────────────────────────────────────────────────────────
# GERENCIADOR DE DADOS
# ─────────────────────────────────────────────────────────────────────────────

class GerenciadorDados:
    def __init__(self):
        self.multas: list[dict] = []
        self._carregar()

    def _carregar(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    self.multas = json.load(f)
                logger.info(f"📂 {len(self.multas)} registro(s) carregado(s) do disco.")
            except Exception:
                self.multas = []

    def salvar(self):
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(self.multas, f, ensure_ascii=False, indent=2)

    def adicionar(self, novas: list[dict]):
        self.multas.extend(novas)
        self.salvar()

    def estatisticas_hoje(self) -> dict:
        hoje = date.today().strftime("%Y-%m-%d")
        hoje_lista = [m for m in self.multas if m.get("coletado_em", "").startswith(hoje)]
        por_hora = defaultdict(int)
        for m in hoje_lista:
            hora = m.get("coletado_em", "")[11:13]
            if hora:
                por_hora[hora] += 1

        return {
            "total_hoje":   len(hoje_lista),
            "total_geral":  len(self.multas),
            "por_hora":     dict(sorted(por_hora.items())),
            "ultima_coleta": self.multas[-1]["coletado_em"] if self.multas else "—",
        }

    def exportar_excel(self, caminho: str):
        if not self.multas:
            raise ValueError("Nenhum dado para exportar.")
        df = pd.DataFrame(self.multas)
        df.drop(columns=["id", "raw"], errors="ignore", inplace=True)
        df.to_excel(caminho, index=False)
        logger.info(f"📊 Excel exportado: {caminho}")

# ─────────────────────────────────────────────────────────────────────────────
# INTERFACE DESKTOP
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    # ── Cores e estilos ───────────────────────────────────────────────────────
    COR_BG       = "#0D1117"
    COR_PAINEL   = "#161B22"
    COR_BORDA    = "#30363D"
    COR_AZUL     = "#58A6FF"
    COR_VERDE    = "#3FB950"
    COR_AMARELO  = "#D29922"
    COR_VERMELHO = "#F85149"
    COR_TEXTO    = "#E6EDF3"
    COR_MUTED    = "#8B949E"
    FONTE_TITULO = ("Consolas", 13, "bold")
    FONTE_CORPO  = ("Consolas", 10)
    FONTE_NUM    = ("Consolas", 28, "bold")
    FONTE_LABEL  = ("Consolas", 9)

    def __init__(self):
        super().__init__()
        self.title("Maas Serviços Ltda.")
        self.geometry("1060x700")
        self.minsize(900, 600)
        self.configure(bg=self.COR_BG)
        self.resizable(True, True)

        self.coletor  = SenatranColetor()
        self.dados    = GerenciadorDados()
        self._rodando = False
        self._thread  = None

        self._build_ui()
        self._atualizar_stats()
        self._log("Sistema iniciado. Conecte ao navegador e inicie a coleta.")

    # ── Construção da UI ──────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Cabeçalho ─────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=self.COR_PAINEL, pady=10)
        header.pack(fill="x", side="top")

        tk.Label(
            header, text="⚡ Central de Monitoramento de Multas",
            font=("Consolas", 16, "bold"),
            fg=self.COR_AZUL, bg=self.COR_PAINEL
        ).pack(side="left", padx=20)

        self.lbl_status = tk.Label(
            header, text="● DESCONECTADO",
            font=self.FONTE_CORPO, fg=self.COR_VERMELHO, bg=self.COR_PAINEL
        )
        self.lbl_status.pack(side="left", padx=10)

        tk.Label(
            header,
            text=f"Horário comercial: {HORARIO_INICIO} – {HORARIO_FIM}  |  Intervalo: {INTERVALO_SEG//60} min",
            font=self.FONTE_LABEL, fg=self.COR_MUTED, bg=self.COR_PAINEL
        ).pack(side="right", padx=20)

        sep = tk.Frame(self, bg=self.COR_BORDA, height=1)
        sep.pack(fill="x")

        # ── Corpo principal ───────────────────────────────────────────────────
        corpo = tk.Frame(self, bg=self.COR_BG)
        corpo.pack(fill="both", expand=True, padx=16, pady=12)

        # Coluna esquerda: cards de estatísticas + controles
        esq = tk.Frame(corpo, bg=self.COR_BG, width=260)
        esq.pack(side="left", fill="y", padx=(0, 12))
        esq.pack_propagate(False)

        self._build_cards(esq)
        self._build_controles(esq)

        # Coluna direita: tabela + log
        dir_ = tk.Frame(corpo, bg=self.COR_BG)
        dir_.pack(side="left", fill="both", expand=True)

        self._build_tabela(dir_)
        self._build_log(dir_)

    def _card(self, pai, titulo: str, var: tk.StringVar, cor_num: str) -> tk.Frame:
        frame = tk.Frame(pai, bg=self.COR_PAINEL, padx=14, pady=10,
                         highlightbackground=self.COR_BORDA, highlightthickness=1)
        frame.pack(fill="x", pady=5)
        tk.Label(frame, text=titulo, font=self.FONTE_LABEL,
                 fg=self.COR_MUTED, bg=self.COR_PAINEL).pack(anchor="w")
        tk.Label(frame, textvariable=var, font=self.FONTE_NUM,
                 fg=cor_num, bg=self.COR_PAINEL).pack(anchor="w")
        return frame

    def _build_cards(self, pai):
        tk.Label(pai, text="ESTATÍSTICAS — HOJE",
                 font=self.FONTE_TITULO, fg=self.COR_MUTED,
                 bg=self.COR_BG).pack(anchor="w", pady=(0, 4))

        self.var_hoje  = tk.StringVar(value="0")
        self.var_geral = tk.StringVar(value="0")
        self.var_ult   = tk.StringVar(value="—")

        self._card(pai, "INFRAÇÕES HOJE",    self.var_hoje,  self.COR_AMARELO)
        self._card(pai, "TOTAL HISTÓRICO",   self.var_geral, self.COR_AZUL)
        self._card(pai, "ÚLTIMA ATUALIZAÇÃO",self.var_ult,   self.COR_VERDE)

    def _build_controles(self, pai):
        tk.Frame(pai, bg=self.COR_BORDA, height=1).pack(fill="x", pady=12)
        tk.Label(pai, text="CONTROLES", font=self.FONTE_TITULO,
                 fg=self.COR_MUTED, bg=self.COR_BG).pack(anchor="w", pady=(0, 6))

        # Porta de debug
        row = tk.Frame(pai, bg=self.COR_BG)
        row.pack(fill="x", pady=3)
        tk.Label(row, text="Porta debug:", font=self.FONTE_LABEL,
                 fg=self.COR_MUTED, bg=self.COR_BG, width=12, anchor="w").pack(side="left")
        self.entry_porta = tk.Entry(row, font=self.FONTE_CORPO, width=6,
                                    bg=self.COR_PAINEL, fg=self.COR_TEXTO,
                                    insertbackground=self.COR_TEXTO,
                                    relief="flat", bd=4)
        self.entry_porta.insert(0, "9222")
        self.entry_porta.pack(side="left")

        # Botão conectar
        self.btn_conectar = tk.Button(
            pai, text="🔌  CONECTAR AO NAVEGADOR",
            font=self.FONTE_CORPO, bg=self.COR_AZUL, fg="#000",
            activebackground="#79B8FF", relief="flat", bd=0, pady=8, cursor="hand2",
            command=self._conectar
        )
        self.btn_conectar.pack(fill="x", pady=(6, 3))

        # Botão iniciar/parar
        self.btn_iniciar = tk.Button(
            pai, text="▶  INICIAR COLETA",
            font=self.FONTE_CORPO, bg=self.COR_VERDE, fg="#000",
            activebackground="#56D364", relief="flat", bd=0, pady=8, cursor="hand2",
            command=self._toggle_coleta, state="disabled"
        )
        self.btn_iniciar.pack(fill="x", pady=3)

        # Botão coletar agora
        self.btn_agora = tk.Button(
            pai, text="⚡  COLETAR AGORA",
            font=self.FONTE_CORPO, bg=self.COR_AMARELO, fg="#000",
            activebackground="#F0B429", relief="flat", bd=0, pady=8, cursor="hand2",
            command=self._coletar_agora, state="disabled"
        )
        self.btn_agora.pack(fill="x", pady=3)

        # Botão exportar
        self.btn_export = tk.Button(
            pai, text="📊  EXPORTAR EXCEL",
            font=self.FONTE_CORPO, bg=self.COR_PAINEL, fg=self.COR_TEXTO,
            activebackground=self.COR_BORDA, relief="flat", bd=0, pady=8, cursor="hand2",
            highlightbackground=self.COR_BORDA, highlightthickness=1,
            command=self._exportar
        )
        self.btn_export.pack(fill="x", pady=3)

        # Barra de progresso do próximo ciclo
        tk.Label(pai, text="Próxima coleta em:", font=self.FONTE_LABEL,
                 fg=self.COR_MUTED, bg=self.COR_BG).pack(anchor="w", pady=(12, 2))
        self.progress = ttk.Progressbar(pai, maximum=INTERVALO_SEG, length=230)
        self.progress.pack(fill="x")
        self.lbl_contagem = tk.Label(pai, text="—", font=self.FONTE_LABEL,
                                     fg=self.COR_MUTED, bg=self.COR_BG)
        self.lbl_contagem.pack(anchor="e")

    def _build_tabela(self, pai):
        frame = tk.Frame(pai, bg=self.COR_PAINEL,
                         highlightbackground=self.COR_BORDA, highlightthickness=1)
        frame.pack(fill="both", expand=True, pady=(0, 8))

        cabecalho = tk.Frame(frame, bg=self.COR_PAINEL)
        cabecalho.pack(fill="x", padx=10, pady=8)
        tk.Label(cabecalho, text="INFRAÇÕES COLETADAS",
                 font=self.FONTE_TITULO, fg=self.COR_AZUL, bg=self.COR_PAINEL).pack(side="left")
        self.lbl_qtd = tk.Label(cabecalho, text="0 registros",
                                 font=self.FONTE_LABEL, fg=self.COR_MUTED, bg=self.COR_PAINEL)
        self.lbl_qtd.pack(side="right")

        colunas = ("coletado_em", "data_infracao", "placa", "veiculo",
                   "descricao", "local", "valor", "status")
        self.tree = ttk.Treeview(frame, columns=colunas, show="headings", height=10)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview",
                         background=self.COR_PAINEL, fieldbackground=self.COR_PAINEL,
                         foreground=self.COR_TEXTO, font=self.FONTE_LABEL,
                         rowheight=24, borderwidth=0)
        style.configure("Treeview.Heading",
                         background=self.COR_BG, foreground=self.COR_MUTED,
                         font=self.FONTE_LABEL, borderwidth=0, relief="flat")
        style.map("Treeview", background=[("selected", self.COR_AZUL)])

        widths = {"coletado_em": 130, "data_infracao": 100, "placa": 80,
                  "veiculo": 90, "descricao": 160, "local": 120,
                  "valor": 70, "status": 80}
        titulos = {"coletado_em": "Coletado em", "data_infracao": "Data Infração",
                   "placa": "Placa", "veiculo": "Veículo",
                   "descricao": "Descrição", "local": "Local",
                   "valor": "Valor", "status": "Status"}
        for col in colunas:
            self.tree.heading(col, text=titulos.get(col, col))
            self.tree.column(col, width=widths.get(col, 100), minwidth=60, anchor="w")

        sb_v = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        sb_h = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0))
        sb_v.pack(side="right", fill="y")
        sb_h.pack(side="bottom", fill="x", padx=10)

    def _build_log(self, pai):
        frame = tk.Frame(pai, bg=self.COR_PAINEL,
                         highlightbackground=self.COR_BORDA, highlightthickness=1)
        frame.pack(fill="x")
        tk.Label(frame, text="LOG DO SISTEMA", font=self.FONTE_LABEL,
                 fg=self.COR_MUTED, bg=self.COR_PAINEL, pady=4).pack(anchor="w", padx=10)
        self.txt_log = tk.Text(frame, height=5, font=("Consolas", 9),
                                bg=self.COR_BG, fg=self.COR_VERDE,
                                insertbackground=self.COR_TEXTO, relief="flat",
                                state="disabled", wrap="none")
        self.txt_log.pack(fill="x", padx=10, pady=(0, 8))

    # ── Ações ─────────────────────────────────────────────────────────────────

    def _conectar(self):
        try:
            porta = int(self.entry_porta.get())
        except ValueError:
            messagebox.showerror("Erro", "Porta inválida.")
            return

        self.coletor.porta_debug = porta
        self._log(f"Tentando conectar na porta {porta}...")

        def _thread():
            ok = self.coletor.conectar()
            self.after(0, self._pos_conectar, ok)

        threading.Thread(target=_thread, daemon=True).start()

    def _pos_conectar(self, ok: bool):
        if ok:
            self.lbl_status.config(text="● CONECTADO", fg=self.COR_VERDE)
            self.btn_iniciar.config(state="normal")
            self.btn_agora.config(state="normal")
            self._log("✅ Conectado ao Chrome com sucesso!")
        else:
            self.lbl_status.config(text="● ERRO DE CONEXÃO", fg=self.COR_VERMELHO)
            self._log(
                "❌ Não foi possível conectar. Verifique se o Chrome está aberto com:\n"
                f"    --remote-debugging-port={self.coletor.porta_debug}\n"
                "    Exemplo: chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\\ChromeDebug"
            )

    def _toggle_coleta(self):
        if self._rodando:
            self._parar()
        else:
            self._iniciar()

    def _iniciar(self):
        self._rodando = True
        self.btn_iniciar.config(text="⏹  PARAR COLETA", bg=self.COR_VERMELHO)
        self.lbl_status.config(text="● COLETANDO", fg=self.COR_VERDE)
        self._log("▶ Coleta automática iniciada (horário comercial).")
        self._thread = threading.Thread(target=self._loop_coleta, daemon=True)
        self._thread.start()

    def _parar(self):
        self._rodando = False
        self.btn_iniciar.config(text="▶  INICIAR COLETA", bg=self.COR_VERDE)
        self.lbl_status.config(text="● CONECTADO", fg=self.COR_AZUL)
        self.lbl_contagem.config(text="—")
        self.progress["value"] = 0
        self._log("⏹ Coleta pausada pelo usuário.")

    def _coletar_agora(self):
        self._log("⚡ Coleta manual iniciada...")
        threading.Thread(target=self._executar_coleta, daemon=True).start()

    def _loop_coleta(self):
        while self._rodando:
            agora = datetime.now()
            inicio = datetime.strptime(HORARIO_INICIO, "%H:%M").replace(
                year=agora.year, month=agora.month, day=agora.day)
            fim = datetime.strptime(HORARIO_FIM, "%H:%M").replace(
                year=agora.year, month=agora.month, day=agora.day)

            if inicio <= agora <= fim:
                self._executar_coleta()
                # Contagem regressiva
                for seg_restantes in range(INTERVALO_SEG, 0, -1):
                    if not self._rodando:
                        break
                    self.after(0, self._atualizar_progresso, seg_restantes)
                    time.sleep(1)
            else:
                self.after(0, self._log,
                    f"⏰ Fora do horário comercial ({HORARIO_INICIO}–{HORARIO_FIM}). "
                    "Aguardando...")
                time.sleep(60)

    def _executar_coleta(self):
        novas = self.coletor.coletar_multas()
        if novas:
            self.dados.adicionar(novas)
            self.after(0, self._atualizar_tabela, novas)
            self.after(0, self._log, f"✅ {len(novas)} nova(s) infração(ões) adicionada(s).")
        else:
            self.after(0, self._log, "🔍 Nenhuma infração nova encontrada.")
        self.after(0, self._atualizar_stats)

    def _exportar(self):
        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=f"senatran_{date.today()}.xlsx"
        )
        if not caminho:
            return
        try:
            self.dados.exportar_excel(caminho)
            messagebox.showinfo("Exportação", f"Arquivo salvo em:\n{caminho}")
            self._log(f"📊 Exportado para {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # ── Atualização de UI ─────────────────────────────────────────────────────

    def _atualizar_stats(self):
        stats = self.dados.estatisticas_hoje()
        self.var_hoje.set(str(stats["total_hoje"]))
        self.var_geral.set(str(stats["total_geral"]))
        ultima = stats["ultima_coleta"]
        self.var_ult.set(ultima[11:16] if len(ultima) >= 16 else ultima)

    def _atualizar_tabela(self, novas: list[dict]):
        for m in novas:
            self.tree.insert("", 0, values=(
                m.get("coletado_em", ""),
                m.get("data_infracao", ""),
                m.get("placa", ""),
                m.get("veiculo", ""),
                m.get("descricao", ""),
                m.get("local", ""),
                m.get("valor", ""),
                m.get("status", ""),
            ))
        total = len(self.tree.get_children())
        self.lbl_qtd.config(text=f"{total} registro(s)")

    def _atualizar_progresso(self, restam: int):
        self.progress["value"] = INTERVALO_SEG - restam
        minutos = restam // 60
        segundos = restam % 60
        self.lbl_contagem.config(text=f"{minutos:02d}:{segundos:02d}")

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        linha = f"[{ts}] {msg}\n"
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", linha)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        logger.info(msg)

    def on_close(self):
        self._rodando = False
        self.coletor.desconectar()
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# PONTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
