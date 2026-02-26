# AppTransferencia.py
# Requisitos:
#   pip install --upgrade selenium PyPDF2
#
# Execute:
#   python AppTransferencia.py

import os
import re
import time
import shutil
import threading
from pathlib import Path
from dataclasses import dataclass
from typing import Optional
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from PyPDF2 import PdfReader

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

CONFIG_FILE = "config_transferencia_app.txt"
DEFAULT_PDF_FOLDER = r"C:\Users\ll3868\Documents\Transferencias de equipamentos"


# ==========================
# Utils PDF
# ==========================
def renomear_e_mover_pdf(chamado_numero: str, tombo: str, destino_pasta: str) -> str:
    downloads = Path.home() / "Downloads"
    pdfs = list(downloads.glob("*.pdf"))
    if not pdfs:
        raise Exception("Nenhum PDF encontrado na pasta Downloads.")

    pdf_mais_recente = max(pdfs, key=os.path.getmtime)

    reader = PdfReader(str(pdf_mais_recente))
    texto_completo = ""
    for pagina in reader.pages:
        texto_completo += (pagina.extract_text() or "") + "\n"

    padrao_grp = r"Número\s*-\s*(\d+)"
    match = re.search(padrao_grp, texto_completo)
    if not match:
        raise Exception("Número da GRP não encontrado no PDF.")

    numero_grp = match.group(1)
    from datetime import datetime

    chamado_numero = (chamado_numero or "").strip() or "SEM_CHAMADO"
    tombo = (tombo or "").strip() or "SEM_TOMBO"
    data_hoje = datetime.now().strftime("%Y-%m-%d")
    novo_nome = f"CH-{chamado_numero}_T-{tombo}_GRP-{numero_grp}_{data_hoje}.pdf"
    os.makedirs(destino_pasta, exist_ok=True)
    caminho_destino = os.path.join(destino_pasta, novo_nome)
    shutil.move(str(pdf_mais_recente), caminho_destino)
    return caminho_destino


def esperar_download_pdf(download_dir: Path, start_time: float, timeout: int = 60) -> Path:
    """
    Espera aparecer um PDF novo em Downloads após start_time e aguarda terminar (sem .crdownload).
    Retorna o caminho do PDF final.
    """
    deadline = time.time() + timeout

    while time.time() < deadline:
        # ainda baixando
        if list(download_dir.glob("*.crdownload")):
            time.sleep(0.4)
            continue

        pdfs = [p for p in download_dir.glob("*.pdf") if p.stat().st_mtime >= start_time]
        if pdfs:
            return max(pdfs, key=lambda p: p.stat().st_mtime)

        time.sleep(0.4)

    raise TimeoutException("Timeout aguardando o download do PDF em Downloads.")


# ==========================
# Dados do formulário
# ==========================
@dataclass
class DadosTransferencia:
    login: str
    senha: str
    origem: str
    destino: str
    tombo: str
    chamado: str = ""
    pasta_pdf: str = DEFAULT_PDF_FOLDER
    manter_chrome_aberto: bool = False


# ==========================
# Motor Selenium
# ==========================
class TransferenciaGRP:
    def __init__(self, stop_event: threading.Event, log_fn):
        self.stop_event = stop_event
        self.log = log_fn

    def _check_stop(self):
        if self.stop_event.is_set():
            raise Exception("Processo interrompido pelo usuário.")

    def _make_driver(self, manter_aberto: bool) -> webdriver.Chrome:
        options = Options()
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-popup-blocking")

        # ✅ Preferir download direto do PDF (sem viewer)
        prefs = {
            "download.default_directory": str(Path.home() / "Downloads"),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
        }
        options.add_experimental_option("prefs", prefs)

        if manter_aberto:
            options.add_experimental_option("detach", True)

        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)

    def _login(self, driver: webdriver.Chrome, wait: WebDriverWait, login: str, senha: str):
        self.log("Abrindo GRP...")
        driver.get("https://tjse.thema.inf.br/grp/home.faces")

        wait.until(EC.presence_of_element_located((By.ID, "loginForm:usuario")))
        driver.find_element(By.ID, "loginForm:usuario").send_keys(login)
        driver.find_element(By.ID, "loginForm:senha").send_keys(senha)

        self._check_stop()
        btn = wait.until(EC.element_to_be_clickable((By.ID, "loginForm:login")))
        driver.execute_script("arguments[0].click();", btn)

        self._check_stop()
        btn_admin = wait.until(EC.element_to_be_clickable((By.ID, "formAdministracao:administracao_1")))
        btn_admin.click()

        self.log("Login OK.")

    def _ir_transferencia(self, wait: WebDriverWait):
        self._check_stop()
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Transferência de Bens']/ancestor::a")))
        btn.click()
        time.sleep(0.8)

    def _criar(self, wait: WebDriverWait):
        self._check_stop()
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Criar']")))
        btn.click()
        time.sleep(0.8)

    def _salvar(self, driver: webdriver.Chrome, wait: WebDriverWait, origem: str, destino: str):
        self._check_stop()
        driver.find_element(By.ID, "form_transferenciaBemM:codigoLocalOrigem:field").send_keys(origem)
        driver.find_element(By.ID, "form_transferenciaBemM:codigoLocalDestino:field").send_keys(destino)

        btn_salvar = wait.until(EC.element_to_be_clickable((By.ID, "form_transferenciaBemM:cmdl_salvar")))
        btn_salvar.click()
        time.sleep(0.8)

        btn_fechar = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Fechar Mensagem']")))
        btn_fechar.click()
        time.sleep(0.4)

    def _abrir_aba_bens(self, wait: WebDriverWait):
        self._check_stop()
        btn = wait.until(EC.element_to_be_clickable((By.ID, "form_transferenciaBemM:aba_670280:header:inactive")))
        btn.click()
        time.sleep(0.7)

    def _fechar_todas_msg(self, wait: WebDriverWait):
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Fechar todas as mensagens']")))
            btn.click()
        except Exception:
            pass

    def _inserir_bem(self, driver: webdriver.Chrome, wait: WebDriverWait, tombo: str):
        self._check_stop()
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Inserir Bens']")))
        btn.click()
        time.sleep(0.8)

        campo = driver.find_element(By.ID, "form_transferenciaBemM:campoCodigoBem:fieldNumerico1:field")
        campo.send_keys(tombo)
        campo.send_keys(Keys.ENTER)
        time.sleep(0.8)

        self._check_stop()
        btn_linha = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "form_transferenciaBemM:dataTableModalInsereBens:0:divDataAquisicaoItemTransferenciaBem")
            )
        )
        btn_linha.click()
        time.sleep(0.4)

        self._check_stop()
        btn_inserir = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@id, 'msgModalInserirBensItemTransferencia')]//input[@value='Inserir']")
            )
        )
        btn_inserir.click()
        time.sleep(0.8)

        self._fechar_todas_msg(wait)

    def _encerrar(self, driver: webdriver.Chrome, wait: WebDriverWait):
        self._check_stop()
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Encerrar']")))
        btn.click()
        time.sleep(0.8)

        wait.until(EC.alert_is_present())
        driver.switch_to.alert.accept()

        btn_fechar = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='Fechar Mensagem']")))
        btn_fechar.click()
        time.sleep(0.6)

    # ==========================
    # PDF (igual ao ChamadoTransferenciasBot.py)
    # ==========================
    def _clicar_pdf_estilo_bot(self, driver: webdriver.Chrome, wait: WebDriverWait):
        """
        Procedimento igual ao ChamadoTransferenciasBot.py:
          - hover no elemento com id contendo '3z_1_2_1_2_3_label'
          - clicar no item PDF: rf-ddm-itm + linkRelatorio + span 'PDF'
        """
        self._check_stop()

        xpath_label = "//*[contains(@id, '3z_1_2_1_2_3_label')]"
        xpath_pdf = "//div[contains(@class, 'rf-ddm-itm') and contains(@class, 'linkRelatorio') and .//span[text()='PDF']]"

        label = wait.until(EC.presence_of_element_located((By.XPATH, xpath_label)))
        ActionChains(driver).move_to_element(label).perform()
        time.sleep(0.6)

        pdf_item = wait.until(EC.presence_of_element_located((By.XPATH, xpath_pdf)))
        driver.execute_script("arguments[0].click();", pdf_item)

    def _clicar_download_apryse_shadow(self, driver: webdriver.Chrome, wait: WebDriverWait):
        """
        Fallback igual ao bot: clica no botão download dentro do <apryse-webviewer> shadowRoot.
        """
        self._check_stop()

        wait.until(EC.presence_of_element_located((By.TAG_NAME, "apryse-webviewer")))
        driver.execute_script(
            """
            const viewer = document.querySelector("apryse-webviewer");
            if (!viewer) { return; }
            const shadow = viewer.shadowRoot;
            if (!shadow) { return; }

            const btn = shadow.querySelector('[data-element="downloadButton"]');
            if (btn) {
                btn.scrollIntoView({behavior: "smooth", block: "center"});
                btn.click();
            }
            """
        )

    def _gerar_pdf(self, driver: webdriver.Chrome, wait: WebDriverWait):
        """
        1) Clica PDF do jeito do BOT (hover label + click PDF).
        2) Tenta detectar download direto.
        3) Se não baixar, tenta clicar no download do Apryse e aguarda.
        """
        self.log("Gerando PDF (procedimento do bot: hover na impressora -> PDF)...")
        self._check_stop()

        start_time = time.time()
        download_dir = Path.home() / "Downloads"

        # 1) clicar no PDF do jeito do bot (com tentativas)
        last_err = None
        for tentativa in range(1, 6):
            self._check_stop()
            try:
                self.log(f"Abrindo menu/selecionando PDF (tentativa {tentativa}/5)...")
                self._clicar_pdf_estilo_bot(driver, wait)
                break
            except (TimeoutException, WebDriverException) as e:
                last_err = e
                time.sleep(0.8)
        else:
            raise Exception(f"Não consegui clicar em PDF (menu do bot). Último erro: {last_err}")

        # 2) esperar baixar direto (se o Chrome respeitar always_open_pdf_externally)
        try:
            pdf_path = esperar_download_pdf(download_dir, start_time, timeout=20)
            self.log(f"PDF baixado direto em Downloads: {pdf_path.name}")
            return
        except TimeoutException:
            self.log("Não baixou direto. Tentando download pelo viewer (Apryse)...")

        # 3) fallback: clicar download no Apryse e esperar baixar
        try:
            self._clicar_download_apryse_shadow(driver, wait)
            pdf_path = esperar_download_pdf(download_dir, start_time, timeout=60)
            self.log(f"PDF baixado via Apryse em Downloads: {pdf_path.name}")
        except Exception as e:
            raise Exception(f"Falha ao baixar PDF via Apryse: {e}")

    def executar(self, dados: DadosTransferencia) -> str:
        driver = None
        try:
            driver = self._make_driver(dados.manter_chrome_aberto)
            wait = WebDriverWait(driver, 25)

            self._login(driver, wait, dados.login, dados.senha)

            self.log("Abrindo módulo Transferência de Bens...")
            self._ir_transferencia(wait)

            self.log("Criando nova transferência...")
            self._criar(wait)

            self.log("Preenchendo origem/destino e salvando...")
            self._salvar(driver, wait, dados.origem, dados.destino)

            self.log("Abrindo aba de bens...")
            self._abrir_aba_bens(wait)

            self.log(f"Inserindo tombo {dados.tombo} ...")
            self._inserir_bem(driver, wait, dados.tombo)

            self.log("Encerrando transferência...")
            self._encerrar(driver, wait)

            self._gerar_pdf(driver, wait)

            self.log("Renomeando e movendo PDF...")
            caminho = renomear_e_mover_pdf(dados.chamado, dados.tombo, dados.pasta_pdf)

            self.log(f"✅ Concluído! PDF salvo em: {caminho}")
            return caminho

        finally:
            try:
                if driver and not dados.manter_chrome_aberto:
                    driver.quit()
            except Exception:
                pass


# ==========================
# UI
# ==========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação - Transferência de Equipamentos (GRP)")
        self.geometry("860x590")
        self.minsize(780, 540)

        self.stop_event = threading.Event()
        self.worker: Optional[threading.Thread] = None

        self._build()
        self._load_config_silent()

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Transferência GRP", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(root, text="1 tombo por vez • Sem Telegram • Log em tempo real", font=("Segoe UI", 10)).pack(
            anchor="w", pady=(4, 12)
        )

        form = ttk.LabelFrame(root, text="Dados", padding=12)
        form.pack(fill="x")

        self.var_login = tk.StringVar()
        self.var_senha = tk.StringVar()
        self.var_origem = tk.StringVar()
        self.var_destino = tk.StringVar()
        self.var_tombo = tk.StringVar()
        self.var_chamado = tk.StringVar()
        self.var_pasta = tk.StringVar(value=DEFAULT_PDF_FOLDER)
        self.var_detach = tk.BooleanVar(value=False)

        form.columnconfigure(1, weight=1)

        def row(r, label, var, show=None, hint=None):
            ttk.Label(form, text=label).grid(row=r, column=0, sticky="w", padx=(0, 10), pady=6)
            e = ttk.Entry(form, textvariable=var, show=show)
            e.grid(row=r, column=1, sticky="ew", pady=6)
            if hint:
                ttk.Label(form, text=hint, foreground="#666").grid(row=r, column=2, sticky="w", padx=(10, 0))
            return e

        row(0, "Login", self.var_login, hint="usuário GRP")
        row(1, "Senha", self.var_senha, show="•", hint="não é salva por padrão")
        row(2, "Origem", self.var_origem, hint="código local origem")
        row(3, "Destino", self.var_destino, hint="código local destino")
        row(4, "Tombo", self.var_tombo, hint="patrimônio/tombo")
        row(5, "Chamado (opcional)", self.var_chamado, hint="vai no nome do PDF")

        ttk.Label(form, text="Pasta PDF").grid(row=6, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(form, textvariable=self.var_pasta).grid(row=6, column=1, sticky="ew", pady=6)
        ttk.Label(form, text="(padrão já configurado)", foreground="#666").grid(row=6, column=2, sticky="w", padx=(10, 0))

        ttk.Checkbutton(form, text="Modo debug (manter Chrome aberto)", variable=self.var_detach).grid(
            row=7, column=1, sticky="w", pady=(4, 0)
        )

        actions = ttk.Frame(root)
        actions.pack(fill="x", pady=(12, 0))

        self.btn_start = ttk.Button(actions, text="▶ Iniciar", command=self._start)
        self.btn_start.pack(side="left")

        self.btn_stop = ttk.Button(actions, text="⏹ Parar", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=8)

        ttk.Button(actions, text="💾 Salvar config", command=self._save_config).pack(side="right")

        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill="x", pady=(12, 8))

        ttk.Label(root, text="Log").pack(anchor="w")
        self.txt = tk.Text(root, height=13, wrap="word")
        self.txt.pack(fill="both", expand=True)
        self.txt.configure(state="disabled")

        self._log("Pronto. Preencha os dados e clique em Iniciar.")

    def _log(self, msg: str):
        self.txt.configure(state="normal")
        self.txt.insert("end", f"{time.strftime('%H:%M:%S')}  {msg}\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _thread_log(self, msg: str):
        self.after(0, lambda: self._log(msg))

    def _validate(self) -> Optional[str]:
        if not self.var_login.get().strip():
            return "Informe o Login."
        if not self.var_senha.get().strip():
            return "Informe a Senha."
        if not self.var_origem.get().strip():
            return "Informe a Origem."
        if not self.var_destino.get().strip():
            return "Informe o Destino."
        if not self.var_tombo.get().strip():
            return "Informe o Tombo."
        if not self.var_pasta.get().strip():
            return "Informe a pasta de destino do PDF."
        return None

    def _lock(self, running: bool):
        self.btn_start.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")

    def _start(self):
        err = self._validate()
        if err:
            messagebox.showwarning("Atenção", err)
            return

        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Em execução", "Já existe uma execução em andamento.")
            return

        self.stop_event.clear()
        self._lock(True)
        self.progress.start(10)

        dados = DadosTransferencia(
            login=self.var_login.get().strip(),
            senha=self.var_senha.get().strip(),
            origem=self.var_origem.get().strip(),
            destino=self.var_destino.get().strip(),
            tombo=self.var_tombo.get().strip(),
            chamado=self.var_chamado.get().strip(),
            pasta_pdf=self.var_pasta.get().strip(),
            manter_chrome_aberto=bool(self.var_detach.get()),
        )

        self._log("Iniciando automação...")

        def run():
            try:
                engine = TransferenciaGRP(self.stop_event, self._thread_log)
                engine.executar(dados)
            except Exception as e:
                self._thread_log(f"❌ Erro: {e}")
            finally:
                self.after(0, self._finish)

        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def _finish(self):
        self.progress.stop()
        self._lock(False)
        self._log("Processo finalizado.")

    def _stop(self):
        self.stop_event.set()
        self._log("Solicitado: parar execução (vai interromper no próximo ponto seguro).")

    def _save_config(self):
        data = "\n".join(
            [
                f"login={self.var_login.get().strip()}",
                f"origem={self.var_origem.get().strip()}",
                f"destino={self.var_destino.get().strip()}",
                f"pasta={self.var_pasta.get().strip()}",
            ]
        )
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(data)
        self._log(f"Config salva em: {CONFIG_FILE}")

    def _load_config_silent(self):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
            kv = {}
            for line in lines:
                if "=" in line:
                    k, v = line.split("=", 1)
                    kv[k.strip()] = v.strip()
            if "login" in kv:
                self.var_login.set(kv["login"])
            if "origem" in kv:
                self.var_origem.set(kv["origem"])
            if "destino" in kv:
                self.var_destino.set(kv["destino"])
            if "pasta" in kv:
                self.var_pasta.set(kv["pasta"])
            self._log("Config carregada automaticamente.")
        except Exception:
            pass


if __name__ == "__main__":
    app = App()
    app.mainloop()