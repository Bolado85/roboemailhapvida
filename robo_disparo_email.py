"""
Hapvida — Robô de Disparo de E-mail (paralelo)
Envia e-mails personalizados via Gmail SMTP para uma lista de clientes,
usando um template HTML com placeholders {{COLUNA}}.

Uso normal (clique duplo no iniciar_envio.bat, ou linha de comando):
    python robo_disparo_email.py --teste seuemail@gmail.com   (envia 1 e-mail de teste)
    python robo_disparo_email.py                              (envia para todos os pendentes)
    python robo_disparo_email.py --workers 3 --delay 2 --limite 400
"""

import argparse
import csv
import os
import re
import smtplib
import ssl
import sys
import threading
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote as url_quote

# ── Corrige encoding do console no Windows ────────────────────
# Mesmo problema que já corrigimos no robô de CPF: sem isso, qualquer
# emoji no print() quebra com UnicodeEncodeError em consoles Windows.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════
# ⚙️  CONFIGURAÇÕES — edite aqui antes de usar
# ══════════════════════════════════════════════════════════════
EMAIL_REMETENTE   = "matheusrodriguesgtel2@gmail.com"        # conta Gmail que vai enviar
SENHA_APP         = "ejjv jseq bxde ykzb"       # Senha de App do Gmail (NÃO é a senha normal!)
NOME_REMETENTE    = "Hapvida Odonto"            # nome que aparece no campo "De:"

ASSUNTO           = "Oferta Imperdível: Plano Odontológico Hapvida!" # assunto do e-mail (aceita {{PLACEHOLDERS}})

ARQUIVO_CLIENTES  = "clientes.csv"              # CSV ou Excel com os dados + e-mail
ARQUIVO_TEMPLATE  = "template_email.html"       # template HTML do corpo do e-mail
ANEXOS = []  # sem anexos — o e-mail já traz tudo no próprio corpo (estilo landing page)

WORKERS             = 2     # conexões SMTP simultâneas — Gmail não gosta de muitas, mantenha baixo (2-4)
DELAY_ENTRE_ENVIOS  = 2     # segundos de pausa entre envios de cada worker (evita ser marcado como spam/robô)
LIMITE_POR_EXECUCAO = 450   # não passe de ~450-500/dia no Gmail pessoal (2000/dia no Google Workspace)

ARQUIVO_LOG       = "log_envios.csv"
# ══════════════════════════════════════════════════════════════

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

# ── Mapeamento automático de colunas ───────────────────────────
# O robô entende variações de nome de coluna, então funciona mesmo se
# o arquivo vier com "E-mail", "Email do Cliente", "CPF Buscado" etc.
# Qualquer coluna extra que não estiver aqui continua disponível no
# template pelo próprio nome (em maiúsculas), então você pode usar
# {{QUALQUER_COLUNA}} sem precisar editar este código.
COLUNAS_MAPA = {
    "NOME":     ["NOME", "NOME DO CLIENTE", "CLIENTE", "NOME_CLIENTE"],
    "EMAIL":    ["EMAIL", "E-MAIL", "E MAIL", "EMAIL DO CLIENTE"],
    "CPF":      ["CPF", "CPF BUSCADO", "CPF DO CLIENTE"],
    "PLANO":    ["PLANO", "PLANO (REPIQUE)", "COD PLANO", "CÓDIGO DO PLANO"],
    "CONTRATO": ["CONTRATO"],
    "INICIO":   ["INÍCIO", "INICIO", "DATA INICIO", "DATA DE INICIO"],
}

RE_EMAIL = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')

lock_log   = threading.Lock()
lock_print = threading.Lock()
lock_stat  = threading.Lock()
stop_event = threading.Event()

contadores = {"enviados": 0, "erros": 0, "pulados": 0}
workers_status = {}


# ── Utilitários ─────────────────────────────────────────────────
def log(msg, tipo="info"):
    icones = {"ok": "✅", "err": "❌", "warn": "⚠️", "info": "  "}
    with lock_print:
        print(f"  {icones.get(tipo, '  ')} {msg}", flush=True)


def emitir_status(total, workers_total):
    """Linha padronizada — compatível com o mesmo formato usado no robô
    de CPF, caso você queira plugar um dashboard depois."""
    with lock_stat:
        ativos = sum(1 for s in workers_status.values() if s == 'ativo')
    print(f">>> STATUS | total={total} enviados={contadores['enviados']} "
          f"erros={contadores['erros']} pulados={contadores['pulados']} "
          f"workers_ativos={ativos} workers_total={workers_total}", flush=True)


def limpar_nome_coluna(c) -> str:
    return re.sub(r'\s+', ' ', str(c)).strip().upper()


def mapear_colunas(colunas_originais):
    limpas = {limpar_nome_coluna(c): c for c in colunas_originais}
    mapa = {}
    for padrao, variantes in COLUNAS_MAPA.items():
        for v in variantes:
            if v in limpas:
                mapa[limpas[v]] = padrao
                break
    return mapa


def carregar_clientes(caminho):
    """Lê CSV ou Excel e devolve lista de dicts com colunas já mapeadas
    (ex: {'NOME': 'Maria', 'EMAIL': 'maria@x.com', 'CPF': '...', ...})."""
    linhas = []

    if caminho.lower().endswith((".xlsx", ".xls")):
        import pandas as pd
        df = pd.read_excel(caminho, dtype=str).fillna("")
        mapa = mapear_colunas(list(df.columns))
        for _, row in df.iterrows():
            d = {}
            for col_orig, val in row.items():
                chave = mapa.get(col_orig, limpar_nome_coluna(col_orig))
                d[chave] = str(val).strip()
            linhas.append(d)
    else:
        # Detecta a assinatura de bytes: se o arquivo ".csv" na verdade for um
        # Excel (.xlsx) renomeado/exportado errado, ele começa com "PK" (ZIP).
        with open(caminho, "rb") as fbin:
            assinatura = fbin.read(4)
        if assinatura[:2] == b"PK":
            import pandas as pd
            df = pd.read_excel(caminho, dtype=str).fillna("")
            mapa = mapear_colunas(list(df.columns))
            for _, row in df.iterrows():
                d = {}
                for col_orig, val in row.items():
                    chave = mapa.get(col_orig, limpar_nome_coluna(col_orig))
                    d[chave] = str(val).strip()
                linhas.append(d)
            return linhas

        # Tenta UTF-8 primeiro; se o Excel tiver salvo em ANSI/Windows-1252
        # (muito comum ao usar "Salvar como > CSV" no Windows), cai para
        # cp1252 automaticamente em vez de travar com UnicodeDecodeError.
        texto = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(caminho, "r", encoding=enc) as f:
                    texto = f.read()
                break
            except UnicodeDecodeError:
                continue
        if texto is None:
            raise ValueError(f"Não foi possível decodificar '{caminho}' com utf-8, cp1252 ou latin-1.")

        amostra = texto[:2048]
        sep = ";" if amostra.count(";") > amostra.count(",") else ","
        reader = csv.DictReader(texto.splitlines(), delimiter=sep)
        mapa = mapear_colunas(reader.fieldnames or [])
        for row in reader:
            d = {}
            for col_orig, val in row.items():
                chave = mapa.get(col_orig, limpar_nome_coluna(col_orig))
                d[chave] = (val or "").strip()
            linhas.append(d)

    return linhas


def carregar_template(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        return f.read()


def renderizar_template(template, dados):
    """Substitui {{CHAVE}} pelos valores do dicionário (não faz distinção
    de maiúsculas/minúsculas dentro das chaves)."""
    def substituir(m):
        chave = m.group(1).strip().upper()
        return dados.get(chave, "")
    return re.sub(r"\{\{\s*(\w+)\s*\}\}", substituir, template)


def carregar_ja_enviados():
    enviados = set()
    if not os.path.exists(ARQUIVO_LOG):
        return enviados
    try:
        with open(ARQUIVO_LOG, "r", encoding="utf-8-sig") as f:
            r = csv.DictReader(f, delimiter=";")
            for row in r:
                if row.get("Status") == "enviado":
                    enviados.add((row.get("Email") or "").strip().lower())
    except Exception:
        pass
    return enviados


def gravar_log(email, nome, status, detalhe=""):
    novo = not os.path.exists(ARQUIVO_LOG)
    with lock_log:
        with open(ARQUIVO_LOG, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            if novo:
                w.writerow(["Email", "Nome", "Status", "Detalhe", "DataHora"])
            w.writerow([email, nome, status, detalhe, time.strftime("%Y-%m-%d %H:%M:%S")])


def montar_email(destinatario, dados, template_html, assunto):
    msg = MIMEMultipart("alternative")

    # Versão do nome já codificada para uso seguro em URLs (ex: link do WhatsApp)
    dados_render = dict(dados)
    dados_render["NOME_URL"] = url_quote(dados.get("NOME", "Cliente") or "Cliente")

    msg["Subject"] = renderizar_template(assunto, dados_render)
    msg["From"] = f"{NOME_REMETENTE} <{EMAIL_REMETENTE}>"
    msg["To"] = destinatario

    html_final = renderizar_template(template_html, dados_render)
    msg.attach(MIMEText(html_final, "html", "utf-8"))

    for caminho_anexo in ANEXOS:
        if caminho_anexo and os.path.exists(caminho_anexo):
            with open(caminho_anexo, "rb") as f:
                parte = MIMEApplication(f.read(), Name=os.path.basename(caminho_anexo))
            parte["Content-Disposition"] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
            msg.attach(parte)
        elif caminho_anexo:
            log(f"Anexo não encontrado: {caminho_anexo}", "warn")

    return msg


def conectar_smtp():
    contexto = ssl.create_default_context()
    servidor = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
    servidor.starttls(context=contexto)
    servidor.login(EMAIL_REMETENTE, SENHA_APP)
    return servidor


# ── Worker ───────────────────────────────────────────────────────
def worker(wid, clientes_atribuidos, template_html, delay, total, workers_total):
    tag = f"[Worker {wid:02d}]"
    workers_status[wid] = 'ativo'

    try:
        servidor = conectar_smtp()
        log(f"{tag} Conectado ao Gmail com sucesso.", "ok")
    except Exception as e:
        workers_status[wid] = 'erro'
        log(f"{tag} FALHA ao conectar/logar no Gmail: {e}", "err")
        log("Verifique EMAIL_REMETENTE e SENHA_APP no topo do script.", "err")
        return

    for dados in clientes_atribuidos:
        if stop_event.is_set():
            break

        email_dest = dados.get("EMAIL", "").strip()
        nome_dest  = dados.get("NOME") or email_dest or "(sem nome)"

        if not email_dest or not RE_EMAIL.match(email_dest):
            log(f"{tag} E-mail inválido/vazio para {nome_dest}, pulando.", "warn")
            gravar_log(email_dest, nome_dest, "pulado", "email invalido ou vazio")
            with lock_stat:
                contadores["pulados"] += 1
            emitir_status(total, workers_total)
            continue

        try:
            msg = montar_email(email_dest, dados, template_html, ASSUNTO)
            servidor.sendmail(EMAIL_REMETENTE, email_dest, msg.as_string())
            gravar_log(email_dest, nome_dest, "enviado")
            log(f"{tag} Enviado para {nome_dest} <{email_dest}>", "ok")
            with lock_stat:
                contadores["enviados"] += 1
        except Exception as e:
            gravar_log(email_dest, nome_dest, "erro", str(e)[:150])
            log(f"{tag} ERRO ao enviar para {email_dest}: {e}", "err")
            with lock_stat:
                contadores["erros"] += 1
            # Reconecta em caso de erro de conexão (comum após muitos envios)
            try:
                servidor.quit()
            except Exception:
                pass
            try:
                servidor = conectar_smtp()
            except Exception as e2:
                log(f"{tag} Falha ao reconectar: {e2}", "err")

        emitir_status(total, workers_total)

        if stop_event.wait(delay):
            break

    workers_status[wid] = 'fim'
    try:
        servidor.quit()
    except Exception:
        pass
    log(f"{tag} Encerrado.", "info")


def dividir(lst, n):
    return [lst[i::n] for i in range(n)]


def enviar_teste(destino):
    """Envia 1 e-mail de exemplo com dados fictícios, sem tocar na lista real."""
    template_html = carregar_template(ARQUIVO_TEMPLATE)
    dados_fake = {
        "NOME": "Cliente Teste", "EMAIL": destino, "CPF": "000.000.000-00",
        "PLANO": "0000 (exemplo)", "CONTRATO": "CONTRATO-TESTE", "INICIO": "01/01/2026",
    }
    servidor = conectar_smtp()
    msg = montar_email(destino, dados_fake, template_html, ASSUNTO)
    servidor.sendmail(EMAIL_REMETENTE, destino, msg.as_string())
    servidor.quit()
    print(f"\n  ✅ E-mail de teste enviado para {destino}! Confira a caixa de entrada (e o spam).\n")


# ── Main ─────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Robô de disparo de e-mail Hapvida")
    parser.add_argument("--workers", type=int, default=WORKERS)
    parser.add_argument("--delay",   type=int, default=DELAY_ENTRE_ENVIOS)
    parser.add_argument("--limite",  type=int, default=LIMITE_POR_EXECUCAO)
    parser.add_argument("--teste",   type=str, default=None,
                         help="Envia 1 e-mail de teste para o endereço informado e sai.")
    args = parser.parse_args()

    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║   HAPVIDA — ROBÔ DE DISPARO DE E-MAIL                ║")
    print("  ╚══════════════════════════════════════════════════════╝")
    print()

    if not EMAIL_REMETENTE or "seuemail" in EMAIL_REMETENTE or "xxxx" in SENHA_APP:
        print("  [ERRO] Configure EMAIL_REMETENTE e SENHA_APP no topo do arquivo antes de usar!")
        print("  Veja as instruções de 'Senha de App' no README.")
        input("  Pressione Enter para sair...")
        return

    if args.teste:
        print(f"  Enviando e-mail de TESTE para: {args.teste}\n")
        try:
            enviar_teste(args.teste)
        except Exception as e:
            print(f"  ❌ Erro ao enviar teste: {e}")
        input("  Pressione Enter para sair...")
        return

    if not os.path.exists(ARQUIVO_CLIENTES):
        print(f"  ❌ Arquivo '{ARQUIVO_CLIENTES}' não encontrado!")
        input("  Enter para sair..."); return

    if not os.path.exists(ARQUIVO_TEMPLATE):
        print(f"  ❌ Template '{ARQUIVO_TEMPLATE}' não encontrado!")
        input("  Enter para sair..."); return

    clientes = carregar_clientes(ARQUIVO_CLIENTES)
    template_html = carregar_template(ARQUIVO_TEMPLATE)

    if not clientes:
        print("  ❌ Nenhum cliente encontrado no arquivo!")
        input("  Enter para sair..."); return

    ja_enviados = carregar_ja_enviados()

    # Filtra pendentes e remove duplicados de e-mail dentro da própria lista
    vistos, pendentes = set(), []
    for c in clientes:
        e = c.get("EMAIL", "").strip().lower()
        if not e or e in ja_enviados or e in vistos:
            continue
        vistos.add(e)
        pendentes.append(c)

    if args.limite and len(pendentes) > args.limite:
        print(f"  ⚠️  Limitando a {args.limite} envios nesta execução (de {len(pendentes)} pendentes).")
        pendentes = pendentes[:args.limite]

    total = len(clientes)
    n_workers = min(args.workers, len(pendentes)) if pendentes else 1

    print(f"  📋 Total de clientes:     {total}")
    print(f"  ✅ Já enviados antes:     {len(ja_enviados)}")
    print(f"  ⏳ Pendentes agora:       {len(pendentes)}")
    print(f"  🔧 Workers (conexões):    {n_workers}")
    print(f"  ⏱️  Delay entre envios:   {args.delay}s")
    print()

    if not pendentes:
        print("  ✅ Todos os e-mails já foram enviados!")
        input("  Enter para sair..."); return

    print("  Ctrl+C para parar com segurança — o progresso fica salvo em log_envios.csv")
    print("  e a próxima execução pula quem já recebeu.")
    print()
    time.sleep(1.5)

    lotes = dividir(pendentes, n_workers)
    for wid in range(1, n_workers + 1):
        workers_status[wid] = 'pendente'

    threads = [
        threading.Thread(target=worker, args=(
            wid, lote, template_html, args.delay, total, n_workers
        ), daemon=True)
        for wid, lote in enumerate(lotes, start=1)
    ]

    try:
        for t in threads: t.start()
        for t in threads: t.join()
    except KeyboardInterrupt:
        print("\n  ⚠️  Interrompido! Progresso salvo em log_envios.csv")
        stop_event.set()
        for t in threads: t.join(timeout=10)

    print()
    print("  ╔══════════════════════════════════════════╗")
    print("  ║  CONCLUÍDO!                              ║")
    print(f"  ║  Enviados : {contadores['enviados']:<28}║")
    print(f"  ║  Erros    : {contadores['erros']:<28}║")
    print(f"  ║  Pulados  : {contadores['pulados']:<28}║")
    print("  ╚══════════════════════════════════════════╝")
    print()


if __name__ == "__main__":
    main()
