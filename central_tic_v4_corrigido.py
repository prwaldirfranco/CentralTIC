#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
===============================================================================
CENTRAL DE ATENDIMENTO TIC - PETROBRAS
Versão 4.0
Backend principal
===============================================================================

Autor: Waldir Franco Mesquita Junior
Ano: 2026

Arquitetura:
- Python puro
- http.server
- socketserver
- Sem Flask
- Sem Django
- Sem bibliotecas externas
- API REST JSON
- Autenticação por sessão
- Controle de permissões
- PBKDF2-HMAC-SHA256
- Controle de sessão
- Leitura nativa de XLSX
- Scripts de atendimento
- Knowledge Base
- Aplicações
- Dashboard
- Administração
- Logs
- Auditoria
- CRUD de scripts
- CRUD de Knowledge Base
- Gerenciamento de usuários
- Proteções básicas HTTP
- Limite de requisição
- Tratamento de erros
- Estrutura preparada para evolução da CentralTIC

===============================================================================
"""

import http.server
import socketserver
import threading
import socket
import json
import uuid
import secrets
import hashlib
import hmac
import logging
import os
import zipfile
import xml.etree.ElementTree as ET
import mimetypes

from pathlib import Path
from datetime import datetime, timedelta
from urllib.parse import urlparse, unquote, parse_qs
from typing import Dict, List, Optional, Tuple, Any


# =============================================================================
# CONFIGURAÇÃO
# =============================================================================

VERSAO = "4.0.0"

NOME_SISTEMA = "Central de Atendimento TIC"

PORTA_PREFERIDA = int(
    os.getenv("PORT", "8895")
)

HOST = os.getenv(
    "HOST",
    "0.0.0.0"
)

# Em hospedagem HTTPS (Render etc.), marca cookie Secure
COOKIE_SECURE = os.getenv("COOKIE_SECURE", "").lower() in ("1", "true", "yes") or bool(os.getenv("RENDER"))


def flag_cookie_secure() -> str:
    return "; Secure" if COOKIE_SECURE else ""


BASE_DIR = Path(
    __file__
).resolve().parent

DADOS_DIR = BASE_DIR / "dados"

TEMPLATES_DIR = BASE_DIR / "templates"

STATIC_DIR = BASE_DIR / "static"

CSS_DIR = STATIC_DIR / "css"

JS_DIR = STATIC_DIR / "js"

IMG_DIR = STATIC_DIR / "img"

LOGS_DIR = BASE_DIR / "logs"

BACKUP_DIR = BASE_DIR / "backup"


# =============================================================================
# ARQUIVOS
# =============================================================================

EXCEL_FILE = DADOS_DIR / "aplicacoes.xlsx"

USUARIOS_FILE = DADOS_DIR / "usuarios.json"

SCRIPTS_FILE = DADOS_DIR / "scripts.json"

MODELO_DIALOGO_FILE = DADOS_DIR / "modelo_dialogo.json"

MODELO_DIALOGO_URL_PADRAO = (
    "https://petrobras.service-now.com/cs?id=kb_article_view&sysparm_article=KB0029502"
)

KB_FILE = DADOS_DIR / "kb.json"

AUDITORIA_FILE = LOGS_DIR / "auditoria.json"


# =============================================================================
# SEGURANÇA
# =============================================================================

SESSION_TIMEOUT = 3600

MAX_BODY_SIZE = 2 * 1024 * 1024

MAX_UPLOAD_SIZE = 15 * 1024 * 1024  # 15 MB (planilha de aplicações)

MAX_TITULO = 100

MAX_TEXTO = 5000

MAX_CATEGORIA = 80

MAX_EMOJI = 8

MAX_USUARIO = 80

MAX_NOME = 150

MAX_SENHA = 200

PBKDF2_ITERACOES = 120_000


# =============================================================================
# CABEÇALHOS HTTP
# =============================================================================

HEADERS_SEGURANCA = {
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "SAMEORIGIN",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "Cache-Control": "no-store",
}


# =============================================================================
# DIRETÓRIOS
# =============================================================================

def criar_diretorios():

    diretorios = [
        DADOS_DIR,
        TEMPLATES_DIR,
        STATIC_DIR,
        CSS_DIR,
        JS_DIR,
        IMG_DIR,
        LOGS_DIR,
        BACKUP_DIR,
    ]

    for diretorio in diretorios:

        diretorio.mkdir(
            parents=True,
            exist_ok=True
        )


criar_diretorios()


# =============================================================================
# LOGGING
# =============================================================================

LOG_FILE = LOGS_DIR / "servidor.log"

logging.basicConfig(
    level=logging.INFO,
    format=(
        "%(asctime)s - "
        "%(levelname)s - "
        "%(message)s"
    ),
    handlers=[
        logging.FileHandler(
            LOG_FILE,
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(
    "CentralTIC"
)


# =============================================================================
# LOCKS
# =============================================================================

usuarios_lock = threading.RLock()

scripts_lock = threading.RLock()

kb_lock = threading.RLock()

sessoes_lock = threading.RLock()

auditoria_lock = threading.RLock()


# =============================================================================
# SESSÕES
# =============================================================================

sessoes: Dict[str, Dict[str, Any]] = {}


# =============================================================================
# CACHE DE APLICAÇÕES
# =============================================================================

apps_list: List[Dict[str, Any]] = []


# =============================================================================
# ESTATÍSTICAS
# =============================================================================

estatisticas = {
    "inicio": datetime.now().isoformat(),
    "requisicoes": 0,
    "logins_sucesso": 0,
    "logins_falha": 0,
    "scripts_criados": 0,
    "scripts_alterados": 0,
    "scripts_excluidos": 0,
    "kb_criados": 0,
    "kb_alterados": 0,
    "kb_excluidos": 0,
}


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def agora_iso() -> str:

    return datetime.now().isoformat(
        timespec="seconds"
    )


def limitar_texto(
    valor: Any,
    limite: int
) -> str:

    return str(
        valor or ""
    ).strip()[:limite]


def gerar_id(
    prefixo: str = ""
) -> str:

    identificador = str(
        uuid.uuid4()
    )

    if prefixo:

        return f"{prefixo}-{identificador}"

    return identificador


def resposta_padrao_erro(
    mensagem: str
) -> Dict[str, Any]:

    return {
        "sucesso": False,
        "erro": mensagem,
        "versao": VERSAO,
        "timestamp": agora_iso()
    }


# =============================================================================
# JSON
# =============================================================================

def ler_json(
    caminho: Path,
    padrao: Any
) -> Any:

    try:

        if not caminho.exists():

            return padrao

        with caminho.open(
            "r",
            encoding="utf-8"
        ) as arquivo:

            return json.load(
                arquivo
            )

    except json.JSONDecodeError as erro:

        logger.error(
            "JSON inválido em %s: %s",
            caminho,
            erro
        )

        return padrao

    except Exception as erro:

        logger.error(
            "Erro ao ler JSON %s: %s",
            caminho,
            erro
        )

        return padrao


def salvar_json(
    caminho: Path,
    dados: Any
):

    caminho.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    temporario = caminho.with_suffix(
        caminho.suffix + ".tmp"
    )

    try:

        with temporario.open(
            "w",
            encoding="utf-8"
        ) as arquivo:

            json.dump(
                dados,
                arquivo,
                ensure_ascii=False,
                indent=2
            )

            arquivo.flush()

            os.fsync(
                arquivo.fileno()
            )

        temporario.replace(
            caminho
        )

    except Exception:

        if temporario.exists():

            try:
                temporario.unlink()
            except Exception:
                pass

        raise


# =============================================================================
# BACKUP
# =============================================================================

def criar_backup(
    caminho: Path
):

    if not caminho.exists():

        return

    try:

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        destino = (
            BACKUP_DIR /
            f"{caminho.stem}_{timestamp}.json.bak"
        )

        destino.write_bytes(
            caminho.read_bytes()
        )

        logger.info(
            "Backup criado: %s",
            destino
        )

    except Exception as erro:

        logger.error(
            "Erro ao criar backup de %s: %s",
            caminho,
            erro
        )


def salvar_json_com_backup(
    caminho: Path,
    dados: Any
):

    if caminho.exists():

        criar_backup(
            caminho
        )

    salvar_json(
        caminho,
        dados
    )


# =============================================================================
# AUDITORIA
# =============================================================================

def registrar_auditoria(
    usuario: str,
    acao: str,
    recurso: str,
    detalhes: Optional[Dict[str, Any]] = None
):

    registro = {
        "id": gerar_id("AUD"),
        "data": agora_iso(),
        "usuario": usuario,
        "acao": acao,
        "recurso": recurso,
        "detalhes": detalhes or {}
    }

    with auditoria_lock:

        dados = ler_json(
            AUDITORIA_FILE,
            []
        )

        if not isinstance(
            dados,
            list
        ):

            dados = []

        dados.append(
            registro
        )

        # Mantém no máximo 5000 registros
        if len(dados) > 5000:

            dados = dados[-5000:]

        salvar_json(
            AUDITORIA_FILE,
            dados
        )

    logger.info(
        "AUDITORIA | %s | %s | %s",
        usuario,
        acao,
        recurso
    )


# =============================================================================
# SENHAS
# =============================================================================

def gerar_salt() -> str:

    return secrets.token_hex(
        16
    )


def gerar_hash_senha(
    senha: str,
    salt: Optional[str] = None
) -> str:

    if salt is None:

        salt = gerar_salt()

    derivada = hashlib.pbkdf2_hmac(
        "sha256",
        senha.encode("utf-8"),
        salt.encode("utf-8"),
        PBKDF2_ITERACOES
    )

    return (
        f"{salt}$"
        f"{derivada.hex()}"
    )


def verificar_senha(
    senha: str,
    senha_hash: str
) -> bool:

    try:

        if not senha_hash:

            return False

        if "$" not in senha_hash:

            return False

        salt, hash_original = (
            senha_hash.split(
                "$",
                1
            )
        )

        derivada = hashlib.pbkdf2_hmac(
            "sha256",
            senha.encode("utf-8"),
            salt.encode("utf-8"),
            PBKDF2_ITERACOES
        )

        return hmac.compare_digest(
            derivada.hex(),
            hash_original
        )

    except Exception as erro:

        logger.error(
            "Erro ao verificar senha: %s",
            erro
        )

        return False


def validar_senha_formato(
    senha: str
) -> Tuple[bool, str]:

    if not senha:

        return (
            False,
            "Senha é obrigatória."
        )

    if len(senha) < 6:

        return (
            False,
            "A senha deve possuir pelo menos 6 caracteres."
        )

    if len(senha) > MAX_SENHA:

        return (
            False,
            "Senha muito grande."
        )

    return True, "OK"


# =============================================================================
# TOKEN
# =============================================================================

def gerar_token() -> str:

    return secrets.token_urlsafe(
        48
    )


# =============================================================================
# USUÁRIOS
# =============================================================================

def usuarios_padrao():

    return {
        "usuarios": [
            {
                "id": "USR-ADMIN",
                "usuario": "admin",
                "senha": gerar_hash_senha(
                    "admin123"
                ),
                "nome": "Administrador",
                "email": "",
                "permissao": "admin",
                "ativo": True,
                "criado_em": agora_iso()
            },
            {
                "id": "USR-AGENTE",
                "usuario": "agente",
                "senha": gerar_hash_senha(
                    "agente123"
                ),
                "nome": "Agente de Atendimento",
                "email": "",
                "permissao": "usuario",
                "ativo": True,
                "criado_em": agora_iso()
            }
        ]
    }


def carregar_usuarios():

    dados = ler_json(
        USUARIOS_FILE,
        {"usuarios": []}
    )

    if not isinstance(
        dados,
        dict
    ):

        dados = {
            "usuarios": []
        }

    if not isinstance(
        dados.get("usuarios"),
        list
    ):

        dados["usuarios"] = []

    return dados


def criar_usuarios_padrao():

    dados = usuarios_padrao()

    with usuarios_lock:

        salvar_json(
            USUARIOS_FILE,
            dados
        )

    logger.info(
        "Arquivo de usuários criado."
    )


def garantir_usuarios_padrao():

    dados = carregar_usuarios()

    usuarios = dados.setdefault(
        "usuarios",
        []
    )

    nomes = {
        str(
            item.get(
                "usuario",
                ""
            )
        ).lower()
        for item in usuarios
    }

    alterou = False

    if "admin" not in nomes:

        usuarios.append(
            {
                "id": "USR-ADMIN",
                "usuario": "admin",
                "senha": gerar_hash_senha(
                    "admin123"
                ),
                "nome": "Administrador",
                "email": "",
                "permissao": "admin",
                "ativo": True,
                "criado_em": agora_iso()
            }
        )

        alterou = True

    if "agente" not in nomes:

        usuarios.append(
            {
                "id": "USR-AGENTE",
                "usuario": "agente",
                "senha": gerar_hash_senha(
                    "agente123"
                ),
                "nome": "Agente de Atendimento",
                "email": "",
                "permissao": "usuario",
                "ativo": True,
                "criado_em": agora_iso()
            }
        )

        alterou = True

    if alterou:

        with usuarios_lock:

            salvar_json_com_backup(
                USUARIOS_FILE,
                dados
            )


def encontrar_usuario(
    usuario: str
) -> Optional[Dict[str, Any]]:

    dados = carregar_usuarios()

    for item in dados.get(
        "usuarios",
        []
    ):

        if (
            str(
                item.get(
                    "usuario",
                    ""
                )
            ).lower()
            ==
            usuario.lower()
        ):

            return item

    return None


def encontrar_usuario_id(
    usuario_id: str
) -> Optional[Dict[str, Any]]:

    dados = carregar_usuarios()

    for item in dados.get(
        "usuarios",
        []
    ):

        if item.get(
            "id"
        ) == usuario_id:

            return item

    return None


def validar_credenciais(
    usuario: str,
    senha: str
) -> Tuple[bool, str, str]:

    dados = carregar_usuarios()

    usuarios = dados.get(
        "usuarios",
        []
    )

    for usuario_data in usuarios:

        if (
            str(
                usuario_data.get(
                    "usuario",
                    ""
                )
            ).lower()
            !=
            usuario.lower()
        ):

            continue

        if not usuario_data.get(
            "ativo",
            True
        ):

            estatisticas[
                "logins_falha"
            ] += 1

            logger.warning(
                "Login de usuário inativo: %s",
                usuario
            )

            return False, "", ""

        senha_armazenada = str(
            usuario_data.get(
                "senha",
                ""
            )
        )

        senha_valida = False

        if "$" in senha_armazenada:

            senha_valida = verificar_senha(
                senha,
                senha_armazenada
            )

        else:

            # Compatibilidade com versões antigas
            try:

                senha_valida = (
                    hmac.compare_digest(
                        senha,
                        senha_armazenada
                    )
                )

                if senha_valida:

                    usuario_data[
                        "senha"
                    ] = gerar_hash_senha(
                        senha
                    )

                    with usuarios_lock:

                        salvar_json_com_backup(
                            USUARIOS_FILE,
                            dados
                        )

            except Exception:

                senha_valida = False

        if not senha_valida:

            estatisticas[
                "logins_falha"
            ] += 1

            logger.warning(
                "Senha inválida para usuário: %s",
                usuario
            )

            return False, "", ""

        estatisticas[
            "logins_sucesso"
        ] += 1

        nome = usuario_data.get(
            "nome",
            usuario
        )

        permissao = usuario_data.get(
            "permissao",
            "usuario"
        )

        registrar_auditoria(
            usuario,
            "LOGIN",
            "autenticacao"
        )

        return (
            True,
            nome,
            permissao
        )

    estatisticas[
        "logins_falha"
    ] += 1

    return False, "", ""


# =============================================================================
# SESSÕES
# =============================================================================

def criar_sessao(
    usuario: str,
    nome: str,
    permissao: str
) -> str:

    token = gerar_token()

    agora = datetime.now()

    with sessoes_lock:

        sessoes[token] = {
            "usuario": usuario,
            "nome": nome,
            "permissao": permissao,
            "criada_em": agora.isoformat(),
            "ultimo_acesso": agora.isoformat(),
            "expiracao": (
                agora.timestamp()
                + SESSION_TIMEOUT
            )
        }

    logger.info(
        "Sessão criada | %s",
        usuario
    )

    return token


def validar_sessao(
    token: Optional[str]
) -> Tuple[bool, Optional[Dict[str, Any]]]:

    if not token:

        return False, None

    with sessoes_lock:

        sessao = sessoes.get(
            token
        )

        if not sessao:

            return False, None

        agora = datetime.now()

        expiracao = float(
            sessao.get(
                "expiracao",
                0
            )
        )

        if agora.timestamp() > expiracao:

            sessoes.pop(
                token,
                None
            )

            return False, None

        sessao[
            "ultimo_acesso"
        ] = agora.isoformat()

        sessao[
            "expiracao"
        ] = (
            agora.timestamp()
            + SESSION_TIMEOUT
        )

        return True, dict(
            sessao
        )


def destruir_sessao(
    token: Optional[str]
):

    if not token:

        return

    with sessoes_lock:

        sessao = sessoes.pop(
            token,
            None
        )

    if sessao:

        registrar_auditoria(
            sessao.get(
                "usuario",
                ""
            ),
            "LOGOUT",
            "autenticacao"
        )


def usuario_e_admin(
    token: Optional[str]
) -> bool:

    valido, usuario = validar_sessao(
        token
    )

    if not valido or not usuario:

        return False

    return (
        usuario.get(
            "permissao"
        ) == "admin"
    )


# =============================================================================
# COOKIE
# =============================================================================

def extrair_cookie(
    headers,
    nome: str
) -> Optional[str]:

    cookies = headers.get(
        "Cookie",
        ""
    )

    for item in cookies.split(";"):

        item = item.strip()

        if "=" not in item:

            continue

        chave, valor = item.split(
            "=",
            1
        )

        if chave.strip() == nome:

            return valor.strip()

    return None


def usuario_da_requisicao(
    handler
) -> Tuple[bool, Optional[Dict[str, Any]]]:

    token = extrair_cookie(
        handler.headers,
        "session_token"
    )

    return validar_sessao(
        token
    )


def exigir_admin(
    handler
) -> Tuple[bool, Optional[Dict[str, Any]]]:

    valido, usuario = (
        usuario_da_requisicao(
            handler
        )
    )

    if not valido:

        return False, None

    if not usuario:

        return False, None

    if usuario.get(
        "permissao"
    ) != "admin":

        return False, usuario

    return True, usuario


# =============================================================================
# XLSX
# =============================================================================

NS = {
    "main":
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",

    "rel":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

    "pkgrel":
        "http://schemas.openxmlformats.org/package/2006/relationships"
}


def ler_excel(
    arquivo: Path
) -> List[Dict[str, Any]]:

    if not arquivo.exists():

        logger.warning(
            "Excel não encontrado: %s",
            arquivo
        )

        return []

    try:

        with zipfile.ZipFile(
            arquivo,
            "r"
        ) as zip_file:

            nomes = zip_file.namelist()

            shared_strings = []

            if (
                "xl/sharedStrings.xml"
                in nomes
            ):

                root = ET.fromstring(
                    zip_file.read(
                        "xl/sharedStrings.xml"
                    )
                )

                for item in root.findall(
                    ".//main:si",
                    NS
                ):

                    textos = []

                    for texto in item.findall(
                        ".//main:t",
                        NS
                    ):

                        if texto.text:

                            textos.append(
                                texto.text
                            )

                    shared_strings.append(
                        "".join(
                            textos
                        )
                    )

            workbook_root = ET.fromstring(
                zip_file.read(
                    "xl/workbook.xml"
                )
            )

            sheets = workbook_root.find(
                "main:sheets",
                NS
            )

            if sheets is None:

                return []

            primeira_sheet = (
                sheets.find(
                    "main:sheet",
                    NS
                )
            )

            if primeira_sheet is None:

                return []

            rel_id = primeira_sheet.get(
                "{%s}id"
                % NS["rel"]
            )

            if not rel_id:

                return []

            rels_root = ET.fromstring(
                zip_file.read(
                    "xl/_rels/workbook.xml.rels"
                )
            )

            target = None

            for rel in rels_root:

                if rel.get(
                    "Id"
                ) == rel_id:

                    target = rel.get(
                        "Target"
                    )

                    break

            if not target:

                return []

            target = target.replace(
                "\\",
                "/"
            )

            if target.startswith("/"):

                target = target.lstrip("/")

            if not target.startswith(
                "xl/"
            ):

                target = (
                    "xl/"
                    + target
                )

            sheet_root = ET.fromstring(
                zip_file.read(
                    target
                )
            )

            rows = sheet_root.findall(
                ".//main:sheetData/main:row",
                NS
            )

            dados = []

            cabecalho = []

            for indice, row in enumerate(
                rows
            ):

                valores = []

                for cell in row.findall(
                    "main:c",
                    NS
                ):

                    valor = cell.find(
                        "main:v",
                        NS
                    )

                    tipo = cell.get(
                        "t"
                    )

                    if tipo == "inlineStr":

                        textos = []

                        for texto in cell.findall(
                            ".//main:t",
                            NS
                        ):

                            if texto.text:

                                textos.append(
                                    texto.text
                                )

                        valores.append(
                            "".join(
                                textos
                            ).strip()
                        )

                        continue

                    if valor is None:

                        valores.append(
                            ""
                        )

                        continue

                    texto = (
                        valor.text
                        or ""
                    )

                    if (
                        tipo == "s"
                        and texto.isdigit()
                    ):

                        numero = int(
                            texto
                        )

                        if (
                            numero
                            <
                            len(
                                shared_strings
                            )
                        ):

                            texto = (
                                shared_strings[
                                    numero
                                ]
                            )

                    valores.append(
                        str(
                            texto
                        ).strip()
                    )

                if indice == 0:

                    cabecalho = [
                        x.lower().strip()
                        for x in valores
                    ]

                    continue

                if not any(
                    valores
                ):

                    continue

                registro = {}

                for i, coluna in enumerate(
                    cabecalho
                ):

                    if i < len(
                        valores
                    ):

                        registro[
                            coluna
                        ] = valores[i]

                dados.append(
                    registro
                )

            logger.info(
                "Excel carregado: %s registros",
                len(dados)
            )

            return dados

    except Exception as erro:

        logger.exception(
            "Erro ao ler XLSX: %s",
            erro
        )

        return []


def _norm_chave(texto: str) -> str:
    """Normaliza cabeçalho da planilha para comparação."""
    import unicodedata
    t = str(texto or "").strip().lower()
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.replace("_", " ")
    t = " ".join(t.split())
    return t


def _valor_linha(linha: dict, *candidatos: str) -> str:
    if not isinstance(linha, dict):
        return ""
    mapa = { _norm_chave(k): v for k, v in linha.items() }
    for cand in candidatos:
        v = mapa.get(_norm_chave(cand))
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""


def carregar_aplicacoes():

    linhas = ler_excel(
        EXCEL_FILE
    )

    resultado = []

    for linha in linhas:

        codigo = _valor_linha(
            linha,
            "código do software",
            "codigo do software",
            "codigo",
            "código",
            "cod software",
            "cod.",
        )

        sigla = _valor_linha(
            linha,
            "sigla",
            "sigla app",
            "sigla aplicação",
            "sigla aplicacao",
        )

        nome = _valor_linha(
            linha,
            "nome da aplicação",
            "nome da aplicacao",
            "nome aplicação",
            "nome aplicacao",
            "nome",
            "aplicação",
            "aplicacao",
        )

        mesa = _valor_linha(
            linha,
            "mesa",
            "mesa atendimento",
            "fila",
            "grupo",
        )


        if codigo or nome:

            resultado.append(
                {
                    "codigo": str(
                        codigo
                    ).strip(),

                    "sigla": str(
                        sigla
                    ).strip(),

                    "nome": str(
                        nome
                    ).strip(),

                    "mesa": str(
                        mesa
                    ).strip(),

                    "status": "Ativo"
                }
            )

    return resultado


# =============================================================================
# SCRIPTS
# =============================================================================

def scripts_padrao():

    return {
        "scripts": [

            {
                "id": gerar_id("SCR"),
                "emoji": "📞",
                "titulo": "Saudação Inicial",
                "categoria": "Abertura",
                "texto": (
                    "Olá, bom dia/tarde! "
                    "Meu nome é [SEU NOME], "
                    "sou do Atendimento TIC N1.\n"
                    "Em que posso ajudar?"
                ),
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": gerar_id("SCR"),
                "emoji": "🔍",
                "titulo": "Identificação do Problema",
                "categoria": "Diagnóstico",
                "texto": (
                    "Entendido. Qual aplicação "
                    "está apresentando problema?\n"
                    "Qual mensagem de erro aparece?\n"
                    "Desde quando está acontecendo?"
                ),
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": gerar_id("SCR"),
                "emoji": "🛠️",
                "titulo": "Orientação de Cache",
                "categoria": "Orientação",
                "texto": (
                    "Por favor, tente limpar o cache "
                    "utilizando Ctrl + Shift + R "
                    "e atualize a página.\n"
                    "O problema persiste?"
                ),
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": gerar_id("SCR"),
                "emoji": "🔄",
                "titulo": "Escalonamento N2",
                "categoria": "Escalonamento",
                "texto": (
                    "Vou registrar o chamado e "
                    "encaminhar para a equipe N2 "
                    "responsável pela aplicação."
                ),
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": gerar_id("SCR"),
                "emoji": "✅",
                "titulo": "Confirmação de Resolução",
                "categoria": "Encerramento",
                "texto": (
                    "O problema foi resolvido do seu lado?\n"
                    "Posso realizar o encerramento do chamado?"
                ),
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": gerar_id("SCR"),
                "emoji": "🏁",
                "titulo": "Encerramento",
                "categoria": "Encerramento",
                "texto": (
                    "Chamado registrado com sucesso.\n"
                    "Qualquer outra dúvida, "
                    "estamos à disposição.\n"
                    "Tenha um excelente dia!"
                ),
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            }

        ]
    }


def carregar_scripts():

    dados = ler_json(
        SCRIPTS_FILE,
        {"scripts": []}
    )

    if not isinstance(
        dados,
        dict
    ):

        dados = {
            "scripts": []
        }

    if not isinstance(
        dados.get("scripts"),
        list
    ):

        dados["scripts"] = []

    return dados


def criar_scripts_padrao():

    with scripts_lock:

        salvar_json(
            SCRIPTS_FILE,
            scripts_padrao()
        )


# =============================================================================
# KB
# =============================================================================

def kb_padrao():

    return {
        "kb": [

            {
                "id": "KB0265800",
                "titulo": "Solicitar acesso ao Novo SINPEP",
                "sistema": "SINPEP",
                "categoria": "Acesso",
                "descricao": (
                    "Procedimento para solicitação "
                    "de acesso ao Novo SINPEP."
                ),
                "url": "",
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": "KB0290150",
                "titulo": "Acessar SIGEM",
                "sistema": "SIGEM",
                "categoria": "Acesso",
                "descricao": (
                    "Orientação para acesso ao SIGEM."
                ),
                "url": "",
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            },

            {
                "id": "KB0290182",
                "titulo": (
                    "Acessar SIGEM - Sistema Integrado "
                    "de Gerenciamento de Empreendimentos"
                ),
                "sistema": "SIGEM",
                "categoria": "Acesso",
                "descricao": (
                    "Procedimento para acesso ao SIGEM."
                ),
                "url": "",
                "ativo": True,
                "criado_por": "sistema",
                "criado_em": agora_iso()
            }

        ]
    }


def carregar_kb():

    dados = ler_json(
        KB_FILE,
        {"kb": []}
    )

    if not isinstance(
        dados,
        dict
    ):

        dados = {
            "kb": []
        }

    if not isinstance(
        dados.get("kb"),
        list
    ):

        dados["kb"] = []

    return dados


def criar_kb_padrao():

    with kb_lock:

        salvar_json(
            KB_FILE,
            kb_padrao()
        )


# =============================================================================
# VALIDAÇÃO DE SCRIPT
# =============================================================================

def validar_script(
    titulo: str,
    texto: str,
    emoji: str,
    categoria: str
) -> Tuple[bool, str]:

    titulo = limitar_texto(
        titulo,
        MAX_TITULO
    )

    texto = limitar_texto(
        texto,
        MAX_TEXTO
    )

    emoji = limitar_texto(
        emoji,
        MAX_EMOJI
    )

    categoria = limitar_texto(
        categoria,
        MAX_CATEGORIA
    )

    if not titulo:

        return (
            False,
            "Título é obrigatório."
        )

    if not texto:

        return (
            False,
            "Texto é obrigatório."
        )

    if not categoria:

        return (
            False,
            "Categoria é obrigatória."
        )

    return True, "OK"


# =============================================================================
# VALIDAÇÃO KB
# =============================================================================

def validar_kb(
    titulo: str,
    sistema: str,
    categoria: str,
    descricao: str
) -> Tuple[bool, str]:

    if not limitar_texto(
        titulo,
        MAX_TITULO
    ):

        return (
            False,
            "Título é obrigatório."
        )

    if not limitar_texto(
        sistema,
        100
    ):

        return (
            False,
            "Sistema é obrigatório."
        )

    if not limitar_texto(
        categoria,
        MAX_CATEGORIA
    ):

        return (
            False,
            "Categoria é obrigatória."
        )

    if not limitar_texto(
        descricao,
        MAX_TEXTO
    ):

        return (
            False,
            "Descrição é obrigatória."
        )

    return True, "OK"

# =============================================================================
# BUSCA GLOBAL
# =============================================================================

def normalizar_busca(valor):
    """
    Normaliza o texto para melhorar a pesquisa.
    Remove espaços extras e ignora maiúsculas/minúsculas.
    """
    return " ".join(
        str(valor or "").strip().lower().split()
    )


def contem_termo(texto, termo):
    """
    Verifica se o termo aparece no texto.
    """
    return normalizar_busca(termo) in normalizar_busca(texto)


def buscar_global(termo):
    """
    Pesquisa simultaneamente em:

    - Aplicações vindas do Excel
    - Scripts
    - Knowledge Base

    Retorna resultados separados por tipo.
    """

    termo = normalizar_busca(termo)

    resultados = {
        "aplicacoes": [],
        "scripts": [],
        "kb": [],
        "total": 0
    }

    # -------------------------------------------------------------------------
    # APLICAÇÕES
    # -------------------------------------------------------------------------

    for app in apps_list:

        texto_app = " ".join([
            str(app.get("codigo", "")),
            str(app.get("sigla", "")),
            str(app.get("nome", "")),
            str(app.get("mesa", "")),
            str(app.get("status", ""))
        ])

        if not termo or contem_termo(texto_app, termo):

            resultados["aplicacoes"].append({
                "tipo": "aplicacao",
                "codigo": app.get("codigo", ""),
                "sigla": app.get("sigla", ""),
                "nome": app.get("nome", ""),
                "mesa": app.get("mesa", ""),
                "status": app.get("status", "Ativo")
            })

    # -------------------------------------------------------------------------
    # SCRIPTS
    # -------------------------------------------------------------------------

    dados_scripts = carregar_scripts()

    for script in dados_scripts.get("scripts", []):

        texto_script = " ".join([
            str(script.get("id", "")),
            str(script.get("titulo", "")),
            str(script.get("categoria", "")),
            str(script.get("texto", "")),
            str(script.get("tag", ""))
        ])

        if not termo or contem_termo(
            texto_script,
            termo
        ):

            resultados["scripts"].append({
                "tipo": "script",
                "id": script.get("id", ""),
                "emoji": script.get("emoji", "📝"),
                "titulo": script.get("titulo", ""),
                "categoria": (
                    script.get("categoria")
                    or script.get("tag")
                    or "Geral"
                ),
                "texto": script.get("texto", "")
            })

    # -------------------------------------------------------------------------
    # KNOWLEDGE BASE
    # -------------------------------------------------------------------------

    dados_kb = carregar_kb()

    for item in dados_kb.get("kb", []):

        texto_kb = " ".join([
            str(item.get("id", "")),
            str(item.get("titulo", "")),
            str(item.get("sistema", "")),
            str(item.get("categoria", "")),
            str(item.get("descricao", "")),
            str(item.get("texto", "")),
            str(item.get("conteudo", "")),
        ])

        if not termo or contem_termo(
            texto_kb,
            termo
        ):

            resultados["kb"].append({
                "tipo": "kb",
                "id": item.get("id", ""),
                "titulo": item.get("titulo", ""),
                "sistema": item.get("sistema", ""),
                "categoria": item.get("categoria", ""),
                "descricao": (
                    item.get("descricao")
                    or item.get("texto")
                    or item.get("conteudo")
                    or ""
                ),
                "url": item.get("url", "")
            })

    # -------------------------------------------------------------------------
    # TOTAL
    # -------------------------------------------------------------------------

    resultados["total"] = (
        len(resultados["aplicacoes"])
        + len(resultados["scripts"])
        + len(resultados["kb"])
    )

    return resultados

# =============================================================================
# HTTP HANDLER
# =============================================================================



def gravar_arquivo_bytes(caminho: Path, conteudo: bytes):
    """
    Grava bytes de forma resiliente no Windows
    (evita PermissionError em os.replace quando o destino está bloqueado).
    """
    import time
    import shutil

    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)

    temporario = caminho.with_name(caminho.name + ".tmp")
    # se sobrou tmp antigo, remove
    try:
        if temporario.exists():
            temporario.unlink()
    except Exception:
        pass

    temporario.write_bytes(conteudo)

    ultimo_erro = None
    for tentativa in range(8):
        try:
            if caminho.exists():
                try:
                    # remove atributo somente-leitura se houver
                    os.chmod(str(caminho), 0o666)
                except Exception:
                    pass
            os.replace(str(temporario), str(caminho))
            return
        except PermissionError as erro:
            ultimo_erro = erro
            time.sleep(0.15 * (tentativa + 1))
        except OSError as erro:
            ultimo_erro = erro
            time.sleep(0.15 * (tentativa + 1))

    # Fallback: sobrescreve por cópia
    try:
        if caminho.exists():
            try:
                os.chmod(str(caminho), 0o666)
            except Exception:
                pass
        with open(str(caminho), "wb") as destino:
            destino.write(conteudo)
            destino.flush()
            os.fsync(destino.fileno())
        try:
            if temporario.exists():
                temporario.unlink()
        except Exception:
            pass
        return
    except Exception as erro_fallback:
        try:
            if temporario.exists():
                temporario.unlink()
        except Exception:
            pass
        raise PermissionError(
            f"Não foi possível gravar {caminho.name}. "
            f"Feche o arquivo no Excel se estiver aberto e tente de novo. "
            f"Detalhe: {erro_fallback}"
        ) from (ultimo_erro or erro_fallback)


def extrair_arquivo_multipart(content_type: str, body: bytes):
    """
    Extrai o primeiro arquivo de um body multipart/form-data.
    Retorna (nome_arquivo, conteudo_bytes) ou (None, None).
    """
    if not content_type or "multipart/form-data" not in content_type:
        return None, None

    boundary = None
    for parte in content_type.split(";"):
        parte = parte.strip()
        if parte.lower().startswith("boundary="):
            boundary = parte.split("=", 1)[1].strip().strip('"')
            break

    if not boundary:
        return None, None

    delim = b"--" + boundary.encode("utf-8")
    blocos = body.split(delim)

    for bloco in blocos:
        if b"Content-Disposition" not in bloco:
            continue
        if b"filename=" not in bloco and b"filename*=" not in bloco:
            continue

        if b"\r\n\r\n" in bloco:
            cabecalho, conteudo = bloco.split(b"\r\n\r\n", 1)
        elif b"\n\n" in bloco:
            cabecalho, conteudo = bloco.split(b"\n\n", 1)
        else:
            continue

        if conteudo.endswith(b"\r\n"):
            conteudo = conteudo[:-2]
        elif conteudo.endswith(b"\n"):
            conteudo = conteudo[:-1]

        # remove eventual fechamento residual
        if conteudo.endswith(b"--"):
            conteudo = conteudo[:-2].rstrip(b"\r\n")

        nome = None
        for linha in cabecalho.decode("utf-8", errors="ignore").splitlines():
            low = linha.lower()
            if "filename*=" in low:
                try:
                    parte = linha.split("*=", 1)[1].strip()
                    if "''" in parte:
                        nome = parte.split("''", 1)[1].strip().strip('"')
                    else:
                        nome = parte.strip().strip('"')
                except Exception:
                    pass
            elif "filename=" in low:
                try:
                    parte = linha.split("filename=", 1)[1].strip()
                    nome = parte.split(";")[0].strip().strip('"')
                except Exception:
                    pass

        if nome and conteudo is not None and len(conteudo) > 0:
            nome = Path(nome).name
            return nome, conteudo

    return None, None


def validar_xlsx(conteudo: bytes) -> bool:
    """Verifica se o conteúdo parece um arquivo .xlsx válido."""
    if not conteudo or len(conteudo) < 4:
        return False
    if conteudo[:2] != b"PK":
        return False
    try:
        import io
        with zipfile.ZipFile(io.BytesIO(conteudo), "r") as zf:
            nomes = zf.namelist()
            return any(n.startswith("xl/") for n in nomes)
    except Exception:
        return False



def modelo_dialogo_padrao():
    return {
        "fonte_url": MODELO_DIALOGO_URL_PADRAO,
        "kb": "KB0029502",
        "titulo": "Modelo de Diálogo",
        "atualizado_em": None,
        "atualizado_por": None,
        "intervalo_dias": 5,
        "scripts": [],
    }


def carregar_modelo_dialogo() -> dict:
    dados = ler_json(MODELO_DIALOGO_FILE, None)
    if not isinstance(dados, dict):
        return modelo_dialogo_padrao()
    base = modelo_dialogo_padrao()
    base.update(dados)
    if not isinstance(base.get("scripts"), list):
        base["scripts"] = []
    try:
        base["intervalo_dias"] = max(1, int(base.get("intervalo_dias") or 5))
    except Exception:
        base["intervalo_dias"] = 5
    return base


def salvar_modelo_dialogo(dados: dict, usuario: str = "sistema"):
    atual = carregar_modelo_dialogo()
    scripts = dados.get("scripts", atual.get("scripts") or [])
    if not isinstance(scripts, list):
        raise ValueError("Lista de scripts inválida.")

    # validação leve
    limpos = []
    for i, item in enumerate(scripts):
        if not isinstance(item, dict):
            continue
        texto = str(item.get("texto") or "").strip()
        titulo = str(item.get("titulo") or "").strip()
        if not texto and not titulo:
            continue
        limpos.append({
            "id": str(item.get("id") or f"mdl_{i+1}"),
            "emoji": str(item.get("emoji") or "💬")[:8],
            "titulo": titulo[:200] or f"Script {i+1}",
            "tag": str(item.get("tag") or item.get("categoria") or "Geral")[:120],
            "ordem": str(item.get("ordem") or ""),
            "texto": texto[:20000],
            "categoria": str(item.get("categoria") or item.get("tag") or "Geral")[:120],
            "sistema": bool(item.get("sistema", True)),
            "ativo": bool(item.get("ativo", True)),
        })

    if not limpos:
        raise ValueError("Nenhum script válido encontrado no modelo.")

    agora = datetime.now().isoformat(timespec="seconds")
    novo = {
        "fonte_url": str(
            dados.get("fonte_url")
            or atual.get("fonte_url")
            or MODELO_DIALOGO_URL_PADRAO
        ),
        "kb": str(dados.get("kb") or atual.get("kb") or "KB0029502"),
        "titulo": str(dados.get("titulo") or atual.get("titulo") or "Modelo de Diálogo"),
        "atualizado_em": agora,
        "atualizado_por": usuario or "sistema",
        "intervalo_dias": int(dados.get("intervalo_dias") or atual.get("intervalo_dias") or 5),
        "scripts": limpos,
    }
    try:
        if MODELO_DIALOGO_FILE.exists():
            BACKUP_DIR.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            (BACKUP_DIR / f"modelo_dialogo_{ts}.json.bak").write_bytes(
                MODELO_DIALOGO_FILE.read_bytes()
            )
    except Exception as e:
        logger.warning("Backup modelo diálogo: %s", e)

    salvar_json(MODELO_DIALOGO_FILE, novo)
    return novo


def status_modelo_dialogo(dados: dict = None) -> dict:
    dados = dados or carregar_modelo_dialogo()
    intervalo = int(dados.get("intervalo_dias") or 5)
    atualizado_em = dados.get("atualizado_em")
    dias = None
    vencido = True
    if atualizado_em:
        try:
            dt = datetime.fromisoformat(str(atualizado_em).replace("Z", ""))
            dias = (datetime.now() - dt).days
            vencido = dias >= intervalo
        except Exception:
            vencido = True
            dias = None
    return {
        "fonte_url": dados.get("fonte_url") or MODELO_DIALOGO_URL_PADRAO,
        "kb": dados.get("kb") or "KB0029502",
        "atualizado_em": atualizado_em,
        "atualizado_por": dados.get("atualizado_por"),
        "intervalo_dias": intervalo,
        "dias_desde_atualizacao": dias,
        "precisa_atualizar": vencido,
        "total_scripts": len(dados.get("scripts") or []),
    }


class Handler(
    http.server.BaseHTTPRequestHandler
):

    server_version = (
        "CentralTIC/"
        + VERSAO
    )

    sys_version = ""

    # =========================================================================
    # HEADERS
    # =========================================================================

    def enviar_headers_seguranca(self):

        for chave, valor in (
            HEADERS_SEGURANCA.items()
        ):

            self.send_header(
                chave,
                valor
            )

    # =========================================================================
    # JSON
    # =========================================================================

    def enviar_json(
        self,
        dados,
        status=200,
        headers=None
    ):

        corpo = json.dumps(
            dados,
            ensure_ascii=False
        ).encode(
            "utf-8"
        )

        self.send_response(
            status
        )

        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8"
        )

        self.send_header(
            "Content-Length",
            str(len(corpo))
        )

        self.enviar_headers_seguranca()

        if headers:

            for chave, valor in (
                headers.items()
            ):

                self.send_header(
                    chave,
                    valor
                )

        self.end_headers()

        try:
            self.wfile.write(
                corpo
            )
        except (
            ConnectionAbortedError,
            ConnectionResetError,
            BrokenPipeError,
            OSError,
        ):
            logger.debug(
                "Conexão encerrada pelo cliente ao enviar JSON."
            )
            return

    # =========================================================================
    # HTML
    # =========================================================================

    def enviar_html(
        self,
        conteudo,
        status=200
    ):

        corpo = conteudo.encode(
            "utf-8"
        )

        try:

            self.send_response(
                status
            )

            self.send_header(
                "Content-Type",
                "text/html; charset=utf-8"
            )

            self.send_header(
                "Content-Length",
                str(len(corpo))
            )

            self.enviar_headers_seguranca()

            self.end_headers()

            self.wfile.write(
                corpo
            )

        except (
            ConnectionAbortedError,
            ConnectionResetError,
            BrokenPipeError,
            OSError,
        ):
            # Cliente fechou a aba/navegou antes do fim do envio
            logger.debug(
                "Conexão encerrada pelo cliente ao enviar HTML."
            )
            return

    # =========================================================================
    # REDIRECIONAMENTO
    # =========================================================================

    def redirecionar(
        self,
        destino
    ):

        self.send_response(
            302
        )

        self.send_header(
            "Location",
            destino
        )

        self.send_header(
            "Cache-Control",
            "no-store"
        )

        self.end_headers()

    def redirecionar_com_cookie(
        self,
        destino,
        cookie
    ):

        self.send_response(
            302
        )

        self.send_header(
            "Location",
            destino
        )

        self.send_header(
            "Set-Cookie",
            cookie
        )

        self.send_header(
            "Cache-Control",
            "no-store"
        )

        self.end_headers()

    # =========================================================================
    # BODY
    # =========================================================================

    def ler_body_json(self):

        try:

            tamanho = int(
                self.headers.get(
                    "Content-Length",
                    "0"
                )
            )

        except ValueError:

            raise ValueError(
                "Content-Length inválido."
            )

        if tamanho <= 0:

            raise ValueError(
                "Corpo da requisição vazio."
            )

        if tamanho > MAX_BODY_SIZE:

            raise ValueError(
                "Requisição muito grande."
            )

        corpo = self.rfile.read(
            tamanho
        )

        if not corpo:

            raise ValueError(
                "Corpo vazio."
            )

        return json.loads(
            corpo.decode(
                "utf-8"
            )
        )

    # =========================================================================
    # AUTENTICAÇÃO
    # =========================================================================

    def exigir_login_api(self):

        valido, usuario = (
            usuario_da_requisicao(
                self
            )
        )

        if not valido or not usuario:

            self.enviar_json(
                resposta_padrao_erro(
                    "Não autenticado."
                ),
                401
            )

            return None

        return usuario

    # =========================================================================
    # GET
    # =========================================================================

    def do_GET(self):

        estatisticas[
            "requisicoes"
        ] += 1

        try:

            url = urlparse(
                self.path
            )

            caminho = unquote(
                url.path
            )

            # ============================================================
            # PÁGINA DE APLICAÇÕES
            # ============================================================
            if caminho == "/aplicacoes":

                return self.servir_arquivo(
                    "aplicacoes.html"
                )

            # -----------------------------------------------------------------
            # CHAMADO (singular) — mesma tela de Aplicações na aba Chamados
            # -----------------------------------------------------------------
            if caminho in ("/chamado", "/chamados"):

                return self.servir_arquivo(
                    "aplicacoes.html"
                )

            # -----------------------------------------------------------------
            # LOGIN
            # -----------------------------------------------------------------

            if caminho == "/login":

                self.pagina_login()

                return

            # -----------------------------------------------------------------
            # LOGOUT
            # -----------------------------------------------------------------

            if caminho == "/logout":

                token = extrair_cookie(
                    self.headers,
                    "session_token"
                )

                destruir_sessao(
                    token
                )

                cookie = (
                    "session_token=; "
                    "Path=/; "
                    "HttpOnly; "
                    "SameSite=Lax; "
                    "Max-Age=0"
                    f"{flag_cookie_secure()}"
                )

                self.redirecionar_com_cookie(
                    "/login",
                    cookie
                )

                return

            # -----------------------------------------------------------------
            # HOME
            # -----------------------------------------------------------------

            if caminho in (
                "/",
                "/index.html"
            ):

                valido, usuario = (
                    usuario_da_requisicao(
                        self
                    )
                )

                if not valido:

                    self.redirecionar(
                        "/login"
                    )

                    return

                self.pagina_home(
                    usuario
                )

                return

            # -----------------------------------------------------------------
            # APLICAÇÕES
            # ----------

            if caminho == "/aplicacoes":
                return self.servir_arquivo("aplicacoes.html")

            # -----------------------------------------------------------------
            # ADMIN
            # -----------------------------------------------------------------

            if caminho == "/admin":

                valido, usuario = (
                    exigir_admin(
                        self
                    )
                )

                if not valido:

                    if usuario:

                        self.enviar_json(
                            resposta_padrao_erro(
                                "Acesso administrativo não autorizado."
                            ),
                            403
                        )

                    else:

                        self.redirecionar(
                            "/login"
                        )

                    return

                self.pagina_admin(
                    usuario
                )

                return

            # -----------------------------------------------------------------
            # DIÁLOGO
            # -----------------------------------------------------------------

            if caminho == "/dialogo":

                valido, usuario = (
                    usuario_da_requisicao(
                        self
                    )
                )

                if not valido:

                    self.redirecionar(
                        "/login"
                    )

                    return

                self.pagina_dialogo(
                    usuario
                )

                return

            # -----------------------------------------------------------------
            # API ME
            # -----------------------------------------------------------------

            
            # -----------------------------------------------------------------
            # API USUÁRIOS (lista - admin)
            # -----------------------------------------------------------------

            if caminho == "/api/usuarios":

                valido, usuario = exigir_admin(self)

                if not valido:
                    if usuario:
                        self.enviar_json(
                            resposta_padrao_erro(
                                "Acesso administrativo não autorizado."
                            ),
                            403
                        )
                    else:
                        self.enviar_json(
                            resposta_padrao_erro(
                                "Não autenticado."
                            ),
                            401
                        )
                    return

                dados = carregar_usuarios()
                lista = []

                for item in dados.get("usuarios", []):
                    lista.append({
                        "id": item.get("id", ""),
                        "usuario": item.get("usuario", ""),
                        "nome": item.get("nome", ""),
                        "email": item.get("email", ""),
                        "permissao": item.get("permissao", "usuario"),
                        "ativo": item.get("ativo", True),
                        "criado_em": item.get("criado_em", ""),
                        "criado_por": item.get("criado_por", ""),
                        "alterado_em": item.get("alterado_em", ""),
                        "alterado_por": item.get("alterado_por", ""),
                    })

                self.enviar_json({
                    "sucesso": True,
                    "total": len(lista),
                    "dados": lista
                })
                return

            if caminho == "/api/me":

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                self.enviar_json(
                    {
                        "sucesso": True,
                        "usuario": {
                            "usuario":
                                usuario["usuario"],
                            "nome":
                                usuario["nome"],
                            "permissao":
                                usuario["permissao"]
                        }
                    }
                )

                return

            # -----------------------------------------------------------------
            # API HEALTH
            # -----------------------------------------------------------------

            if caminho == "/api/health":

                self.enviar_json(
                    {
                        "sucesso": True,
                        "sistema":
                            NOME_SISTEMA,
                        "versao":
                            VERSAO,
                        "status":
                            "online",
                        "timestamp":
                            agora_iso()
                    }
                )

                return

            # -----------------------------------------------------------------
            # API APPS
            # -----------------------------------------------------------------

            
            if caminho in ("/api/modelo-dialogo", "/api/modelo_dialogo"):

                usuario = self.exigir_login_api()
                if not usuario:
                    return

                dados = carregar_modelo_dialogo()
                st = status_modelo_dialogo(dados)
                self.enviar_json({
                    "sucesso": True,
                    "dados": {
                        **st,
                        "scripts": dados.get("scripts") or [],
                        "titulo": dados.get("titulo") or "Modelo de Diálogo",
                    }
                })
                return

            if caminho in ("/api/modelo-dialogo/status", "/api/modelo_dialogo/status"):

                usuario = self.exigir_login_api()
                if not usuario:
                    return

                self.enviar_json({
                    "sucesso": True,
                    "dados": status_modelo_dialogo()
                })
                return


            if caminho in ("/api/apps", "/api/aplicacoes"):

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                global apps_list
                try:
                    apps_list = carregar_aplicacoes()
                except Exception as erro:
                    logger.exception("Erro ao carregar aplicações")
                    self.enviar_json(
                        resposta_padrao_erro(
                            f"Falha ao ler planilha: {erro}"
                        ),
                        500
                    )
                    return

                self.enviar_json(
                    {
                        "sucesso": True,
                        "total": len(apps_list),
                        "dados": apps_list
                    }
                )

                return

            # -----------------------------------------------------------------
            # API SCRIPTS
            # -----------------------------------------------------------------

            if caminho == "/api/scripts":

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                dados = (
                    carregar_scripts()
                )

                scripts = [
                    script
                    for script in dados.get(
                        "scripts",
                        []
                    )
                    if script.get(
                        "ativo",
                        True
                    )
                ]

                self.enviar_json(
                    {
                        "sucesso": True,
                        "total": len(
                            scripts
                        ),
                        "dados": scripts
                    }
                )

                return

            # -----------------------------------------------------------------
            # API KB
            # -----------------------------------------------------------------

            if caminho == "/api/kb":

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                dados = carregar_kb()

                itens = [
                    item
                    for item in dados.get(
                        "kb",
                        []
                    )
                    if item.get(
                        "ativo",
                        True
                    )
                ]

                self.enviar_json(
                    {
                        "sucesso": True,
                        "total": len(
                            itens
                        ),
                        "dados": itens
                    }
                )

                return

            # -------------------------------------------------------------------------
            # API BUSCA GLOBAL
            # -------------------------------------------------------------------------

            if caminho == "/api/busca":

                valido, usuario = usuario_da_requisicao(self)

                if not valido:
                    self.enviar_json(
                        resposta_padrao_erro("Não autenticado."),
                        401
                    )
                    return

                termo = ""
                try:
                    termo = (parametros.get("q", [""])[0] or "").strip()
                except Exception:
                    termo = ""

                if "buscar_global" in globals() and callable(buscar_global):
                    resultado = buscar_global(termo)
                else:
                    resultado = {
                        "total": 0,
                        "aplicacoes": [],
                        "scripts": [],
                        "kb": []
                    }

                try:
                    registrar_auditoria(
                        (usuario or {}).get("usuario", ""),
                        "BUSCA_GLOBAL",
                        "busca",
                        {"termo": termo}
                    )
                except Exception:
                    pass

                self.enviar_json({
                    "sucesso": True,
                    "consulta": termo,
                    "total": resultado.get("total", 0),
                    "aplicacoes": resultado.get("aplicacoes", []),
                    "scripts": resultado.get("scripts", []),
                    "kb": resultado.get("kb", [])
                })
                return

            # -----------------------------------------------------------------
            # API STATS
            # -----------------------------------------------------------------

            if caminho == "/api/stats":

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                scripts = (
                    carregar_scripts()
                )

                kb = (
                    carregar_kb()
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "dados": {
                            "aplicacoes":
                                len(apps_list),
                            "scripts":
                                len(
                                    scripts.get(
                                        "scripts",
                                        []
                                    )
                                ),
                            "kb":
                                len(
                                    kb.get(
                                        "kb",
                                        []
                                    )
                                ),
                            "sessoes_ativas":
                                len(sessoes),
                            "requisicoes":
                                estatisticas[
                                    "requisicoes"
                                ],
                            "logins_sucesso":
                                estatisticas[
                                    "logins_sucesso"
                                ],
                            "logins_falha":
                                estatisticas[
                                    "logins_falha"
                                ],
                            "inicio":
                                estatisticas[
                                    "inicio"
                                ]
                        }
                    }
                )

                return

            # -----------------------------------------------------------------
            # API AUDITORIA
            # -----------------------------------------------------------------

            if caminho == "/api/auditoria":

                valido, usuario = (
                    exigir_admin(
                        self
                    )
                )

                if not valido:

                    if usuario:

                        self.enviar_json(
                            resposta_padrao_erro(
                                "Acesso administrativo não autorizado."
                            ),
                            403
                        )

                    else:

                        self.enviar_json(
                            resposta_padrao_erro(
                                "Não autenticado."
                            ),
                            401
                        )

                    return

                dados = ler_json(
                    AUDITORIA_FILE,
                    []
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "total": len(
                            dados
                        ),
                        "dados": dados[-500:]
                    }
                )

                return

            # -----------------------------------------------------------------
            # STATIC
            # -----------------------------------------------------------------

            if caminho.startswith(
                "/static/"
            ):

                self.servir_static(
                    caminho
                )

                return

            # -----------------------------------------------------------------
            # 404
            # -----------------------------------------------------------------

            self.enviar_json(
                resposta_padrao_erro(
                    "Página não encontrada."
                ),
                404
            )

        except Exception:

            logger.exception(
                "Erro GET %s",
                self.path
            )

            try:

                self.enviar_json(
                    resposta_padrao_erro(
                        "Erro interno do servidor."
                    ),
                    500
                )

            except Exception:

                pass

    # =========================================================================
    # POST
    # =========================================================================

    def do_POST(self):

        estatisticas[
            "requisicoes"
        ] += 1

        try:

            url = urlparse(
                self.path
            )

            caminho = unquote(
                url.path
            )

            # -----------------------------------------------------------------
            # LOGIN
            # -----------------------------------------------------------------

            if caminho == "/login":

                self.processar_login()

                return

            # -----------------------------------------------------------------
            # SCRIPT
            # -----------------------------------------------------------------

            if caminho == "/api/scripts":

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                dados = (
                    self.ler_body_json()
                )

                emoji = limitar_texto(
                    dados.get(
                        "emoji",
                        ""
                    ),
                    MAX_EMOJI
                )

                titulo = limitar_texto(
                    dados.get(
                        "titulo",
                        ""
                    ),
                    MAX_TITULO
                )

                categoria = limitar_texto(
                    dados.get(
                        "categoria",
                        "Geral"
                    ),
                    MAX_CATEGORIA
                )

                texto = limitar_texto(
                    dados.get(
                        "texto",
                        ""
                    ),
                    MAX_TEXTO
                )

                ok, mensagem = (
                    validar_script(
                        titulo,
                        texto,
                        emoji,
                        categoria
                    )
                )

                if not ok:

                    self.enviar_json(
                        resposta_padrao_erro(
                            mensagem
                        ),
                        400
                    )

                    return

                dados_scripts = (
                    carregar_scripts()
                )

                novo = {
                    "id":
                        gerar_id("SCR"),
                    "emoji":
                        emoji,
                    "titulo":
                        titulo,
                    "categoria":
                        categoria,
                    "texto":
                        texto,
                    "ativo":
                        True,
                    "criado_por":
                        usuario["usuario"],
                    "criado_em":
                        agora_iso()
                }

                dados_scripts.setdefault(
                    "scripts",
                    []
                ).append(
                    novo
                )

                with scripts_lock:

                    salvar_json_com_backup(
                        SCRIPTS_FILE,
                        dados_scripts
                    )

                estatisticas[
                    "scripts_criados"
                ] += 1

                registrar_auditoria(
                    usuario["usuario"],
                    "CRIAR",
                    "script",
                    {
                        "id": novo["id"],
                        "titulo": titulo
                    }
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Script criado com sucesso.",
                        "dados":
                            novo
                    },
                    201
                )

                return

            # -----------------------------------------------------------------
            # KB
            # -----------------------------------------------------------------

            if caminho == "/api/kb":

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                dados = (
                    self.ler_body_json()
                )

                titulo = limitar_texto(
                    dados.get(
                        "titulo",
                        ""
                    ),
                    MAX_TITULO
                )

                sistema = limitar_texto(
                    dados.get(
                        "sistema",
                        ""
                    ),
                    100
                )

                categoria = limitar_texto(
                    dados.get(
                        "categoria",
                        ""
                    ),
                    MAX_CATEGORIA
                )

                descricao = limitar_texto(
                    dados.get(
                        "descricao",
                        ""
                    ),
                    MAX_TEXTO
                )

                url_item = limitar_texto(
                    dados.get(
                        "url",
                        ""
                    ),
                    500
                )

                ok, mensagem = (
                    validar_kb(
                        titulo,
                        sistema,
                        categoria,
                        descricao
                    )
                )

                if not ok:

                    self.enviar_json(
                        resposta_padrao_erro(
                            mensagem
                        ),
                        400
                    )

                    return

                dados_kb = (
                    carregar_kb()
                )

                novo = {
                    "id":
                        gerar_id("KB"),
                    "titulo":
                        titulo,
                    "sistema":
                        sistema,
                    "categoria":
                        categoria,
                    "descricao":
                        descricao,
                    "url":
                        url_item,
                    "ativo":
                        True,
                    "criado_por":
                        usuario["usuario"],
                    "criado_em":
                        agora_iso()
                }

                dados_kb.setdefault(
                    "kb",
                    []
                ).append(
                    novo
                )

                with kb_lock:

                    salvar_json_com_backup(
                        KB_FILE,
                        dados_kb
                    )

                estatisticas[
                    "kb_criados"
                ] += 1

                registrar_auditoria(
                    usuario["usuario"],
                    "CRIAR",
                    "kb",
                    {
                        "id": novo["id"],
                        "titulo": titulo
                    }
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Artigo KB criado com sucesso.",
                        "dados":
                            novo
                    },
                    201
                )

                return

            # -----------------------------------------------------------------
            # USUÁRIO
            # -----------------------------------------------------------------

            if caminho == "/api/usuarios":

                valido, usuario_admin = (
                    exigir_admin(
                        self
                    )
                )

                if not valido:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Acesso administrativo não autorizado."
                        ),
                        403
                    )

                    return

                dados = (
                    self.ler_body_json()
                )

                if not isinstance(dados, dict):
                    dados = {}

                usuario_novo = limitar_texto(
                    dados.get("usuario")
                    or dados.get("username")
                    or dados.get("login")
                    or "",
                    MAX_USUARIO
                )

                nome = limitar_texto(
                    dados.get("nome")
                    or dados.get("name")
                    or "",
                    MAX_NOME
                )

                senha = str(
                    dados.get("senha")
                    or dados.get("password")
                    or ""
                )

                permissao = limitar_texto(
                    dados.get(
                        "permissao",
                        "usuario"
                    ),
                    30
                )

                if not usuario_novo:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Usuário é obrigatório."
                        ),
                        400
                    )

                    return

                if not nome:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Nome é obrigatório."
                        ),
                        400
                    )

                    return

                senha_ok, senha_msg = (
                    validar_senha_formato(
                        senha
                    )
                )

                if not senha_ok:

                    self.enviar_json(
                        resposta_padrao_erro(
                            senha_msg
                        ),
                        400
                    )

                    return

                if permissao not in (
                    "admin",
                    "usuario"
                ):

                    permissao = "usuario"

                if encontrar_usuario(
                    usuario_novo
                ):

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Usuário já existe."
                        ),
                        409
                    )

                    return

                dados_usuarios = (
                    carregar_usuarios()
                )

                novo = {
                    "id":
                        gerar_id("USR"),
                    "usuario":
                        usuario_novo,
                    "senha":
                        gerar_hash_senha(
                            senha
                        ),
                    "nome":
                        nome,
                    "email":
                        limitar_texto(
                            dados.get(
                                "email",
                                ""
                            ),
                            200
                        ),
                    "permissao":
                        permissao,
                    "ativo":
                        True,
                    "criado_por":
                        usuario_admin[
                            "usuario"
                        ],
                    "criado_em":
                        agora_iso()
                }

                dados_usuarios.setdefault(
                    "usuarios",
                    []
                ).append(
                    novo
                )

                with usuarios_lock:

                    salvar_json_com_backup(
                        USUARIOS_FILE,
                        dados_usuarios
                    )

                registrar_auditoria(
                    usuario_admin["usuario"],
                    "CRIAR",
                    "usuario",
                    {
                        "id": novo["id"],
                        "usuario": usuario_novo
                    }
                )

                resposta = dict(
                    novo
                )

                resposta.pop(
                    "senha",
                    None
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Usuário criado com sucesso.",
                        "dados":
                            resposta
                    },
                    201
                )

                return


            # -----------------------------------------------------------------
            # UPLOAD PLANILHA DE APLICAÇÕES (admin)
            # -----------------------------------------------------------------
            if caminho in ("/api/apps/upload", "/api/aplicacoes/upload"):

                valido, usuario_admin = exigir_admin(self)

                if not valido:
                    self.enviar_json(
                        resposta_padrao_erro(
                            "Acesso administrativo não autorizado."
                        ),
                        403
                    )
                    return

                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" not in content_type:
                    self.enviar_json(
                        resposta_padrao_erro(
                            "Envie um arquivo multipart (campo arquivo)."
                        ),
                        400
                    )
                    return

                try:
                    tamanho = int(self.headers.get("Content-Length", "0"))
                except ValueError:
                    tamanho = 0

                if tamanho <= 0:
                    self.enviar_json(
                        resposta_padrao_erro("Arquivo vazio."),
                        400
                    )
                    return

                if tamanho > MAX_UPLOAD_SIZE:
                    self.enviar_json(
                        resposta_padrao_erro(
                            "Arquivo muito grande (máx. 15 MB)."
                        ),
                        400
                    )
                    return

                body = self.rfile.read(tamanho)
                nome_arq, conteudo = extrair_arquivo_multipart(
                    content_type,
                    body
                )

                if not nome_arq or not conteudo:
                    self.enviar_json(
                        resposta_padrao_erro(
                            "Nenhum arquivo encontrado no envio."
                        ),
                        400
                    )
                    return

                nome_lower = nome_arq.lower()
                if not (
                    nome_lower.endswith(".xlsx")
                    or nome_lower.endswith(".xlsm")
                ):
                    self.enviar_json(
                        resposta_padrao_erro(
                            "Envie um arquivo Excel (.xlsx)."
                        ),
                        400
                    )
                    return

                if not validar_xlsx(conteudo):
                    self.enviar_json(
                        resposta_padrao_erro(
                            "O arquivo não é um Excel .xlsx válido."
                        ),
                        400
                    )
                    return

                # Backup do arquivo atual
                try:
                    DADOS_DIR.mkdir(parents=True, exist_ok=True)
                    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
                    if EXCEL_FILE.exists():
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        destino = BACKUP_DIR / f"aplicacoes_{ts}.xlsx.bak"
                        destino.write_bytes(EXCEL_FILE.read_bytes())
                except Exception as erro_backup:
                    logger.warning(
                        "Falha ao criar backup da planilha: %s",
                        erro_backup
                    )

                # Grava novo arquivo (resiliente no Windows)
                try:
                    gravar_arquivo_bytes(EXCEL_FILE, conteudo)
                except Exception as erro_grava:
                    logger.exception("Erro ao gravar planilha")
                    self.enviar_json(
                        resposta_padrao_erro(
                            "Não foi possível salvar o arquivo. "
                            "Feche o Excel se a planilha estiver aberta e tente novamente. "
                            f"Detalhe: {erro_grava}"
                        ),
                        500
                    )
                    return

                # Valida leitura e atualiza cache em memória
                try:
                    global apps_list
                    apps = carregar_aplicacoes()
                    apps_list = apps if isinstance(apps, list) else []
                    total = len(apps_list)
                except Exception as erro_lei:
                    logger.exception("Erro ao ler planilha enviada")
                    self.enviar_json(
                        resposta_padrao_erro(
                            f"Arquivo salvo, mas falhou ao ler: {erro_lei}"
                        ),
                        500
                    )
                    return

                try:
                    registrar_auditoria(
                        usuario_admin.get("usuario", "admin"),
                        "UPLOAD",
                        "aplicacoes",
                        {
                            "arquivo": nome_arq,
                            "bytes": len(conteudo),
                            "total": total
                        }
                    )
                except Exception:
                    pass

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem": (
                            f"Planilha atualizada com sucesso. "
                            f"{total} aplicação(ões) carregada(s)."
                        ),
                        "dados": {
                            "arquivo": nome_arq,
                            "total": total
                        }
                    }
                )
                return



            if caminho in ("/api/modelo-dialogo", "/api/modelo_dialogo", "/api/modelo-dialogo/upload"):

                valido, usuario_admin = exigir_admin(self)
                if not valido:
                    self.enviar_json(
                        resposta_padrao_erro("Acesso administrativo não autorizado."),
                        403
                    )
                    return

                content_type = self.headers.get("Content-Type", "")
                dados_in = None

                if "multipart/form-data" in content_type:
                    try:
                        tamanho = int(self.headers.get("Content-Length", "0"))
                    except ValueError:
                        tamanho = 0
                    if tamanho <= 0 or tamanho > MAX_UPLOAD_SIZE:
                        self.enviar_json(
                            resposta_padrao_erro("Arquivo inválido ou muito grande."),
                            400
                        )
                        return
                    body = self.rfile.read(tamanho)
                    nome_arq, conteudo = extrair_arquivo_multipart(content_type, body)
                    if not conteudo:
                        self.enviar_json(
                            resposta_padrao_erro("Nenhum arquivo enviado."),
                            400
                        )
                        return
                    try:
                        dados_in = json.loads(conteudo.decode("utf-8"))
                    except Exception:
                        self.enviar_json(
                            resposta_padrao_erro("O arquivo precisa ser um JSON válido do modelo."),
                            400
                        )
                        return
                else:
                    dados_in = self.ler_body_json()

                if not isinstance(dados_in, dict):
                    self.enviar_json(
                        resposta_padrao_erro("JSON inválido."),
                        400
                    )
                    return

                # Aceita {scripts:[...]} ou o objeto completo do modelo
                try:
                    salvo = salvar_modelo_dialogo(
                        dados_in,
                        usuario=usuario_admin.get("usuario", "admin")
                    )
                except ValueError as erro:
                    self.enviar_json(resposta_padrao_erro(str(erro)), 400)
                    return
                except Exception as erro:
                    logger.exception("Erro ao salvar modelo diálogo")
                    self.enviar_json(
                        resposta_padrao_erro(f"Falha ao salvar: {erro}"),
                        500
                    )
                    return

                try:
                    registrar_auditoria(
                        usuario_admin.get("usuario", "admin"),
                        "UPLOAD",
                        "modelo_dialogo",
                        {"total": len(salvo.get("scripts") or [])}
                    )
                except Exception:
                    pass

                self.enviar_json({
                    "sucesso": True,
                    "mensagem": (
                        f"Modelo de Diálogo atualizado com "
                        f"{len(salvo.get('scripts') or [])} script(s)."
                    ),
                    "dados": status_modelo_dialogo(salvo)
                })
                return

            if caminho in ("/api/modelo-dialogo/revisado", "/api/modelo_dialogo/revisado"):

                valido, usuario_admin = exigir_admin(self)
                if not valido:
                    self.enviar_json(
                        resposta_padrao_erro("Acesso administrativo não autorizado."),
                        403
                    )
                    return

                atual = carregar_modelo_dialogo()
                atual["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
                atual["atualizado_por"] = usuario_admin.get("usuario", "admin")
                # mantém scripts
                body = {}
                try:
                    body = self.ler_body_json() or {}
                except Exception:
                    body = {}
                if isinstance(body, dict) and body.get("intervalo_dias"):
                    try:
                        atual["intervalo_dias"] = max(1, int(body["intervalo_dias"]))
                    except Exception:
                        pass
                salvar_json(MODELO_DIALOGO_FILE, atual)
                self.enviar_json({
                    "sucesso": True,
                    "mensagem": "Data de revisão atualizada.",
                    "dados": status_modelo_dialogo(atual)
                })
                return


            self.enviar_json(
                resposta_padrao_erro(
                    "Endpoint não encontrado."
                ),
                404
            )

        except json.JSONDecodeError:

            self.enviar_json(
                resposta_padrao_erro(
                    "JSON inválido."
                ),
                400
            )

        except ValueError as erro:

            self.enviar_json(
                resposta_padrao_erro(
                    str(erro)
                ),
                400
            )

        except Exception:

            logger.exception(
                "Erro POST %s",
                self.path
            )

            self.enviar_json(
                resposta_padrao_erro(
                    "Erro interno do servidor."
                ),
                500
            )

    # =========================================================================
    # PUT
    # =========================================================================

    def do_PUT(self):

        estatisticas[
            "requisicoes"
        ] += 1

        try:

            url = urlparse(
                self.path
            )

            caminho = unquote(
                url.path
            )

            # -----------------------------------------------------------------
            # SCRIPT
            # -----------------------------------------------------------------

            if caminho.startswith(
                "/api/scripts/"
            ):

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                id_script = caminho.split(
                    "/"
                )[-1]

                if not id_script:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "ID inválido."
                        ),
                        400
                    )

                    return

                dados = (
                    self.ler_body_json()
                )

                titulo = limitar_texto(
                    dados.get(
                        "titulo",
                        ""
                    ),
                    MAX_TITULO
                )

                texto = limitar_texto(
                    dados.get(
                        "texto",
                        ""
                    ),
                    MAX_TEXTO
                )

                emoji = limitar_texto(
                    dados.get(
                        "emoji",
                        ""
                    ),
                    MAX_EMOJI
                )

                categoria = limitar_texto(
                    dados.get(
                        "categoria",
                        "Geral"
                    ),
                    MAX_CATEGORIA
                )

                ok, mensagem = (
                    validar_script(
                        titulo,
                        texto,
                        emoji,
                        categoria
                    )
                )

                if not ok:

                    self.enviar_json(
                        resposta_padrao_erro(
                            mensagem
                        ),
                        400
                    )

                    return

                dados_scripts = (
                    carregar_scripts()
                )

                encontrado = None

                for script in dados_scripts.get(
                    "scripts",
                    []
                ):

                    if script.get(
                        "id"
                    ) == id_script:

                        encontrado = script

                        break

                if not encontrado:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Script não encontrado."
                        ),
                        404
                    )

                    return

                encontrado[
                    "emoji"
                ] = emoji

                encontrado[
                    "titulo"
                ] = titulo

                encontrado[
                    "categoria"
                ] = categoria

                encontrado[
                    "texto"
                ] = texto

                encontrado[
                    "alterado_por"
                ] = usuario[
                    "usuario"
                ]

                encontrado[
                    "alterado_em"
                ] = agora_iso()

                if "ativo" in dados:

                    encontrado[
                        "ativo"
                    ] = bool(
                        dados[
                            "ativo"
                        ]
                    )

                with scripts_lock:

                    salvar_json_com_backup(
                        SCRIPTS_FILE,
                        dados_scripts
                    )

                estatisticas[
                    "scripts_alterados"
                ] += 1

                registrar_auditoria(
                    usuario["usuario"],
                    "ALTERAR",
                    "script",
                    {
                        "id":
                            id_script,
                        "titulo":
                            titulo
                    }
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Script atualizado.",
                        "dados":
                            encontrado
                    }
                )

                return

            # -----------------------------------------------------------------
            # KB
            # -----------------------------------------------------------------

            if caminho.startswith(
                "/api/kb/"
            ):

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                id_kb = caminho.split(
                    "/"
                )[-1]

                dados = (
                    self.ler_body_json()
                )

                dados_kb = (
                    carregar_kb()
                )

                encontrado = None

                for item in dados_kb.get(
                    "kb",
                    []
                ):

                    if item.get(
                        "id"
                    ) == id_kb:

                        encontrado = item

                        break

                if not encontrado:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Artigo KB não encontrado."
                        ),
                        404
                    )

                    return

                titulo = limitar_texto(
                    dados.get(
                        "titulo",
                        encontrado.get(
                            "titulo",
                            ""
                        )
                    ),
                    MAX_TITULO
                )

                sistema = limitar_texto(
                    dados.get(
                        "sistema",
                        encontrado.get(
                            "sistema",
                            ""
                        )
                    ),
                    100
                )

                categoria = limitar_texto(
                    dados.get(
                        "categoria",
                        encontrado.get(
                            "categoria",
                            ""
                        )
                    ),
                    MAX_CATEGORIA
                )

                descricao = limitar_texto(
                    dados.get(
                        "descricao",
                        encontrado.get(
                            "descricao",
                            ""
                        )
                    ),
                    MAX_TEXTO
                )

                url_item = limitar_texto(
                    dados.get(
                        "url",
                        encontrado.get(
                            "url",
                            ""
                        )
                    ),
                    500
                )

                ok, mensagem = (
                    validar_kb(
                        titulo,
                        sistema,
                        categoria,
                        descricao
                    )
                )

                if not ok:

                    self.enviar_json(
                        resposta_padrao_erro(
                            mensagem
                        ),
                        400
                    )

                    return

                encontrado[
                    "titulo"
                ] = titulo

                encontrado[
                    "sistema"
                ] = sistema

                encontrado[
                    "categoria"
                ] = categoria

                encontrado[
                    "descricao"
                ] = descricao

                encontrado[
                    "url"
                ] = url_item

                encontrado[
                    "alterado_por"
                ] = usuario[
                    "usuario"
                ]

                encontrado[
                    "alterado_em"
                ] = agora_iso()

                if "ativo" in dados:

                    encontrado[
                        "ativo"
                    ] = bool(
                        dados[
                            "ativo"
                        ]
                    )

                with kb_lock:

                    salvar_json_com_backup(
                        KB_FILE,
                        dados_kb
                    )

                estatisticas[
                    "kb_alterados"
                ] += 1

                registrar_auditoria(
                    usuario["usuario"],
                    "ALTERAR",
                    "kb",
                    {
                        "id":
                            id_kb,
                        "titulo":
                            titulo
                    }
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Artigo KB atualizado.",
                        "dados":
                            encontrado
                    }
                )

                return

            # -----------------------------------------------------------------
            # USUÁRIO
            # -----------------------------------------------------------------

            if caminho.startswith(
                "/api/usuarios/"
            ):

                usuario_admin = (
                    exigir_admin(
                        self
                    )
                )

                if not usuario_admin[0]:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Acesso administrativo não autorizado."
                        ),
                        403
                    )

                    return

                admin = usuario_admin[1]

                id_usuario = caminho.split(
                    "/"
                )[-1]

                dados = (
                    self.ler_body_json()
                )

                dados_usuarios = (
                    carregar_usuarios()
                )

                encontrado = None

                for item in dados_usuarios.get(
                    "usuarios",
                    []
                ):

                    if item.get(
                        "id"
                    ) == id_usuario:

                        encontrado = item

                        break

                if not encontrado:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Usuário não encontrado."
                        ),
                        404
                    )

                    return

                if "nome" in dados:

                    encontrado[
                        "nome"
                    ] = limitar_texto(
                        dados[
                            "nome"
                        ],
                        MAX_NOME
                    )

                if "email" in dados:

                    encontrado[
                        "email"
                    ] = limitar_texto(
                        dados[
                            "email"
                        ],
                        200
                    )

                if "permissao" in dados:

                    permissao = (
                        limitar_texto(
                            dados[
                                "permissao"
                            ],
                            30
                        )
                    )

                    if permissao in (
                        "admin",
                        "usuario"
                    ):

                        encontrado[
                            "permissao"
                        ] = permissao

                if "ativo" in dados:

                    encontrado[
                        "ativo"
                    ] = bool(
                        dados[
                            "ativo"
                        ]
                    )

                if dados.get(
                    "senha"
                ):

                    senha = str(
                        dados[
                            "senha"
                        ]
                    )

                    ok, msg = (
                        validar_senha_formato(
                            senha
                        )
                    )

                    if not ok:

                        self.enviar_json(
                            resposta_padrao_erro(
                                msg
                            ),
                            400
                        )

                        return

                    encontrado[
                        "senha"
                    ] = gerar_hash_senha(
                        senha
                    )

                encontrado[
                    "alterado_por"
                ] = admin[
                    "usuario"
                ]

                encontrado[
                    "alterado_em"
                ] = agora_iso()

                with usuarios_lock:

                    salvar_json_com_backup(
                        USUARIOS_FILE,
                        dados_usuarios
                    )

                registrar_auditoria(
                    admin["usuario"],
                    "ALTERAR",
                    "usuario",
                    {
                        "id":
                            id_usuario
                    }
                )

                resposta = dict(
                    encontrado
                )

                resposta.pop(
                    "senha",
                    None
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Usuário atualizado.",
                        "dados":
                            resposta
                    }
                )

                return

            self.enviar_json(
                resposta_padrao_erro(
                    "Endpoint não encontrado."
                ),
                404
            )

        except json.JSONDecodeError:

            self.enviar_json(
                resposta_padrao_erro(
                    "JSON inválido."
                ),
                400
            )

        except ValueError as erro:

            self.enviar_json(
                resposta_padrao_erro(
                    str(erro)
                ),
                400
            )

        except Exception:

            logger.exception(
                "Erro PUT %s",
                self.path
            )

            self.enviar_json(
                resposta_padrao_erro(
                    "Erro interno do servidor."
                ),
                500
            )

    # =========================================================================
    # DELETE
    # =========================================================================

    def do_DELETE(self):

        estatisticas[
            "requisicoes"
        ] += 1

        try:

            url = urlparse(
                self.path
            )

            caminho = unquote(
                url.path
            )

            # -----------------------------------------------------------------
            # SCRIPT
            # -----------------------------------------------------------------

            if caminho.startswith(
                "/api/scripts/"
            ):

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                id_script = caminho.split(
                    "/"
                )[-1]

                dados_scripts = (
                    carregar_scripts()
                )

                scripts = (
                    dados_scripts.get(
                        "scripts",
                        []
                    )
                )

                encontrado = None

                for script in scripts:

                    if script.get(
                        "id"
                    ) == id_script:

                        encontrado = script

                        break

                if not encontrado:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Script não encontrado."
                        ),
                        404
                    )

                    return

                # Exclusão lógica
                encontrado[
                    "ativo"
                ] = False

                encontrado[
                    "excluido_por"
                ] = usuario[
                    "usuario"
                ]

                encontrado[
                    "excluido_em"
                ] = agora_iso()

                with scripts_lock:

                    salvar_json_com_backup(
                        SCRIPTS_FILE,
                        dados_scripts
                    )

                estatisticas[
                    "scripts_excluidos"
                ] += 1

                registrar_auditoria(
                    usuario["usuario"],
                    "EXCLUIR",
                    "script",
                    {
                        "id":
                            id_script
                    }
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Script desativado com sucesso."
                    }
                )

                return

            # -----------------------------------------------------------------
            # KB
            # -----------------------------------------------------------------

            if caminho.startswith(
                "/api/kb/"
            ):

                usuario = (
                    self.exigir_login_api()
                )

                if not usuario:

                    return

                id_kb = caminho.split(
                    "/"
                )[-1]

                dados_kb = (
                    carregar_kb()
                )

                encontrado = None

                for item in dados_kb.get(
                    "kb",
                    []
                ):

                    if item.get(
                        "id"
                    ) == id_kb:

                        encontrado = item

                        break

                if not encontrado:

                    self.enviar_json(
                        resposta_padrao_erro(
                            "Artigo KB não encontrado."
                        ),
                        404
                    )

                    return

                encontrado[
                    "ativo"
                ] = False

                encontrado[
                    "excluido_por"
                ] = usuario[
                    "usuario"
                ]

                encontrado[
                    "excluido_em"
                ] = agora_iso()

                with kb_lock:

                    salvar_json_com_backup(
                        KB_FILE,
                        dados_kb
                    )

                estatisticas[
                    "kb_excluidos"
                ] += 1

                registrar_auditoria(
                    usuario["usuario"],
                    "EXCLUIR",
                    "kb",
                    {
                        "id":
                            id_kb
                    }
                )

                self.enviar_json(
                    {
                        "sucesso": True,
                        "mensagem":
                            "Artigo KB desativado com sucesso."
                    }
                )

                return

            self.enviar_json(
                resposta_padrao_erro(
                    "Endpoint não encontrado."
                ),
                404
            )

        except Exception:

            logger.exception(
                "Erro DELETE %s",
                self.path
            )

            self.enviar_json(
                resposta_padrao_erro(
                    "Erro interno do servidor."
                ),
                500
            )

    # =========================================================================
    # LOGIN
    # =========================================================================

    def processar_login(self):

        dados = (
            self.ler_body_json()
        )

        usuario = limitar_texto(
            dados.get(
                "usuario",
                ""
            ),
            MAX_USUARIO
        )

        senha = str(
            dados.get(
                "senha",
                ""
            )
        )

        if not usuario or not senha:

            self.enviar_json(
                resposta_padrao_erro(
                    "Usuário e senha são obrigatórios."
                ),
                400
            )

            return

        valido, nome, permissao = (
            validar_credenciais(
                usuario,
                senha
            )
        )

        if not valido:

            self.enviar_json(
                resposta_padrao_erro(
                    "Usuário ou senha incorretos."
                ),
                401
            )

            return

        token = criar_sessao(
            usuario,
            nome,
            permissao
        )

        cookie = (
            f"session_token={token}; "
            "Path=/; "
            "HttpOnly; "
            "SameSite=Lax"
            f"{flag_cookie_secure()}"
        )

        self.enviar_json(
            {
                "sucesso": True,
                "mensagem":
                    f"Bem-vindo, {nome}!",
                "usuario": {
                    "usuario":
                        usuario,
                    "nome":
                        nome,
                    "permissao":
                        permissao
                }
            },
            200,
            {
                "Set-Cookie":
                    cookie
            }
        )

    # =========================================================================
    # TEMPLATE
    # =========================================================================

    def ler_template(
        self,
        nome: str
    ) -> str:

        caminho = (
            TEMPLATES_DIR /
            nome
        )

        if not caminho.exists():

            return ""

        try:

            return caminho.read_text(
                encoding="utf-8"
            )

        except Exception as erro:

            logger.error(
                "Erro ao ler template %s: %s",
                nome,
                erro
            )

            return ""

    # =========================================================================
    # LOGIN
    # =========================================================================

    def pagina_login(self):

        html = self.ler_template(
            "login.html"
        )

        if not html:

            html = """
<!DOCTYPE html>
<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>Central TIC</title>

<style>

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    font-family: Arial, sans-serif;
    background:
        linear-gradient(
            135deg,
            #0b2742,
            #155b88
        );
}

.container {
    width: 92%;
    max-width: 430px;
    background: #ffffff;
    border-radius: 18px;
    padding: 35px;
    box-shadow:
        0 20px 60px
        rgba(0,0,0,.25);
}

.logo {
    text-align: center;
    font-size: 48px;
}

h1 {
    margin:
        10px 0 5px;
    text-align: center;
    color: #12395b;
}

.subtitulo {
    text-align: center;
    color: #667085;
    margin-bottom: 25px;
}

label {
    display: block;
    margin:
        15px 0 7px;
    font-weight: bold;
}

input {
    width: 100%;
    padding: 13px;
    border:
        1px solid #d0d5dd;
    border-radius: 9px;
    font-size: 15px;
}

button {
    width: 100%;
    margin-top: 22px;
    padding: 14px;
    border: 0;
    border-radius: 9px;
    background: #12395b;
    color: white;
    font-size: 16px;
    font-weight: bold;
    cursor: pointer;
}

button:hover {
    background: #0d2b45;
}

.mensagem {
    min-height: 22px;
    margin-top: 15px;
    text-align: center;
}

.info {
    margin-top: 25px;
    padding: 14px;
    border-radius: 9px;
    background: #f2f4f7;
    color: #475467;
    font-size: 13px;
}

</style>

</head>

<body>

<div class="container">

<div class="logo">
🛠️
</div>

<h1>
Central TIC
</h1>

<div class="subtitulo">
Central de Atendimento TIC
</div>

<form id="loginForm">

<label>
Usuário
</label>

<input
    id="usuario"
    type="text"
    autocomplete="username"
    required
>

<label>
Senha
</label>

<input
    id="senha"
    type="password"
    autocomplete="current-password"
    required
>

<button type="submit">
Entrar
</button>

</form>

<div
    id="mensagem"
    class="mensagem"
></div>

<div class="info">

<strong>Acesso inicial</strong>

<br><br>

Administrador:
<strong>admin</strong>

<br>

Senha:
<strong>admin123</strong>

<br><br>

Agente:
<strong>agente</strong>

<br>

Senha:
<strong>agente123</strong>

</div>

</div>

<script>

const form =
    document.getElementById(
        "loginForm"
    );

const mensagem =
    document.getElementById(
        "mensagem"
    );

form.addEventListener(
    "submit",
    async function(event) {

        event.preventDefault();

        mensagem.textContent =
            "Autenticando...";

        try {

            const resposta =
                await fetch(
                    "/login",
                    {
                        method: "POST",
                        credentials:
                            "same-origin",
                        headers: {
                            "Content-Type":
                                "application/json"
                        },
                        body:
                            JSON.stringify({
                                usuario:
                                    document
                                    .getElementById(
                                        "usuario"
                                    )
                                    .value
                                    .trim(),

                                senha:
                                    document
                                    .getElementById(
                                        "senha"
                                    )
                                    .value
                            })
                    }
                );

            const dados =
                await resposta.json();

            if (!resposta.ok) {

                mensagem.textContent =
                    dados.erro ||
                    "Falha no login.";

                return;
            }

            mensagem.textContent =
                "Login realizado.";

            window.location.href =
                "/";

        } catch (erro) {

            console.error(
                erro
            );

            mensagem.textContent =
                "Erro de comunicação com o servidor.";
        }

    }
);

</script>

</body>

</html>
"""

        self.enviar_html(
            html
        )

    # =========================================================================
    # HOME
    # =========================================================================

    def pagina_home(
        self,
        usuario
    ):

        html = self.ler_template(
            "index.html"
        )

        if not html:

            nome = (
                usuario.get(
                    "nome",
                    usuario.get(
                        "usuario",
                        ""
                    )
                )
            )

            admin_link = ""

            if usuario.get(
                "permissao"
            ) == "admin":

                admin_link = """
<a href="/admin">
⚙️ Administração
</a>
"""

            html = f"""
<!DOCTYPE html>

<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>
Central TIC
</title>

<style>

body {{
    margin: 0;
    font-family: Arial, sans-serif;
    background: #f4f6f8;
}}

header {{
    background: #12395b;
    color: white;
    padding: 25px;
}}

main {{
    max-width: 1200px;
    margin: auto;
    padding: 30px;
}}

.card {{
    background: white;
    padding: 25px;
    margin-bottom: 20px;
    border-radius: 14px;
    box-shadow:
        0 4px 15px
        rgba(0,0,0,.07);
}}

.menu {{
    display: grid;
    grid-template-columns:
        repeat(
            auto-fit,
            minmax(220px, 1fr)
        );
    gap: 15px;
}}

.menu a {{
    padding: 20px;
    border-radius: 12px;
    background: #eef3f7;
    color: #12395b;
    text-decoration: none;
    font-weight: bold;
}}

.menu a:hover {{
    background: #dfe9f1;
}}

.sair {{
    display: inline-block;
    padding: 12px 20px;
    background: #b42318;
    color: white;
    text-decoration: none;
    border-radius: 9px;
}}

</style>

</head>

<body>

<header>

<h1>
Central de Atendimento TIC
</h1>

<p>
Versão {VERSAO}
</p>

</header>

<main>

<div class="card">

<h2>
Olá, {nome}!
</h2>

<p>
Usuário:
<strong>
{usuario["usuario"]}
</strong>
</p>

<p>
Perfil:
<strong>
{usuario["permissao"]}
</strong>
</p>

</div>

<div class="card">

<h2>
Recursos
</h2>

<div class="menu">

<a href="/dialogo">
📋 Modelo de Diálogo
</a>

<a href="/api/apps">
🖥️ Aplicações
</a>

<a href="/api/scripts">
📜 Scripts
</a>

<a href="/api/kb">
📚 Knowledge Base
</a>

<a href="/api/stats">
📊 Dashboard
</a>

{admin_link}

</div>

</div>

<div class="card">

<a
    class="sair"
    href="/logout"
>
Sair
</a>

</div>

</main>

</body>

</html>
"""

        self.enviar_html(
            html
        )

    # =========================================================================
    # ADMIN
    # =========================================================================

    def pagina_admin(
        self,
        usuario
    ):

        html = self.ler_template(
            "admin.html"
        )

        if not html:

            scripts = carregar_scripts()

            kb = carregar_kb()

            usuarios = carregar_usuarios()

            try:
                apps_list = carregar_aplicacoes()
            except Exception:
                apps_list = []
            if isinstance(apps_list, dict):
                apps_list = apps_list.get("aplicacoes") or apps_list.get("dados") or []
            if not isinstance(apps_list, list):
                apps_list = []

            html = f"""
<!DOCTYPE html>

<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>
Administração - Central TIC
</title>

<style>

body {{
    margin: 0;
    padding: 30px;
    font-family: Arial, sans-serif;
    background: #f4f6f8;
}}

.container {{
    max-width: 1100px;
    margin: auto;
}}

.card {{
    background: white;
    padding: 25px;
    margin-bottom: 20px;
    border-radius: 14px;
}}

.grid {{
    display: grid;
    grid-template-columns:
        repeat(
            auto-fit,
            minmax(200px, 1fr)
        );
    gap: 15px;
}}

.box {{
    padding: 20px;
    background: #eef3f7;
    border-radius: 10px;
}}

.numero {{
    font-size: 32px;
    font-weight: bold;
    color: #12395b;
}}

a {{
    color: #12395b;
}}

</style>

</head>

<body>

<div class="container">

<div class="card">

<h1>
⚙️ Administração
</h1>

<p>
Administrador:
<strong>
{usuario["nome"]}
</strong>
</p>

</div>

<div class="card">

<h2>
📊 Visão geral
</h2>

<div class="grid">

<div class="box">

<div class="numero">
{len(apps_list)}
</div>

Aplicações

</div>

<div class="box">

<div class="numero">
{len(scripts.get("scripts", []))}
</div>

Scripts

</div>

<div class="box">

<div class="numero">
{len(kb.get("kb", []))}
</div>

Knowledge Base

</div>

<div class="box">

<div class="numero">
{len(usuarios.get("usuarios", []))}
</div>

Usuários

</div>

</div>

</div>

<div class="card">

<h2>
🔧 APIs administrativas
</h2>

<p>
<a href="/api/auditoria">
Consultar auditoria
</a>
</p>

<p>
<a href="/api/stats">
Consultar estatísticas
</a>
</p>

</div>

<div class="card">

<a href="/">
← Voltar para a Central
</a>

</div>

</div>

</body>

</html>
"""

        self.enviar_html(
            html
        )

    # =========================================================================
    # DIÁLOGO
    # =========================================================================

    def pagina_dialogo(
        self,
        usuario
    ):

        html = self.ler_template(
            "dialogo.html"
        )

        if not html:

            html = f"""
<!DOCTYPE html>

<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>
Modelo de Diálogo
</title>

<style>

body {{
    margin: 0;
    padding: 30px;
    font-family: Arial, sans-serif;
    background: #f4f6f8;
}}

.container {{
    max-width: 1100px;
    margin: auto;
}}

.card {{
    background: white;
    padding: 25px;
    border-radius: 14px;
}}

h1 {{
    color: #12395b;
}}

.info {{
    padding: 15px;
    background: #eef3f7;
    border-radius: 10px;
}}

a {{
    color: #12395b;
}}

</style>

</head>

<body>

<div class="container">

<div class="card">

<h1>
📋 Modelo de Diálogo
</h1>

<p>
Atendente:
<strong>
{usuario["nome"]}
</strong>
</p>

<div class="info">

Esta interface será conectada aos
Scripts, Knowledge Base e Aplicações
nas próximas etapas da versão 4.0.

</div>

<br>

<a href="/">
← Voltar para a Central
</a>

</div>

</div>

</body>

</html>
"""

        self.enviar_html(
            html
        )

    # =========================================================================
    # STATIC
    # =========================================================================


    def servir_arquivo(self, nome: str, exige_login: bool = True):
        """Serve um template HTML da pasta templates/."""
        if exige_login:
            valido, usuario = usuario_da_requisicao(self)
            if not valido:
                self.redirecionar("/login")
                return

        html = self.ler_template(nome)
        if not html:
            self.enviar_json(
                resposta_padrao_erro(
                    f"Template não encontrado: {nome}"
                ),
                404
            )
            return

        self.enviar_html(html)

    def servir_static(
        self,
        caminho
    ):

        relativo = caminho.replace(
            "/static/",
            "",
            1
        )

        arquivo = (
            STATIC_DIR /
            relativo
        ).resolve()

        static_resolvido = (
            STATIC_DIR.resolve()
        )

        try:

            arquivo.relative_to(
                static_resolvido
            )

        except ValueError:

            self.enviar_json(
                resposta_padrao_erro(
                    "Arquivo inválido."
                ),
                403
            )

            return

        if not arquivo.exists():

            self.enviar_json(
                resposta_padrao_erro(
                    "Arquivo não encontrado."
                ),
                404
            )

            return

        if not arquivo.is_file():

            self.enviar_json(
                resposta_padrao_erro(
                    "Recurso inválido."
                ),
                403
            )

            return

        extensao = (
            arquivo.suffix.lower()
        )

        tipos = {

            ".css":
                "text/css; charset=utf-8",

            ".js":
                "application/javascript; charset=utf-8",

            ".json":
                "application/json; charset=utf-8",

            ".html":
                "text/html; charset=utf-8",

            ".png":
                "image/png",

            ".jpg":
                "image/jpeg",

            ".jpeg":
                "image/jpeg",

            ".webp":
                "image/webp",

            ".svg":
                "image/svg+xml",

            ".ico":
                "image/x-icon",

            ".woff":
                "font/woff",

            ".woff2":
                "font/woff2"
        }

        tipo = tipos.get(
            extensao,
            mimetypes.guess_type(
                str(arquivo)
            )[0]
            or
            "application/octet-stream"
        )

        try:

            conteudo = (
                arquivo.read_bytes()
            )

            self.send_response(
                200
            )

            self.send_header(
                "Content-Type",
                tipo
            )

            self.send_header(
                "Content-Length",
                str(
                    len(conteudo)
                )
            )

            self.send_header(
                "Cache-Control",
                "public, max-age=3600"
            )

            self.enviar_headers_seguranca()

            self.end_headers()

            self.wfile.write(
                conteudo
            )

        except Exception as erro:

            logger.error(
                "Erro ao servir static %s: %s",
                arquivo,
                erro
            )

            self.enviar_json(
                resposta_padrao_erro(
                    "Erro ao carregar recurso."
                ),
                500
            )

    # =========================================================================
    # LOG HTTP
    # =========================================================================

    def log_message(
        self,
        formato,
        *args
    ):

        logger.info(
            "%s - %s",
            self.address_string(),
            formato % args
        )


# =============================================================================
# PORTA
# =============================================================================

def encontrar_porta(
    preferida: int,
    tentativas: int = 20
) -> int:

    for porta in range(
        preferida,
        preferida + tentativas
    ):

        sock = socket.socket(
            socket.AF_INET,
            socket.SOCK_STREAM
        )

        try:

            sock.setsockopt(
                socket.SOL_SOCKET,
                socket.SO_REUSEADDR,
                1
            )

            sock.bind(
                ("", porta)
            )

            sock.close()

            return porta

        except OSError:

            sock.close()

    raise RuntimeError(
        "Nenhuma porta disponível."
    )


# =============================================================================
# LIMPEZA DE SESSÕES
# =============================================================================

def limpar_sessoes_expiradas():

    agora = datetime.now().timestamp()

    removidas = 0

    with sessoes_lock:

        tokens = list(
            sessoes.keys()
        )

        for token in tokens:

            sessao = sessoes.get(
                token
            )

            if not sessao:

                continue

            if (
                agora
                >
                float(
                    sessao.get(
                        "expiracao",
                        0
                    )
                )
            ):

                sessoes.pop(
                    token,
                    None
                )

                removidas += 1

    if removidas:

        logger.info(
            "Sessões expiradas removidas: %s",
            removidas
        )


def iniciar_limpeza_sessoes():

    def rotina():

        while True:

            try:

                limpar_sessoes_expiradas()

            except Exception:

                logger.exception(
                    "Erro na limpeza de sessões."
                )

            threading.Event().wait(
                300
            )

    thread = threading.Thread(
        target=rotina,
        daemon=True
    )

    thread.start()


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================

def inicializar_dados():

    # -------------------------------------------------------------------------
    # Usuários
    # -------------------------------------------------------------------------

    if not USUARIOS_FILE.exists():

        criar_usuarios_padrao()

    else:

        garantir_usuarios_padrao()

    # -------------------------------------------------------------------------
    # Scripts
    # -------------------------------------------------------------------------

    if not SCRIPTS_FILE.exists():

        criar_scripts_padrao()

    # -------------------------------------------------------------------------
    # KB
    # -------------------------------------------------------------------------

    if not KB_FILE.exists():

        criar_kb_padrao()

    # -------------------------------------------------------------------------
    # Auditoria
    # -------------------------------------------------------------------------

    if not AUDITORIA_FILE.exists():

        salvar_json(
            AUDITORIA_FILE,
            []
        )


# =============================================================================
# SERVIDOR
# =============================================================================

class Servidor(
    socketserver.ThreadingTCPServer
):

    allow_reuse_address = True

    daemon_threads = True


# =============================================================================
# MAIN
# =============================================================================

def main():

    global apps_list

    logger.info(
        "=" * 70
    )

    logger.info(
        "%s",
        NOME_SISTEMA
    )

    logger.info(
        "Versão %s",
        VERSAO
    )

    logger.info(
        "Inicializando..."
    )

    inicializar_dados()

    apps_list = (
        carregar_aplicacoes()
    )

    iniciar_limpeza_sessoes()

    # Render/Railway etc. exigem a porta EXATA da variável PORT
    if os.getenv("RENDER") or os.getenv("FORCE_PORT", "").lower() in ("1", "true", "yes"):
        porta = int(os.getenv("PORT", str(PORTA_PREFERIDA)))
    else:
        porta = encontrar_porta(PORTA_PREFERIDA)

    logger.info(
        "Aplicações carregadas: %s",
        len(apps_list)
    )

    logger.info(
        "Servidor iniciado."
    )

    logger.info(
        "Host: %s | Porta: %s",
        HOST,
        porta
    )
    logger.info(
        "Home local: http://localhost:%s",
        porta
    )

    logger.info(
        "Login: http://localhost:%s/login",
        porta
    )

    logger.info(
        "Admin: http://localhost:%s/admin",
        porta
    )

    logger.info(
        "=" * 70
    )

    print()

    print(
        "=" * 70
    )

    print(
        " CENTRAL DE ATENDIMENTO TIC - PETROBRAS"
    )

    print(
        f" VERSÃO {VERSAO}"
    )

    print(
        "=" * 70
    )

    print()

    print(
        f" Home:       http://localhost:{porta}"
    )

    print(
        f" Login:      http://localhost:{porta}/login"
    )

    print(
        f" Admin:      http://localhost:{porta}/admin"
    )

    print(
        f" Diálogo:    http://localhost:{porta}/dialogo"
    )

    print(
        f" Health:     http://localhost:{porta}/api/health"
    )

    print()

    print(
        f" Aplicações carregadas: "
        f"{len(apps_list)}"
    )

    print()

    print(
        " Usuários iniciais:"
    )

    print(
        "   admin  / admin123"
    )

    print(
        "   agente / agente123"
    )

    print()

    print(
        " Pressione CTRL+C para encerrar."
    )

    print()

    print(
        "=" * 70
    )

    print()

    servidor = Servidor(
        (
            HOST,
            porta
        ),
        Handler
    )

    try:

        servidor.serve_forever()

    except KeyboardInterrupt:

        print()

        print(
            "Servidor encerrado."
        )

        logger.info(
            "Servidor encerrado pelo usuário."
        )

    finally:

        servidor.server_close()


# =============================================================================
# EXECUÇÃO
# =============================================================================

if __name__ == "__main__":

    main()
