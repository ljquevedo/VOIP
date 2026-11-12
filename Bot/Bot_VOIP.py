"""
Bot de Telegram - Guia de Ramais VOIP
Hospital Universitário São Francisco de Paula (2026)

Uso:
  - Mensagem direta ao bot: digite parte do nome do setor ou o número do ramal.
  - Modo inline: em qualquer chat, digite "@seu_bot nome_do_setor".

Configuração:
  1. Defina a variável de ambiente TELEGRAM_BOT_TOKEN com o token do @BotFather.
  2. (Opcional, mas necessário para o modo inline) No @BotFather, use /setinline
     e ative o modo inline para o seu bot.
  3. Instale as dependências: pip install -r requirements.txt --break-system-packages
  4. Rode: python bot.py
"""

import asyncio
import csv
import datetime
import io
import json
import logging
import os
import platform
import re
import subprocess
import time
import unicodedata
from collections import Counter
from uuid import uuid4

from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InlineQueryResultArticle,
    InputTextMessageContent,
    Update,
)
from telegram.constants import ParseMode
from telegram.error import TelegramError

try:
    from PIL import Image
    from pyzbar.pyzbar import decode as decode_barcodes
    LEITURA_CODIGO_BARRAS_DISPONIVEL = True
except Exception:
    # Além de ImportError (biblioteca não instalada), o pyzbar pode falhar
    # ao carregar a DLL nativa do zbar no Windows (OSError/FileNotFoundError
    # — ex.: "libzbar-64.dll" ausente/incompatível), o que derrubaria o bot
    # inteiro na inicialização se não fosse pego aqui. Nesses casos a leitura
    # de código de barras só fica indisponível — o resto do bot funciona normalmente.
    LEITURA_CODIGO_BARRAS_DISPONIVEL = False

from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    InlineQueryHandler,
    MessageHandler,
    filters,
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ── Base de dados de ramais ──────────────────────────────────────────────
# (Nome do setor, Ramal, Sigla da unidade, Prédio, Andar)
# Prédio/Andar vêm da planilha Ramais_HUSFP_2026.xlsx, aba "Consolidado"
# (cruzados pelo número do ramal).
DATA = [
    ("Almoxarifado (Central de Materiais)", "8560", "HU", "HU", "1º"),
    ("Almoxarifado (Recebimentos)", "8560", "HU", "HU", "1º"),
    ("Bloco Cirúrgico (Farmácia)", "8350", "HU", "HU", "5º"),
    ("Bloco Cirúrgico (Marcação)", "8376", "HU", "HU", "5º"),
    ("Bloco Cirúrgico (Posto de Enfermagem)", "8591", "HU", "HU", "5º"),
    ("Bloco Cirúrgico (Sala de Recuperação)", "8347", "HU", "HU", "5º"),
    ("CAF", "8331", "HU", "HU", "1º"),
    ("Casa da Gestante", "8360", "HU", "Casa da gestante", "1º"),
    ("CDI - Administrativo", "8496", "HU", "HU", "2º"),
    ("CDI (Convênios)", "8310", "HU", "HU", "2º"),
    ("CDI (Sala 01)", "8449", "HU", "HU", "2º"),
    ("CDI (Sala 02)", "8563", "HU", "HU", "2º"),
    ("CDI (Sala 03)", "8380", "HU", "HU", "2º"),
    ("CDI (SUS)", "8318", "HU", "HU", "1º"),
    ("Central de Cadastro - MV", "8507", "ADM", "ADM", "1º"),
    ("Central de Cadastro SIGH", "8506", "ADM", "ADM", "2º"),
    ("Centro Acadêmico (Recepção)", "8386", "HU", "HU", "1º"),
    ("Clínica Cirúrgica (Escriturários)", "8557", "HU", "HU", "5º"),
    ("Clínica Cirúrgica (Posto de Enfermagem)", "8320", "HU", "HU", "5º"),
    ("CME", "8336", "HU", "HU", "4º"),
    ("Comunicação e Marketing", "8368", "PA", "PA", "2º"),
    ("Compras", "8326", "HU", "HU", "1º"),
    ("Conceição (Posto de Enfermagem)", "8575", "HU", "HU", "5º"),
    ("Conceição (Sala de Prescrição)", "8324", "HU", "HU", "5º"),
    ("Conforto Médico (Clínica Médica)", "8357", "HU", "HU", "6º"),
    ("Conforto Médico (Ginecologia)", "8361", "HU", "HU", "3º"),
    ("Conforto Pediatria", "8367", "HU", "HU", "6º"),
    ("Conforto UTI Neo/Ped", "8585", "HU", "HU", "4º"),
    ("Consultório 1 Multiprofissional", "8469", "HU", "HU", "2º"),
    ("Consultório 2 Nefrologia", "8426", "HU", "HU", "2º"),
    ("Consultório 3 Multiprofissional", "8427", "HU", "HU", "2º"),
    ("Consultório Infantil (PA)", "8594", "PA", "PA", "1º"),
    ("Contabilidade", "8504", "ADM", "ADM", "2º"),
    ("Contas Médicas", "8503", "ADM", "ADM", "2º"),
    ("COREME", "8341", "DIR", "Direção", "1º"),
    ("Dietoterapia", "8333", "HU", "HU", "1º"),
    ("Direção (Recepção)", "8358", "DIR", "Direção", "1º"),
    ("Eletrocardiograma", "8311", "HU", "HU", "2º"),
    ("Endoscopia (Sala 02)", "8348", "HU", "HU", "2º"),
    ("Endoscopia Administrativo", "8348", "HU", "HU", "2º"),
    ("Engenharia Clínica", "8390", "HU", "HU", "2º"),
    ("Escriturários (4º Piso)", "8329", "HU", "HU", "4º"),
    ("Farmácia (4º Piso)", "8332", "HU", "HU", "4º"),
    ("Farmácia de Manipulação", "8343", "HU", "HU", "1º"),
    ("Fátima (Posto de Enfermagem)", "8302", "HU", "HU", "4º"),
    ("Fátima (Supervisão)", "8596", "HU", "HU", "4º"),
    ("Faturamento (Convênios)", "8321", "ADM", "ADM", "2º"),
    ("Faturamento SUS", "8374", "ADM", "ADM", "2º"),
    ("Financeiro (ADM)", "8309", "ADM", "ADM", "2º"),
    ("Financeiro / Caixa (HU)", "8379", "HU", "HU", "2º"),
    ("Fisioterapia", "8593", "HU", "HU", "5º"),
    ("Gerência Administrativa", "8559", "ADM", "ADM", "2º"),
    ("Gerência Assistencial", "8314", "HU", "HU", "3º"),
    ("Gerência de Hotelaria", "8485", "HU", "HU", "2º"),
    ("Gerência de Suprimentos", "8487", "HU", "HU", "1º"),
    ("Gestão de Pessoas", "8344", "ADM", "ADM", "1º"),
    ("Guadalupe (Posto de Enfermagem)", "8552", "HU", "HU", "4º"),
    ("Hemodiálise (Administrativo)", "8364", "HU", "HU", "2º"),
    ("Hemodiálise (Supervisão)", "8562", "HU", "HU", "2º"),
    ("Hemodinâmica", "8558", "HU", "HU", "2º"),
    ("Higienização", "8369", "HU", "HU", "1º"),
    ("Internação (Convênios)", "8371", "HU", "HU", "2º"),
    ("Internação (Supervisão)", "8599", "HU", "HU", "2º"),
    ("Internação SUS", "8391", "HU", "HU", "2º"),
    ("Jurídico", "8349", "ADM", "ADM", "2º"),
    ("Lavanderia", "8362", "HU", "HU", "1º"),
    ("LEAC (Administrativo)", "8359", "HU", "HU", "1º"),
    ("LEAC (Agência Transfusional)", "8330", "HU", "HU", "1º"),
    ("LEAC (Bioquímica)", "8572", "HU", "HU", "1º"),
    ("LEAC (Convênios)", "8573", "HU", "HU", "1º"),
    ("LEAC (Gerência)", "8385", "HU", "HU", "1º"),
    ("LEAC (Hematologia)", "8572", "HU", "HU", "1º"),
    ("LEAC (Microbiologia)", "8572", "HU", "HU", "1º"),
    ("LEAC (Supervisão Administrativa)", "8307", "HU", "HU", "1º"),
    ("LEAC (Triagem)", "8551", "HU", "HU", "1º"),
    ("LEAC SUS", "8573", "HU", "HU", "1º"),
    ("Manutenção", "8339", "MAN", "Manutenção", "1º"),
    ("Maternidade (Convênios)", "8356", "HU", "HU", "3º"),
    ("Maternidade (SUS)", "8322", "HU", "HU", "3º"),
    ("NIR", "8394", "HU", "HU", "1º"),
    ("Nutrição (Produção)", "8351", "HU", "HU", "1º"),
    ("OPME", "8466", "HU", "HU", "6º"),
    ("Ouvidoria", "8375", "HU", "HU", "2º"),
    ("PA Convênios Maternidade", "8508", "HU", "HU", "3º"),
    ("Pediatria (4º Piso)", "8366", "HU", "HU", "4º"),
    ("Portaria", "8319", "HU", "HU", "2º"),
    ("Posto de Enfermagem (Clínica Médica)", "8301", "HU", "HU", "3º"),
    ("Projetos", "8564", "ADM", "ADM", "1º"),
    ("Pronto Atendimento (Recepção)", "8378", "PA", "PA", "1º"),
    ("Pronto Atendimento (Posto de Enfermagem)", "8353", "PA", "PA", "1º"),
    ("Pronto Socorro (Acolhimento)", "8592", "PS", "HU", "1º"),
    ("Pronto Socorro (Cirurgia/Bucomaxilofacial)", "8550", "PS", "HU", "1º"),
    ("Pronto Socorro (SCIH)", "8568", "PS", "HU", "1º"),
    ("Pronto Socorro (Coleta Laboratorial)", "8567", "PS", "HU", "1º"),
    ("Pronto Socorro (Compras / Financeiro)", "8576", "PS", "HU", "1º"),
    ("Pronto Socorro (Consultório Pediátrico)", "8377", "PS", "HU", "1º"),
    ("Pronto Socorro (Direção)", "8574", "PS", "HU", "1º"),
    ("Pronto Socorro (Emergência)", "8304", "PS", "HU", "1º"),
    ("Pronto Socorro (Escriturários)", "8567", "PS", "HU", "1º"),
    ("Pronto Socorro (Estabilização Adulto)", "8592", "PS", "HU", "1º"),
    ("Pronto Socorro (Farmácia)", "8582", "PS", "HU", "1º"),
    ("Pronto Socorro (Portaria)", "8577", "PS", "HU", "1º"),
    ("Pronto Socorro (Posto de Enfermagem)", "8588", "PS", "HU", "1º"),
    ("Pronto Socorro (Posto Enfermagem Pediatria)", "8579", "PS", "HU", "1º"),
    ("Pronto Socorro (Raio-X)", "8303", "PS", "HU", "1º"),
    ("Pronto Socorro (Recepção)", "8305", "PS", "HU", "1º"),
    ("Pronto Socorro (Serviço Social)", "8584", "PS", "HU", "1º"),
    ("Pronto Socorro (Triagem Médica)", "8586", "PS", "HU", "1º"),
    ("Prontuários (Direção)", "8589", "HU", "HU", "2º"),
    ("Psicologia", "8396", "HU", "HU", "5º"),
    ("Qualidade", "8482", "HU", "HU", "3º"),
    ("Radiologia", "8397", "HU", "HU", "2º"),
    ("Raio-X (Laudos)", "8384", "HU", "HU", "2º"),
    ("Recém Nascido", "8325", "HU", "HU", "3º"),
    ("Recepção (5º Piso)", "8486", "HU", "HU", "5º"),
    ("Rede Fisio", "8389", "HU", "HU", "4º"),
    ("RH", "8345", "ADM", "ADM", "1º"),
    ("RT Farmácia", "8335", "HU", "HU", "3º"),
    ("Sala de Diálise 1", "8313", "HU", "HU", "2º"),
    ("Sala de Diálise 3", "8312", "HU", "HU", "2º"),
    ("Sala de Diálise 4", "8328", "HU", "HU", "2º"),
    ("Sala de Estudos (3º Piso)", "8398", "HU", "HU", "3º"),
    ("Sala de Estudos (4º Piso - Conceição)", "8422", "HU", "HU", "4º"),
    ("Sala de Estudos (5º Piso)", "8569", "HU", "HU", "5º"),
    ("Sala de Laudos", "8317", "HU", "HU", "2º"),
    ("São Camilo (Posto de Enfermagem)", "8571", "HU", "HU", "3º"),
    ("Saúde + Fácil (Agendamento de Consultas)", "8383", "PA", "PA", "2º"),
    ("Saúde + Fácil (Agendamento Reumato SUS)", "8583", "PA", "PA", "2º"),
    ("Saúde + Fácil (Recepção)", "8570", "PA", "PA", "1º"),
    ("Saúde + Fácil (Supervisão)", "8488", "PA", "PA", "2º"),
    ("Saúde Ocupacional (Administrativo)", "8561", "PA", "PA", "1º"),
    ("Saúde Ocupacional (Recepção)", "8355", "PA", "PA", "1º"),
    ("SCIH", "8340", "HU", "HU", "3º"),
    ("SESMT (Consultório 1)", "8433", "PA", "PA", "1º"),
    ("SESMT Interno", "8352", "ADM", "ADM", "1º"),
    ("Serviço Social (Guadalupe)", "8388", "HU", "HU", "4º"),
    ("Supervisão de Enfermagem (Materno/Infantil)", "8553", "HU", "HU", "3º"),
    ("Supervisão de Hotelaria (2º Piso)", "8491", "HU", "HU", "2º"),
    ("Supervisão de Restaurante", "8501", "HU", "HU", "1º"),
    ("Supervisão UTIs", "8497", "HU", "HU", "5º"),
    ("Telefonia", "8327", "HU", "HU", "6º"),
    ("TI", "8306", "ADM", "ADM", "1º"),
    ("Tomografia", "8393", "HU", "HU", "2º"),
    ("UTI Geral (Farmácia)", "8363", "HU", "HU", "2º"),
    ("UTI Geral (Posto de Enfermagem / Área 1)", "8334", "HU", "HU", "2º"),
    ("UTI Geral (Médicos)", "8392", "HU", "HU", "2º"),
    ("UTI Neo Natal", "8365", "HU", "HU", "4º"),
    ("UTI Pediátrica", "8590", "HU", "HU", "4º"),
    ("UTSI", "8329", "HU", "HU", "4º"),
    ("UTU Área 2", "8308", "HU", "HU", "2º"),
]

# ── Detalhes de equipamento (PAT / N° Série / MAC / IP) ─────────────────
# Extraídos da planilha VOIP_HUSFP.xlsx, aba "Consolidado", cruzados por
# nome do setor + ramal. Um setor pode ter mais de um aparelho cadastrado
# (linhas discadoras compartilhadas). Setores sem chave aqui não têm
# equipamento registrado nessa planilha.
# ── Filas compartilhadas ─────────────────────────────────────────────────
# (Nome do setor, Ramal) -> número da fila (quando o setor toca em uma fila
# compartilhada, além do ramal direto do aparelho). Vem da planilha
# Ramais_HUSFP_2026.xlsx, coluna "FILA".
FILA_INFO = {
    ("CDI (Sala 01)", "8449"): "8316",
    ("CDI (Sala 02)", "8563"): "8316",
    ("CDI (Sala 03)", "8380"): "8316",
    ("Consultório 1 Multiprofissional", "8469"): "8554",
    ("Consultório 2 Nefrologia", "8426"): "8554",
    ("Consultório 3 Multiprofissional", "8427"): "8554",
}


DEVICE_INFO = {
    ("Almoxarifado (Central de Materiais)", "8560"): [
        {"pat": "9514", "n_serie": "Q3D1245019226", "mac": "0C383E6113E3", "ip": "192.168.11.28"},
        {"pat": "9970", "n_serie": "Q3G1252002277", "mac": "0C383E66EB3B1", "ip": "192.168.10.208"},
    ],
    ("Almoxarifado (Recebimentos)", "8560"): [
        {"pat": "9599", "n_serie": None, "mac": "0C383E6EB308", "ip": "192.168.10.29"},
    ],
    ("Bloco Cirúrgico (Farmácia)", "8350"): [
        {"pat": "9956", "n_serie": "Q3G124A010205", "mac": "0C383E69B59A", "ip": "192.168.51.11"},
    ],
    ("Bloco Cirúrgico (Marcação)", "8376"): [
        {"pat": None, "n_serie": None, "mac": "0C383E", "ip": "192.168.51.65"},
    ],
    ("Bloco Cirúrgico (Posto de Enfermagem)", "8591"): [
        {"pat": "9597", "n_serie": "Q3G124A015023", "mac": "0C383E69E3C4", "ip": "192.168.51.65"},
    ],
    ("Bloco Cirúrgico (Sala de Recuperação)", "8347"): [
        {"pat": "9971", "n_serie": "Q3G1252002279", "mac": "0C383E6EB3B3", "ip": "192.168.51.66"},
    ],
    ("CAF", "8331"): [
        {"pat": "9516", "n_serie": "Q3D124519237", "mac": "0C383E6113EE", "ip": "192.168.11.31"},
    ],
    ("Casa da Gestante", "8360"): [
        {"pat": "9439", "n_serie": None, "mac": "0C383E611026", "ip": "192.168.70.249"},
    ],
    ("CDI - Administrativo", "8496"): [
        {"pat": None, "n_serie": None, "mac": None, "ip": None},
    ],
    ("CDI (Convênios)", "8310"): [
        {"pat": "9958", "n_serie": "Q3G124A010195", "mac": "0C383E69B590", "ip": "192.168.21.69"},
    ],
    ("CDI (SUS)", "8318"): [
        {"pat": None, "n_serie": None, "mac": "0C383E6EB337", "ip": None},
    ],
    ("Central de Cadastro - MV", "8507"): [
        {"pat": "9021", "n_serie": "Q3D1249002104", "mac": "0C383E661810", "ip": None},
    ],
    ("Central de Cadastro SIGH", "8506"): [
        {"pat": "9987", "n_serie": "Q3G1252002259", "mac": "0C383E6EB39F", "ip": None},
        {"pat": "9558", "n_serie": None, "mac": "0C383E6EB89B", "ip": "192.168.51.9"},
    ],
    ("Centro Acadêmico (Recepção)", "8386"): [
        {"pat": "9576", "n_serie": "Q3G1252002283", "mac": "0C383E6EB387", "ip": "192.168.51.56"},
    ],
    ("Clínica Cirúrgica (Escriturários)", "8557"): [
        {"pat": "9406", "n_serie": "Q3G1252003534", "mac": "0C383E6EB89A", "ip": "192.168.41.14"},
    ],
    ("Clínica Cirúrgica (Posto de Enfermagem)", "8320"): [
        {"pat": "9426", "n_serie": "Q3G1252004128", "mac": "0C383E6EBAEC", "ip": "192.168.41.36"},
    ],
    ("CME", "8336"): [
        {"pat": "9504", "n_serie": "Q3D1245019235", "mac": "0C383E6113EC", "ip": "192.168.11.25"},
        {"pat": "9511", "n_serie": "Q3D1245019245", "mac": "0C383E6113F6", "ip": "192.168.11.27"},
    ],
    ("Comunicação e Marketing", "8368"): [
        {"pat": None, "n_serie": "04QJ2900199KX", "mac": "4851CF53F066", "ip": None},
    ],
    ("Compras", "8326"): [
        {"pat": "9528", "n_serie": "0C383E6113D8", "mac": "0C383E6113D8", "ip": "192.168.11.26"},
        {"pat": "9544", "n_serie": "Q3D1245019241", "mac": "0C383E6113F2", "ip": "192.168.11.24"},
        {"pat": "9600", "n_serie": "Q3G1252002280", "mac": "0C383E6EB3B4", "ip": None},
        {"pat": "9955", "n_serie": "Q3G124A013136", "mac": "0C383E69B592", "ip": "192.168.80.28"},
        {"pat": "9573", "n_serie": "Q3G1252002157", "mac": "0C383E6EB339", "ip": "192.168.51.55"},
    ],
    ("Conceição (Posto de Enfermagem)", "8575"): [
        {"pat": None, "n_serie": None, "mac": None, "ip": None},
    ],
    ("Conceição (Sala de Prescrição)", "8324"): [
        {"pat": "9427", "n_serie": None, "mac": "0C383E61102D", "ip": None},
    ],
    ("Conforto Médico (Clínica Médica)", "8357"): [
        {"pat": "9441", "n_serie": None, "mac": "0C383E61102A", "ip": None},
    ],
    ("Conforto Médico (Ginecologia)", "8361"): [
        {"pat": "9436", "n_serie": "Q3G1252004122", "mac": "0C383E6EBAE6", "ip": "192.168.41.13"},
    ],
    ("Conforto Pediatria", "8367"): [
        {"pat": "9591", "n_serie": None, "mac": "0C383E69D4F6", "ip": "192.168.21.29"},
    ],
    ("Conforto UTI Neo/Ped", "8585"): [
        {"pat": "9590", "n_serie": None, "mac": "0C383E69D4FC", "ip": "192.168.21.36"},
    ],
    ("Consultório Infantil (PA)", "8594"): [
        {"pat": "9512", "n_serie": "Q3D1245019229", "mac": "0C383E6113E6", "ip": "192.168.11.32"},
    ],
    ("Contabilidade", "8504"): [
        {"pat": "9966", "n_serie": "Q3G124A013136", "mac": "0C383E69DC65", "ip": "192.168.21.44"},
    ],
    ("Contas Médicas", "8503"): [
        {"pat": "9566", "n_serie": None, "mac": "0C383E69D4F7", "ip": "192.168.21.34"},
    ],
    ("COREME", "8341"): [
        {"pat": "9984", "n_serie": None, "mac": "0C383EE6111F7", "ip": "192.168.21.32"},
    ],
    ("Dietoterapia", "8333"): [
        {"pat": "10000", "n_serie": None, "mac": "0C383E6111FC", "ip": "192.168.21.31"},
    ],
    ("Direção (Recepção)", "8358"): [
        {"pat": "9963", "n_serie": "Q3D1245018736", "mac": "0C383E6111F9", "ip": "192.168.21.38"},
    ],
    ("Eletrocardiograma", "8311"): [
        {"pat": "9983", "n_serie": "Q3D1245018740", "mac": "0C383E6111F6", "ip": "192.168.21.39"},
    ],
    ("Endoscopia (Sala 02)", "8348"): [
        {"pat": "9994", "n_serie": None, "mac": "0C383E6111FF", "ip": "192.168.41.9"},
    ],
    ("Endoscopia Administrativo", "8348"): [
        {"pat": "9998", "n_serie": "Q3G124A010206", "mac": "0C383E69B59B", "ip": "192.168.41.8"},
    ],
    ("Engenharia Clínica", "8390"): [
        {"pat": "9506", "n_serie": "Q3D1245019227", "mac": "0C383E6113E4", "ip": "192.168.11.30"},
        {"pat": "9598", "n_serie": "Q3G124A010202", "mac": "0C383E69B597", "ip": "192.168.41.12"},
    ],
    ("Escriturários (4º Piso)", "8329"): [
        {"pat": "9572", "n_serie": "Q3G1252002273", "mac": "0C383E6EB3AD", "ip": None},
    ],
    ("Farmácia (4º Piso)", "8332"): [
        {"pat": "8277", "n_serie": None, "mac": "0C383E69B11B", "ip": None},
    ],
    ("Farmácia de Manipulação", "8343"): [
        {"pat": "8283", "n_serie": None, "mac": "0C383E69B11D", "ip": None},
    ],
    ("Fátima (Posto de Enfermagem)", "8302"): [
        {"pat": "8295", "n_serie": None, "mac": "0C383E69B11C", "ip": None},
    ],
    ("Fátima (Supervisão)", "8596"): [
        {"pat": "8285", "n_serie": None, "mac": "0C383E69B11E", "ip": None},
    ],
    ("Faturamento (Convênios)", "8321"): [
        {"pat": "9802", "n_serie": None, "mac": "0C383E69811F", "ip": None},
        {"pat": "8293", "n_serie": None, "mac": "0C383E69B121", "ip": None},
        {"pat": "8294", "n_serie": None, "mac": "0C383E69B122", "ip": None},
        {"pat": "8284", "n_serie": None, "mac": "0C383E69B120", "ip": None},
        {"pat": "9041", "n_serie": None, "mac": "0C383E69B123", "ip": None},
    ],
    ("Faturamento SUS", "8374"): [
        {"pat": "9804", "n_serie": None, "mac": "0C383E69B113", "ip": None},
        {"pat": "9965", "n_serie": None, "mac": "0C383E6114B0", "ip": None},
    ],
    ("Financeiro (ADM)", "8309"): [
        {"pat": "9985", "n_serie": None, "mac": "0C383E6114AC", "ip": None},
        {"pat": "9508", "n_serie": "Q3G1252002565", "mac": "0C383E6EB4D1", "ip": None},
        {"pat": "9425", "n_serie": None, "mac": "0C383E6EBB8A", "ip": "192.168.51.8"},
        {"pat": "9438", "n_serie": None, "mac": "0C383E611028", "ip": None},
        {"pat": "9518", "n_serie": None, "mac": "0C383E69DF84", "ip": "192.168.21.22"},
    ],
    ("Financeiro / Caixa (HU)", "8379"): [
        {"pat": "9565", "n_serie": None, "mac": "0C383E69D4F3", "ip": "Não quiseram"},
    ],
    ("Fisioterapia", "8593"): [
        {"pat": "9550", "n_serie": "Q3G124A007301", "mac": "0C383E69AA42", "ip": "192.168..21.77"},
    ],
    ("Gerência Administrativa", "8559"): [
        {"pat": "9517", "n_serie": "Q3G124A007295", "mac": "0C383E69AA3C", "ip": "192.168.10.214"},
    ],
    ("Gerência Assistencial", "8314"): [
        {"pat": "9600", "n_serie": "Q3G1252002280", "mac": "0C383E6EB3B4", "ip": None},
    ],
    ("Gerência de Hotelaria", "8485"): [
        {"pat": None, "n_serie": None, "mac": None, "ip": None},
    ],
    ("Gerência de Suprimentos", "8487"): [
        {"pat": "9743", "n_serie": "Q3G1252002255", "mac": "0C383E6EB39B", "ip": "192.168.41.25"},
    ],
    ("Gestão de Pessoas", "8344"): [
        {"pat": "9540", "n_serie": "Q3G124A013132", "mac": "0C383E69DC61", "ip": "192.168.41.16"},
        {"pat": None, "n_serie": None, "mac": None, "ip": None},
        {"pat": "9820", "n_serie": None, "mac": "0C383E69DF87", "ip": "192.168.21.19"},
    ],
    ("Guadalupe (Posto de Enfermagem)", "8552"): [
        {"pat": "9837", "n_serie": None, "mac": "0C383E69DF7D", "ip": "192.168.21.25"},
    ],
    ("Hemodiálise (Administrativo)", "8364"): [
        {"pat": "9962", "n_serie": "Q3G124A010194", "mac": "0C383E69B58F", "ip": "192.168.21.70"},
    ],
    ("Hemodiálise (Supervisão)", "8562"): [
        {"pat": "9835", "n_serie": None, "mac": "0C383E69DF46", "ip": "192.168.21.20"},
    ],
    ("Hemodinâmica", "8558"): [
        {"pat": "9954", "n_serie": "Q3G124A010203", "mac": "0C383E69B598", "ip": "192.168.11.35"},
    ],
    ("Higienização", "8369"): [
        {"pat": "9981", "n_serie": None, "mac": "0C383E6114A7", "ip": None},
    ],
    ("Internação (Convênios)", "8371"): [
        {"pat": "9416", "n_serie": "Q3G124A011330", "mac": "0C383E69D557", "ip": "192.168.11.14"},
    ],
    ("Internação (Supervisão)", "8599"): [
        {"pat": "9543", "n_serie": "Q3D1245019221", "mac": "0C383E6113DE", "ip": None},
    ],
    ("Internação SUS", "8391"): [
        {"pat": "9594", "n_serie": "Q3G124A011336", "mac": "0C383E69D55D", "ip": "192.168.11.16"},
    ],
    ("Jurídico", "8349"): [
        {"pat": "9595", "n_serie": "Q3G124A011339", "mac": "0C383E69D560", "ip": "192.168.11.23"},
    ],
    ("Lavanderia", "8362"): [
        {"pat": "9593", "n_serie": "Q3G124A011341", "mac": "0C383E69D562", "ip": "192.168.11.20"},
        {"pat": None, "n_serie": "04QJ2900922TZ", "mac": "4851CF53F338", "ip": "192.168.11.5"},
    ],
    ("LEAC (Administrativo)", "8359"): [
        {"pat": "9415", "n_serie": "Q3G124S011333", "mac": "0C383E69D55A", "ip": "192.168.11.18"},
    ],
    ("LEAC (Agência Transfusional)", "8330"): [
        {"pat": "9419", "n_serie": "Q3G124A011335", "mac": "0C383E69D55C", "ip": "192.168.11.21"},
    ],
    ("LEAC (Bioquímica)", "8572"): [
        {"pat": "9096", "n_serie": "Q3G124A011337", "mac": "0C383E69D55E", "ip": "192.168.11.22"},
    ],
    ("LEAC (Convênios)", "8573"): [
        {"pat": "9095", "n_serie": "Q3G124A011338", "mac": "0C383E69D55F", "ip": "192.168.11.17"},
    ],
    ("LEAC (Gerência)", "8385"): [
        {"pat": "9417", "n_serie": "Q3G124A011332", "mac": "0C383E69D569", "ip": "192.168.11.19"},
    ],
    ("LEAC (Hematologia)", "8572"): [
        {"pat": None, "n_serie": "04QH4600472QN", "mac": "24FD0D54E424", "ip": "192.168.11.6"},
    ],
    ("LEAC (Microbiologia)", "8572"): [
        {"pat": "9524", "n_serie": "Q3G124A007304", "mac": "0C383E69AA45", "ip": None},
    ],
    ("LEAC (Supervisão Administrativa)", "8307"): [
        {"pat": "9541", "n_serie": "Q3G124A007299", "mac": "0C383E69AA40", "ip": None},
    ],
    ("LEAC (Triagem)", "8551"): [
        {"pat": None, "n_serie": "04QH4603448BP", "mac": "24FD0D54D824", "ip": "192.168.11.2"},
    ],
    ("LEAC SUS", "8573"): [
        {"pat": None, "n_serie": "04QJ2901137NU", "mac": "4851CF53F40F", "ip": "192.168.31.2"},
    ],
    ("Manutenção", "8339"): [
        {"pat": None, "n_serie": "04QH4604633FU", "mac": "24FD0D54E5C5", "ip": "192.168.31.7"},
        {"pat": "9507", "n_serie": "Q3G1252002538", "mac": "0C383E6EB4B6", "ip": "192.168.11.58"},
        {"pat": "6719", "n_serie": "04QJ2900189Q1", "mac": "4851CF53F05B", "ip": "192.168.11.10"},
    ],
    ("Maternidade (Convênios)", "8356"): [
        {"pat": "9546", "n_serie": "Q3G124A013130", "mac": "0C383E69DC5F", "ip": "192.168.31.15"},
    ],
    ("Maternidade (SUS)", "8322"): [
        {"pat": "9560", "n_serie": "Q3G1252003532", "mac": "0C383E", "ip": None},
    ],
    ("NIR", "8394"): [
        {"pat": "9565", "n_serie": "Q3G124A001230", "mac": None, "ip": None},
    ],
    ("Nutrição (Produção)", "8351"): [
        {"pat": "9997", "n_serie": "Q3G124A014998", "mac": "0C383E69E3AB", "ip": None},
    ],
    ("OPME", "8466"): [
        {"pat": "9529", "n_serie": "Q3G124A013131", "mac": "0C383E69DC60", "ip": "192.168.31.13"},
    ],
    ("Ouvidoria", "8375"): [
        {"pat": "9531", "n_serie": "Q3G124A013134", "mac": "0C383E69DC63", "ip": "192.168.31.14"},
    ],
    ("Pediatria (4º Piso)", "8366"): [
        {"pat": None, "n_serie": None, "mac": None, "ip": None},
    ],
    ("Portaria", "8319"): [
        {"pat": "9575", "n_serie": None, "mac": "0C383E6114AF", "ip": None},
    ],
    ("Posto de Enfermagem (Clínica Médica)", "8301"): [
        {"pat": "9548", "n_serie": None, "mac": "0C383E69DC62", "ip": "192.168.81.248"},
        {"pat": "9501", "n_serie": "Q3G124A013135", "mac": "0C383E69DC64", "ip": "192.168.81.16"},
    ],
    ("Projetos", "8564"): [
        {"pat": "9509", "n_serie": "Q3G1252002528", "mac": "0C383E6EB4AC", "ip": None},
        {"pat": None, "n_serie": None, "mac": None, "ip": None},
    ],
    ("Pronto Atendimento (Recepção)", "8378"): [
        {"pat": "9505", "n_serie": "Q3G1252002555", "mac": "0C383E6EB4C7", "ip": None},
    ],
    ("Pronto Atendimento (Posto de Enfermagem)", "8353"): [
        {"pat": "9538", "n_serie": "Q3G1252002559", "mac": "0C383E6EB4CB", "ip": None},
    ],
    ("Pronto Socorro (Acolhimento)", "8592"): [
        {"pat": "9819", "n_serie": None, "mac": "0C383E661805", "ip": "192.168.11.83"},
    ],
    ("Pronto Socorro (Cirurgia/Bucomaxilofacial)", "8550"): [
        {"pat": "9520", "n_serie": None, "mac": "0C383E6EB6C9", "ip": "192.168.11.80"},
    ],
    ("Pronto Socorro (SCIH)", "8568"): [
        {"pat": "9532", "n_serie": None, "mac": "0C383E6EB62E", "ip": "192.168.12.71"},
    ],
    ("Pronto Socorro (Coleta Laboratorial)", "8567"): [
        {"pat": "9534", "n_serie": None, "mac": "0C383E6EB6CA", "ip": "192.168.11.79"},
    ],
    ("Pronto Socorro (Compras / Financeiro)", "8576"): [
        {"pat": "9522", "n_serie": None, "mac": "0C383E6EB5DD", "ip": "192.168.11.76"},
    ],
    ("Pronto Socorro (Consultório Pediátrico)", "8377"): [
        {"pat": "9523", "n_serie": None, "mac": "0C383E6EB6D9", "ip": "192.168.11.82"},
    ],
    ("Pronto Socorro (Direção)", "8574"): [
        {"pat": "9525", "n_serie": None, "mac": "0C383E6EB6D3", "ip": "192.168.11.73"},
    ],
    ("Pronto Socorro (Emergência)", "8304"): [
        {"pat": "8286", "n_serie": "Q3G124A013936", "mac": "0C383E69DF85", "ip": "192.168.21.28"},
    ],
    ("Pronto Socorro (Escriturários)", "8567"): [
        {"pat": "9526", "n_serie": None, "mac": "0C383E6EB5E3", "ip": "192.168.11.78"},
    ],
    ("Pronto Socorro (Estabilização Adulto)", "8592"): [
        {"pat": "9819", "n_serie": None, "mac": "0C383E661805", "ip": "192.168.11.83"},
    ],
    ("Pronto Socorro (Farmácia)", "8582"): [
        {"pat": "9022", "n_serie": None, "mac": "0C383E66180F", "ip": "192.168.11.75"},
    ],
    ("Pronto Socorro (Portaria)", "8577"): [
        {"pat": "9530", "n_serie": None, "mac": "0C383E6EB6CB", "ip": "192.168.11.84"},
    ],
    ("Pronto Socorro (Posto de Enfermagem)", "8588"): [
        {"pat": "9836", "n_serie": None, "mac": "0C383E66180A", "ip": "192.168.11.74"},
    ],
    ("Pronto Socorro (Posto Enfermagem Pediatria)", "8579"): [
        {"pat": "9039", "n_serie": None, "mac": "0C383E66180C", "ip": "192.168.11.70"},
    ],
    ("Pronto Socorro (Raio-X)", "8303"): [
        {"pat": "9533", "n_serie": None, "mac": "0C383E6EB6D7", "ip": "192.168.11.81"},
    ],
    ("Pronto Socorro (Recepção)", "8305"): [
        {"pat": "9535", "n_serie": None, "mac": "0C383E6EB6D4", "ip": "192.168.11.77"},
    ],
    ("Pronto Socorro (Serviço Social)", "8584"): [
        {"pat": "9474", "n_serie": "Q3G1252004119", "mac": "0C383E6EBAE3", "ip": "192.168.51.10"},
        {"pat": "9509", "n_serie": None, "mac": "0C383E6EB4AC", "ip": "192.168.11.57"},
    ],
    ("Pronto Socorro (Triagem Médica)", "8586"): [
        {"pat": "9539", "n_serie": "Q3G1252002557", "mac": "0C383E6EB4C9", "ip": "192.168.11.55"},
        {"pat": "9542", "n_serie": "Q3G124A007308", "mac": "0C383E69AA49", "ip": None},
    ],
    ("Prontuários (Direção)", "8589"): [
        {"pat": "9834", "n_serie": "Q3G124A013937", "mac": "0C383E69DF86", "ip": "192.168.21.26"},
    ],
    ("Psicologia", "8396"): [
        {"pat": "9424", "n_serie": "Q3D1245018264", "mac": "0C383E611021", "ip": None},
    ],
    ("Qualidade", "8482"): [
        {"pat": "9536", "n_serie": "Q3G124A007303", "mac": "0C383E69AA44", "ip": "Falta"},
    ],
    ("Radiologia", "8397"): [
        {"pat": "9552", "n_serie": "Q3G1252004156", "mac": "0C383E6EBB08", "ip": "192.168.41.15"},
    ],
    ("Raio-X (Laudos)", "8384"): [
        {"pat": "9968", "n_serie": None, "mac": "0C383E6114A1", "ip": None},
    ],
    ("Recém Nascido", "8325"): [
        {"pat": "9574", "n_serie": None, "mac": "0C383E6114AB", "ip": None},
    ],
    ("Recepção (5º Piso)", "8486"): [
        {"pat": "9999", "n_serie": "Q3D1245018735", "mac": "0C383E6111F8", "ip": "192.168.31.11"},
    ],
    ("Rede Fisio", "8389"): [
        {"pat": "9513", "n_serie": "Q3D1245019238", "mac": "0C383E6113EF", "ip": "192.168.21.18"},
    ],
    ("RH", "8345"): [
        {"pat": "8209", "n_serie": None, "mac": "0C383E69DF78", "ip": "192.168.21.35"},
        {"pat": "8281", "n_serie": "Q3G124A013922", "mac": "0C383E69DF77", "ip": "192.168.21.21"},
    ],
    ("RT Farmácia", "8335"): [
        {"pat": "9964", "n_serie": "Q3D1245018743", "mac": "0C383E611200", "ip": "192.168.31.12"},
    ],
    ("Sala de Diálise 1", "8313"): [
        {"pat": "9553", "n_serie": "Q3G1252003537", "mac": "0C383E6EB89D", "ip": "192.168.51.7"},
    ],
    ("Sala de Diálise 3", "8312"): [
        {"pat": "8274", "n_serie": "Q3G124A013870", "mac": "0C383E69DF43", "ip": "192.168.21.24"},
    ],
    ("Sala de Diálise 4", "8328"): [
        {"pat": "9519", "n_serie": "Q3G124A013924", "mac": "0C383E69DF79", "ip": "192.168.21.33"},
    ],
    ("Sala de Estudos (3º Piso)", "8398"): [
        {"pat": "9537", "n_serie": "Q3G1252002556", "mac": "OC383E6EB4C8", "ip": None},
    ],
    ("Sala de Estudos (4º Piso - Conceição)", "8422"): [
        {"pat": "9822", "n_serie": None, "mac": "0C383E66180B", "ip": "192.168.51.105"},
    ],
    ("Sala de Estudos (5º Piso)", "8569"): [
        {"pat": "9986", "n_serie": "Q3G124A015028", "mac": "0C383E69E3C9", "ip": None},
    ],
    ("Sala de Laudos", "8317"): [
        {"pat": "9982", "n_serie": "Q3G124A015030", "mac": "0C383E69E3CB", "ip": None},
        {"pat": "9545", "n_serie": "Q3G1252002552", "mac": "0C383E6EB4C4", "ip": "192.168.80.21"},
    ],
    ("São Camilo (Posto de Enfermagem)", "8571"): [
        {"pat": "9502", "n_serie": "Q3G124A013137", "mac": "0C383E69DC69", "ip": None},
    ],
    ("Saúde + Fácil (Agendamento de Consultas)", "8383"): [
        {"pat": "9547", "n_serie": "Q3G124A007300", "mac": "0C383E69AA41", "ip": None},
    ],
    ("Saúde + Fácil (Agendamento Reumato SUS)", "8583"): [
        {"pat": "9975", "n_serie": "Q3G124A013138", "mac": "0C383E69DC67", "ip": None},
    ],
    ("Saúde + Fácil (Recepção)", "8570"): [
        {"pat": "9992", "n_serie": "Q3G124A013139", "mac": "0C383E69DC68", "ip": None},
    ],
    ("Saúde + Fácil (Supervisão)", "8488"): [
        {"pat": "9437", "n_serie": "Q3D1245018268", "mac": "0C383E611025", "ip": None},
    ],
    ("Saúde Ocupacional (Administrativo)", "8561"): [
        {"pat": None, "n_serie": "04QH4603445PD", "mac": "24FD0D54D821", "ip": "192.168.81.251"},
        {"pat": "9953", "n_serie": "Q3G124A010208", "mac": "0C383E69B59D", "ip": "192.168.31.9"},
        {"pat": "9957", "n_serie": "Q3G124A010204", "mac": "0C383E69B599", "ip": "192.168.31.10"},
        {"pat": None, "n_serie": None, "mac": "0C383E", "ip": None},
    ],
    ("Saúde Ocupacional (Recepção)", "8355"): [
        {"pat": "9977", "n_serie": None, "mac": "0C383E6114AE", "ip": None},
    ],
    ("SCIH", "8340"): [
        {"pat": "9471", "n_serie": None, "mac": "0C383E6114AA", "ip": None},
        {"pat": "9974", "n_serie": None, "mac": "0C383E6114AB", "ip": None},
    ],
    ("SESMT Interno", "8352"): [
        {"pat": "9515", "n_serie": "Q3G1252002036", "mac": "0C383E6EB2C0", "ip": "192.168.31.22"},
        {"pat": "9440", "n_serie": "Q3D1245018270", "mac": "0C383E611027", "ip": "192.168.11.65"},
        {"pat": "9959", "n_serie": "Q3G124A010199", "mac": "0C383E69B594", "ip": "192.168.30.23"},
    ],
    ("Serviço Social (Guadalupe)", "8388"): [
        {"pat": "9421", "n_serie": "Q3G124A011334", "mac": "0C383E69D55B", "ip": "192.168.11.15"},
    ],
    ("Supervisão de Enfermagem (Materno/Infantil)", "8553"): [
        {"pat": "9503", "n_serie": "Q3G1252002554", "mac": "0C383E6EB4C6", "ip": None},
    ],
    ("Supervisão de Hotelaria (3º Piso)", "8491"): [
        {"pat": None, "n_serie": "04QJ2900926VF", "mac": "4851CF53F33C", "ip": "192.168.31.7"},
    ],
    ("Supervisão de Restaurante", "8501"): [
        {"pat": None, "n_serie": None, "mac": "0C383E", "ip": None},
    ],
    ("Supervisão UTIs", "8497"): [
        {"pat": "9442", "n_serie": "Q3D1245018274", "mac": "0C383E61102B", "ip": None},
    ],
    ("Telefonia", "8327"): [
        {"pat": "9978", "n_serie": None, "mac": "0C383E693EC5", "ip": None},
    ],
    ("TI", "8306"): [
        {"pat": "9563", "n_serie": "0C383E69D4EE", "mac": "0C383E69D4EE", "ip": None},
        {"pat": "9967", "n_serie": "Q3D1245018740", "mac": "0C383E6111FD", "ip": "192.168.20.34"},
        {"pat": "9976", "n_serie": "Q3G124A015025", "mac": "0C383E69E3C6", "ip": None},
        {"pat": "9980", "n_serie": None, "mac": "0C383E69E3C7", "ip": None},
    ],
    ("Tomografia", "8393"): [
        {"pat": "?", "n_serie": "?", "mac": "?", "ip": None},
    ],
    ("UTI Geral (Farmácia)", "8363"): [
        {"pat": "9972", "n_serie": "Q3G124A015027", "mac": "0C383E69E3C8", "ip": None},
    ],
    ("UTI Geral (Posto de Enfermagem / Área 1)", "8334"): [
        {"pat": "9978", "n_serie": "Q3G124A015024", "mac": "0C383E69E365", "ip": None},
    ],
    ("UTI Geral (Médicos)", "8392"): [
        {"pat": "9510", "n_serie": "Q3G124A007310", "mac": "0C383E69AA4B", "ip": None},
    ],
    ("UTI Neo Natal", "8365"): [
        {"pat": "9527", "n_serie": "Q3G124A007306", "mac": "0C383E69AA47", "ip": None},
    ],
    ("UTI Pediátrica", "8590"): [
        {"pat": "9549", "n_serie": "Q3G124A007305", "mac": "0C383E69AA46", "ip": None},
    ],
    ("UTSI", "8329"): [
        {"pat": None, "n_serie": None, "mac": "0C383E", "ip": None},
    ],
    ("UTU Área 2", "8308"): [
        {"pat": "9557", "n_serie": "Q3G1252003475", "mac": "0C383E6EB85F", "ip": None},
    ],
}

GENERAL_PHONE = "(53) 2128-8300"
MAX_RESULTS = 20
TELEGRAM_MSG_LIMIT = 4000  # margem de segurança abaixo do limite real (4096)

# ── Administradores (podem editar PAT / MAC / IP) ───────────────────────
# IDs numéricos do Telegram (obtidos com @userinfobot). Se a variável de
# ambiente TELEGRAM_ADMIN_IDS existir (formato "111,222"), ela tem
# prioridade sobre a lista fixa abaixo.
ADMIN_IDS_PADRAO: set = {223069500, 5688363658, 8951485438}  # Leandro, novo admin (via /meuid)
_env_admins = os.environ.get("TELEGRAM_ADMIN_IDS", "")
ADMIN_IDS = {
    int(x) for x in _env_admins.split(",") if x.strip().isdigit()
} or ADMIN_IDS_PADRAO


def is_admin(user_id: int | None) -> bool:
    return user_id is not None and user_id in ADMIN_IDS




# ── Edições de equipamento (PAT / MAC / IP) feitas por admins ───────────
# Persistidas em disco para sobreviver a reinícios do bot. Cada chave é
# "nome_do_setor|ramal|indice_do_aparelho" e o valor é um dict só com os
# campos que foram sobrescritos (ex.: {"ip": "192.168.1.10"}).
DEVICE_OVERRIDES_FILE = "device_overrides.json"
DEVICE_OVERRIDES: dict = {}


def _override_key(name: str, ramal: str, dev_i: int) -> str:
    return f"{name}|{ramal}|{dev_i}"


def load_device_overrides() -> None:
    global DEVICE_OVERRIDES
    try:
        with open(DEVICE_OVERRIDES_FILE, "r", encoding="utf-8") as f:
            DEVICE_OVERRIDES = json.load(f)
    except FileNotFoundError:
        DEVICE_OVERRIDES = {}
    except Exception:
        logger.exception("Falha ao carregar %s", DEVICE_OVERRIDES_FILE)
        DEVICE_OVERRIDES = {}


def save_device_overrides() -> None:
    try:
        with open(DEVICE_OVERRIDES_FILE, "w", encoding="utf-8") as f:
            json.dump(DEVICE_OVERRIDES, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Falha ao salvar %s", DEVICE_OVERRIDES_FILE)


# ── Edição do nome do setor feita por admins ─────────────────────────────
# Persistida em disco. Chave é o índice do ramal em DATA (fixo — não muda
# entre reinícios), valor é o nome atual escolhido pelo admin. Os campos
# internos (busca por MAC/IP cadastrado, fila, overrides de aparelho etc.)
# continuam usando o nome ORIGINAL de DATA como identidade — só a exibição
# muda.
NOME_OVERRIDES_FILE = "nome_overrides.json"
NOME_OVERRIDES: dict = {}  # idx (int) -> novo nome


def load_nome_overrides() -> None:
    global NOME_OVERRIDES
    try:
        with open(NOME_OVERRIDES_FILE, "r", encoding="utf-8") as f:
            bruto = json.load(f)
        NOME_OVERRIDES = {int(k): v for k, v in bruto.items()}
    except FileNotFoundError:
        NOME_OVERRIDES = {}
    except Exception:
        logger.exception("Falha ao carregar %s", NOME_OVERRIDES_FILE)
        NOME_OVERRIDES = {}


def save_nome_overrides() -> None:
    try:
        with open(NOME_OVERRIDES_FILE, "w", encoding="utf-8") as f:
            json.dump({str(k): v for k, v in NOME_OVERRIDES.items()}, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Falha ao salvar %s", NOME_OVERRIDES_FILE)


def nome_exibido(idx: int) -> str:
    """Retorna o nome atual do setor (considerando edição de admin), ou o
    nome original de DATA se nunca foi editado."""
    return NOME_OVERRIDES.get(idx, DATA[idx][0])


# ── Edição do número do ramal feita por admins ───────────────────────────
# Mesmo esquema do NOME_OVERRIDES acima: persistida em disco, chave é o
# índice do ramal em DATA, valor é o número atualmente exibido/buscável.
# Os campos internos que identificam o aparelho fisicamente (DEVICE_INFO,
# DEVICE_OVERRIDES, FILA_INFO — todos chaveados por nome+ramal ORIGINAL)
# continuam usando o ramal original de DATA como identidade; só a exibição
# e a busca passam a usar o número novo.
RAMAL_OVERRIDES_FILE = "ramal_overrides.json"
RAMAL_OVERRIDES: dict = {}  # idx (int) -> novo número de ramal


def load_ramal_overrides() -> None:
    global RAMAL_OVERRIDES
    try:
        with open(RAMAL_OVERRIDES_FILE, "r", encoding="utf-8") as f:
            bruto = json.load(f)
        RAMAL_OVERRIDES = {int(k): v for k, v in bruto.items()}
    except FileNotFoundError:
        RAMAL_OVERRIDES = {}
    except Exception:
        logger.exception("Falha ao carregar %s", RAMAL_OVERRIDES_FILE)
        RAMAL_OVERRIDES = {}


def save_ramal_overrides() -> None:
    try:
        with open(RAMAL_OVERRIDES_FILE, "w", encoding="utf-8") as f:
            json.dump({str(k): v for k, v in RAMAL_OVERRIDES.items()}, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Falha ao salvar %s", RAMAL_OVERRIDES_FILE)


def ramal_exibido(idx: int) -> str:
    """Retorna o número de ramal atual (considerando edição de admin), ou o
    número original de DATA se nunca foi editado."""
    return RAMAL_OVERRIDES.get(idx, DATA[idx][1])


# ── Exclusão de ramal feita por admin (soft delete, com restaurar) ──────
# O setor continua existindo em DATA no índice original (ninguém precisa
# ser reindexado), só fica marcado como excluído: some de buscas,
# listagens, exportação, checklist e do cálculo de "Fila compartilhada" —
# mas continua acessível pela tela dedicada "🗑️ Ramais excluídos" pra
# poder restaurar, caso tenha sido engano.
RAMAIS_EXCLUIDOS_FILE = "ramais_excluidos.json"
RAMAIS_EXCLUIDOS: set = set()  # conjunto de idx (int) excluídos


def load_ramais_excluidos() -> None:
    global RAMAIS_EXCLUIDOS
    try:
        with open(RAMAIS_EXCLUIDOS_FILE, "r", encoding="utf-8") as f:
            bruto = json.load(f)
        RAMAIS_EXCLUIDOS = {int(x) for x in bruto}
    except FileNotFoundError:
        RAMAIS_EXCLUIDOS = set()
    except Exception:
        logger.exception("Falha ao carregar %s", RAMAIS_EXCLUIDOS_FILE)
        RAMAIS_EXCLUIDOS = set()


def save_ramais_excluidos() -> None:
    try:
        with open(RAMAIS_EXCLUIDOS_FILE, "w", encoding="utf-8") as f:
            json.dump(sorted(RAMAIS_EXCLUIDOS), f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Falha ao salvar %s", RAMAIS_EXCLUIDOS_FILE)


def ramal_excluido(idx: int) -> bool:
    return idx in RAMAIS_EXCLUIDOS


def indices_ativos():
    """Índices de DATA que não estão marcados como excluídos — usado em
    toda busca/listagem pra ignorar ramais excluídos automaticamente."""
    return [i for i in range(len(DATA)) if i not in RAMAIS_EXCLUIDOS]


# ── Ramais criados por admins pelo próprio bot ───────────────────────────
# São acrescentados em DATA (em memória) e persistidos em disco pra
# sobreviverem a reinícios do bot.
RAMAIS_CRIADOS_FILE = "ramais_criados.json"


def _recalcular_ramais_compartilhados() -> None:
    global _RAMAIS_COMPARTILHADOS
    _RAMAIS_COMPARTILHADOS = {
        ramal for ramal, qtd in Counter(ramal_exibido(i) for i in indices_ativos()).items() if qtd > 1
    }


def load_ramais_criados() -> None:
    """Lê os ramais criados pelo bot e os acrescenta em DATA. Deve ser
    chamada uma vez, na inicialização, antes do bot começar a atender."""
    try:
        with open(RAMAIS_CRIADOS_FILE, "r", encoding="utf-8") as f:
            itens = json.load(f)
    except FileNotFoundError:
        itens = []
    except Exception:
        logger.exception("Falha ao carregar %s", RAMAIS_CRIADOS_FILE)
        itens = []

    for item in itens:
        tupla = (item["name"], item["ramal"], item["building"], item["predio"], item["andar"])
        DATA.append(tupla)

    if itens:
        _recalcular_ramais_compartilhados()


def salvar_ramal_criado(name: str, ramal: str, building: str, predio: str, andar: str) -> None:
    try:
        try:
            with open(RAMAIS_CRIADOS_FILE, "r", encoding="utf-8") as f:
                itens = json.load(f)
        except FileNotFoundError:
            itens = []
        itens.append(
            {"name": name, "ramal": ramal, "building": building, "predio": predio, "andar": andar}
        )
        with open(RAMAIS_CRIADOS_FILE, "w", encoding="utf-8") as f:
            json.dump(itens, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Falha ao salvar %s", RAMAIS_CRIADOS_FILE)


# ── Checklist de inventário (conferência física dos ramais) ─────────────
# Marca quais ramais já foram conferidos num ciclo de inventário. Chave é
# o índice em DATA, valor tem se foi conferido e quando/por quem.
INVENTARIO_FILE = "inventario.json"
INVENTARIO: dict = {}  # idx (int) -> {"checado": bool, "quando": str, "por": int|None}


def load_inventario() -> None:
    global INVENTARIO
    try:
        with open(INVENTARIO_FILE, "r", encoding="utf-8") as f:
            bruto = json.load(f)
        INVENTARIO = {int(k): v for k, v in bruto.items()}
    except FileNotFoundError:
        INVENTARIO = {}
    except Exception:
        logger.exception("Falha ao carregar %s", INVENTARIO_FILE)
        INVENTARIO = {}


def save_inventario() -> None:
    try:
        with open(INVENTARIO_FILE, "w", encoding="utf-8") as f:
            json.dump({str(k): v for k, v in INVENTARIO.items()}, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Falha ao salvar %s", INVENTARIO_FILE)


def esta_conferido(idx: int) -> bool:
    return bool(INVENTARIO.get(idx, {}).get("checado"))


def get_effective_devices(name: str, ramal: str) -> list:
    """Retorna a lista de aparelhos do setor com as edições de admin aplicadas.

    Também inclui aparelhos que não existiam no inventário original mas que
    um admin cadastrou do zero (ex.: setor sem equipamento registrado).
    """
    devices = DEVICE_INFO.get((name, ramal), [])
    result = []
    for i, dev in enumerate(devices):
        merged = dict(dev)
        merged.update(DEVICE_OVERRIDES.get(_override_key(name, ramal, i), {}))
        result.append(merged)

    prefix = f"{name}|{ramal}|"
    extra_indices = set()
    for key in DEVICE_OVERRIDES:
        if key.startswith(prefix):
            try:
                i = int(key[len(prefix):])
            except ValueError:
                continue
            if i >= len(devices):
                extra_indices.add(i)

    for i in sorted(extra_indices):
        while len(result) <= i:
            missing_i = len(result)
            base = {"pat": None, "n_serie": None, "mac": None, "ip": None, "tronco": None}
            base.update(DEVICE_OVERRIDES.get(_override_key(name, ramal, missing_i), {}))
            result.append(base)

    return result

# ── Menu por prédio / andar ─────────────────────────────────────────────
# Código curto (usado no callback_data) -> (nome do prédio como está em DATA, rótulo com ícone)
PREDIO_CODES = {
    "HU": ("HU", "🏥 HU"),
    "ADM": ("ADM", "🏢 ADM"),
    "PA": ("PA", "🚑 PA (Pronto Atendimento)"),
    "CG": ("Casa da gestante", "🤰 Casa da Gestante"),
    "DIR": ("Direção", "🏛️ Direção"),
    "MAN": ("Manutenção", "🔧 Manutenção"),
}
PREDIO_NAME_TO_CODE = {v[0]: k for k, v in PREDIO_CODES.items()}


def _andar_sort_key(andar: str):
    match = re.match(r"(\d+)", andar)
    return int(match.group(1)) if match else 0


def predios_disponiveis():
    """Lista códigos de prédio que realmente têm setores cadastrados (não
    excluídos), na ordem de PREDIO_CODES."""
    presentes = {DATA[i][3] for i in indices_ativos()}
    return [code for code, (nome, _label) in PREDIO_CODES.items() if nome in presentes]


def andares_do_predio(predio_nome: str):
    andares = {DATA[i][4] for i in indices_ativos() if DATA[i][3] == predio_nome}
    return sorted(andares, key=_andar_sort_key)


def setores_do_predio_andar(predio_nome: str, andar: str):
    """Retorna pares (idx, item) dos setores não excluídos desse prédio/andar,
    ordenados por nome exibido."""
    itens = [
        (i, DATA[i]) for i in indices_ativos()
        if DATA[i][3] == predio_nome and DATA[i][4] == andar
    ]
    return sorted(itens, key=lambda par: normalize(nome_exibido(par[0])))


def _menu_button_row() -> list:
    return [InlineKeyboardButton("🏠 Menu Inicial", callback_data="menu")]


async def _responder_query_seguro(query, *args, **kwargs) -> None:
    """Chama query.answer() sem derrubar o handler se o clique já estiver
    velho demais (ex.: bot reiniciando, fila de updates acumulada) — o
    Telegram recusa com 'Query is too old...' nesses casos, o que é
    inofensivo e não deveria gerar traceback no log."""
    try:
        await query.answer(*args, **kwargs)
    except TelegramError as exc:
        logger.debug("Callback query expirado/ inválido, ignorando: %s", exc)


def _ping_sync(ip: str) -> bool:
    """Executa um ping ICMP simples (1 pacote) para o IP informado."""
    is_windows = platform.system().lower() == "windows"
    cmd = ["ping", "-n", "1", "-w", "1500", ip] if is_windows else ["ping", "-c", "1", "-W", "2", ip]
    try:
        resultado = subprocess.run(cmd, capture_output=True, timeout=5)
        return resultado.returncode == 0
    except Exception:
        return False


async def ping_host(ip: str) -> bool:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _ping_sync, ip)


def _normalize_mac(mac: str) -> str:
    return re.sub(r"[^0-9a-fA-F]", "", mac or "").lower()


def _refresh_arp_cache_sync(prefixos: set) -> None:
    """Manda um ping de broadcast em cada sub-rede
    informada para forçar os aparelhos a responderem e aparecerem na
    tabela ARP do sistema operacional."""
    if platform.system().lower() == "windows":
        return
    for prefixo in prefixos:
        try:
            subprocess.run(
                ["ping", "-b", "-c", "1", "-W", "1", f"{prefixo}.255"],
                capture_output=True, timeout=3,
            )
        except Exception:
            pass


def _read_system_arp_table_sync() -> dict:
    tabela = {}
    for cmd in (["ip", "neigh"], ["arp", "-an"]):
        try:
            resultado = subprocess.run(cmd, capture_output=True, timeout=5, text=True)
            saida = resultado.stdout
        except Exception:
            continue
        for linha in saida.splitlines():
            m = re.search(
                r"(\d{1,3}(?:\.\d{1,3}){3}).*?([0-9a-fA-F]{2}(?::[0-9a-fA-F]{2}){5})", linha
            )
            if m:
                tabela[_normalize_mac(m.group(2))] = m.group(1)
        if tabela:
            break
    return tabela


def _build_mac_ip_table_sync(prefixos: set) -> dict:
    """Descobre o IP atual de cada MAC nas sub-redes informadas usando só
    ferramentas padrão do sistema (sem depender de pacotes extras): manda
    um ping de broadcast em cada sub-rede para forçar os aparelhos a
    responderem, e depois lê a tabela ARP do sistema."""
    _refresh_arp_cache_sync(prefixos)
    return _read_system_arp_table_sync()


_SUBNET_SCAN_CACHE: dict = {}     # prefixo -> (timestamp, tabela)
_SUBNET_SCAN_INFLIGHT: dict = {}  # prefixo -> asyncio.Task em andamento
_SUBNET_SCAN_TTL = 30  # segundos — evita reescanear a mesma sub-rede várias
# vezes seguidas quando vários aparelhos offline caem na mesma sub-rede.


async def _scan_subnet_cached(prefixo: str) -> dict:
    """Escaneia (via ARP) uma sub-rede /24 em busca de MAC -> IP atual,
    reaproveitando resultado recente em cache e evitando escanear a mesma
    sub-rede em paralelo quando várias chamadas caem no mesmo momento."""
    agora = time.monotonic()
    cache = _SUBNET_SCAN_CACHE.get(prefixo)
    if cache and agora - cache[0] < _SUBNET_SCAN_TTL:
        return cache[1]

    tarefa = _SUBNET_SCAN_INFLIGHT.get(prefixo)
    if tarefa is None:
        loop = asyncio.get_event_loop()
        tarefa = loop.run_in_executor(None, _build_mac_ip_table_sync, {prefixo})
        _SUBNET_SCAN_INFLIGHT[prefixo] = tarefa

    try:
        tabela = await tarefa
    finally:
        _SUBNET_SCAN_INFLIGHT.pop(prefixo, None)

    _SUBNET_SCAN_CACHE[prefixo] = (agora, tabela)
    return tabela


async def resolve_and_ping(dev: dict) -> tuple:
    """Testa se um aparelho está online. Primeiro tenta o IP cadastrado
    manualmente; se não responder e houver MAC cadastrado, procura na rede
    (ARP) qual é o IP atual daquele MAC — útil quando o IP mudou
    por DHCP — e testa de novo nesse IP.

    Retorna (status, ip_via_mac):
    - status: True/False, ou None se o aparelho não tem IP nem MAC para testar.
    - ip_via_mac: o IP em que o aparelho foi encontrado pelo MAC, só quando
      diferente do IP cadastrado manualmente (senão None)."""
    ip_cadastrado = dev.get("ip")
    if ip_cadastrado and await ping_host(ip_cadastrado):
        return True, None

    mac = dev.get("mac")
    if mac and ip_cadastrado:
        partes = ip_cadastrado.split(".")
        if len(partes) == 4:
            prefixo = ".".join(partes[:3])
            tabela = await _scan_subnet_cached(prefixo)
            ip_atual = tabela.get(_normalize_mac(mac))
            if ip_atual and ip_atual != ip_cadastrado:
                ok = await ping_host(ip_atual)
                return ok, (ip_atual if ok else None)

    return (False, None) if ip_cadastrado else (None, None)


def build_predio_keyboard(user_id: int | None = None) -> "InlineKeyboardMarkup":
    botoes = [
        [InlineKeyboardButton("📋 Listar todos os telefones (A-Z)", callback_data="listar_todos")],
        [InlineKeyboardButton("🏢 Listar por prédio/andar", callback_data="listar_por_local")],
        [InlineKeyboardButton("🔴 Listar somente offline", callback_data="listar_offline")],
        [InlineKeyboardButton("🏷️ Buscar por patrimônio (PAT)", callback_data="buscar_pat")],
        [InlineKeyboardButton("🔢 Buscar por ramal", callback_data="buscar_ramal")],
        [InlineKeyboardButton("📤 Exportar lista (CSV)", callback_data="exportar_lista")],
        [InlineKeyboardButton("🧾 Checklist de inventário", callback_data="checklist_inventario")],
    ]
    if is_admin(user_id):
        botoes.append([InlineKeyboardButton("➕ Criar novo ramal", callback_data="criar_ramal")])
        if RAMAIS_EXCLUIDOS:
            botoes.append([InlineKeyboardButton(
                f"🗑️ Ramais excluídos ({len(RAMAIS_EXCLUIDOS)})", callback_data="ramais_excluidos",
            )])
    for code in predios_disponiveis():
        _nome, label = PREDIO_CODES[code]
        botoes.append([InlineKeyboardButton(label, callback_data=f"pr:{code}")])
    return InlineKeyboardMarkup(botoes)


def build_andar_keyboard(code: str) -> "InlineKeyboardMarkup":
    nome, _label = PREDIO_CODES[code]
    botoes = [
        [InlineKeyboardButton(f"{andar} andar", callback_data=f"an:{code}:{andar}")]
        for andar in andares_do_predio(nome)
    ]
    botoes.append([InlineKeyboardButton("🔙 Voltar aos prédios", callback_data="menu")])
    return InlineKeyboardMarkup(botoes)


def build_setores_texto(code: str, andar: str) -> str:
    nome, label = PREDIO_CODES[code]
    qtd = len(setores_do_predio_andar(nome, andar))
    return f"{label} — {andar} andar\n\nToque em um setor para ver o ramal ({qtd} setores):"


def build_setores_keyboard(code: str, andar: str) -> "InlineKeyboardMarkup":
    nome, _label = PREDIO_CODES[code]
    itens = setores_do_predio_andar(nome, andar)
    back_token = f"an:{code}:{andar}"
    botoes = []
    for idx, _item in itens:
        botoes.append([_sector_button(idx, back_token)])
    botoes.append([InlineKeyboardButton("🔙 Voltar aos andares", callback_data=f"pr:{code}")])
    botoes.append(_menu_button_row())
    return InlineKeyboardMarkup(botoes)


async def get_full_listing_ping_status() -> dict:
    """Testa o ping do primeiro aparelho de cada ramal, em paralelo. Tenta o
    IP cadastrado primeiro e, se não responder, procura o IP atual pelo MAC
    (caso o IP tenha mudado por DHCP). Retorna (name, ramal) -> (status, ip_via_mac).
    Ramais sem IP nem MAC não entram."""
    tarefas = {}
    for i in indices_ativos():
        name, ramal, _building, _predio, _andar = DATA[i]
        devices = [d for d in get_effective_devices(name, ramal) if not d.get("excluido")]
        if devices and (devices[0].get("ip") or devices[0].get("mac")):
            tarefas[(name, ramal)] = resolve_and_ping(devices[0])
    if not tarefas:
        return {}
    resultados = await asyncio.gather(*tarefas.values())
    return {chave: r for chave, r in zip(tarefas.keys(), resultados) if r[0] is not None}


def build_export_csv() -> bytes:
    """Monta um CSV com todos os ramais e as informações de equipamento
    cadastradas (PAT, N° Série, MAC, IP, N° Tronco, etc.), pra exportar."""
    buffer = io.StringIO()
    writer = csv.writer(buffer, delimiter=";")
    writer.writerow(
        [
            "Setor", "Ramal", "Prédio", "Andar", "PAT", "N° Série", "MAC", "IP",
            "N° Tronco", "Originar chamada", "Receber chamada",
        ]
    )
    for idx in indices_ativos():
        name, ramal, _building, predio, andar = DATA[idx]
        nome_atual = nome_exibido(idx)
        rc = ramal_com_fila(name, ramal_exibido(idx))
        devices = [d for d in get_effective_devices(name, ramal) if not d.get("excluido")]
        if not devices:
            writer.writerow([nome_atual, rc, predio, andar, "", "", "", "", "", "", ""])
            continue
        for dev in devices:
            writer.writerow(
                [
                    nome_atual,
                    rc,
                    predio,
                    andar,
                    dev.get("pat") or "",
                    dev.get("n_serie") or "",
                    dev.get("mac") or "",
                    dev.get("ip") or "",
                    dev.get("tronco") or "",
                    "Sim" if dev.get("originar", True) else "Não",
                    "Sim" if dev.get("receber", True) else "Não",
                ]
            )
    return buffer.getvalue().encode("utf-8-sig")  # BOM: acentos abrem certo no Excel


def build_checklist_chunks() -> list:
    """Monta a lista de checklist de inventário (✅ conferido / ⬜ pendente),
    agrupada por letra, com um resumo no topo. Dividida em blocos que
    respeitam o limite de tamanho de mensagem do Telegram."""
    itens = sorted(((i, DATA[i]) for i in indices_ativos()), key=lambda item: normalize(nome_exibido(item[0])))
    total = len(itens)
    conferidos = sum(1 for idx, _dados in itens if esta_conferido(idx))

    grouped = {}
    for idx, (name, ramal, _building, _predio, _andar) in itens:
        nome_atual = nome_exibido(idx)
        letter = _strip_accents(nome_atual[0]).upper()
        grouped.setdefault(letter, []).append((idx, nome_atual, name, ramal))

    chunks = []
    current = (
        "🧾 <b>Checklist de inventário</b>\n"
        f"✅ Conferidos: {conferidos}/{total}\n"
        f"⬜ Pendentes: {total - conferidos}\n\n"
    )
    for letter in sorted(grouped.keys()):
        section = f"<b>— {letter} —</b>\n"
        for idx, nome_atual, name, ramal in grouped[letter]:
            marca = "✅" if esta_conferido(idx) else "⬜"
            section += f"• {marca} {nome_atual} — <b>{ramal_com_fila(name, ramal_exibido(idx))}</b>\n"
        section += "\n"

        if len(current) + len(section) > TELEGRAM_MSG_LIMIT:
            chunks.append(current.rstrip())
            current = section
        else:
            current += section

    if current.strip():
        chunks.append(current.rstrip())
    return chunks





def build_full_listing_chunks(ping_status: dict | None = None, apenas_offline: bool = False) -> list:
    """Monta a lista completa de ramais agrupada por letra, dividida em blocos
    que respeitam o limite de tamanho de mensagem do Telegram.
    Se apenas_offline=True, mostra só os ramais cujo ping falhou (🔴).
    ping_status: dict (name, ramal) -> (status, ip_via_mac)."""
    ping_status = ping_status or {}
    itens = sorted(((i, DATA[i]) for i in indices_ativos()), key=lambda item: normalize(nome_exibido(item[0])))
    grouped = {}
    for idx, (name, ramal, building, predio, andar) in itens:
        nome_atual = nome_exibido(idx)
        entry = ping_status.get((name, ramal))
        status = entry[0] if entry else None
        if apenas_offline and status is not False:
            continue
        letter = _strip_accents(nome_atual[0]).upper()
        grouped.setdefault(letter, []).append((idx, name, nome_atual, ramal, building, predio, andar))

    if not grouped:
        return []

    titulo = "🔴 <b>Ramais offline</b>" if apenas_offline else "📋 <b>Lista completa de ramais</b>"
    chunks = []
    current = f"{titulo}\n🟢 Online · 🔴 Sem resposta · ⚪ Sem IP cadastrado\n📡 = achado por MAC em outro IP\n\n"
    for letter in sorted(grouped.keys()):
        section = f"<b>— {letter} —</b>\n"
        for idx, name, nome_atual, ramal, building, predio, andar in grouped[letter]:
            local = f"{predio} — {andar} andar" if predio != "?" else "local não confirmado"
            entry = ping_status.get((name, ramal))
            status = entry[0] if entry else None
            ip_via_mac = entry[1] if entry else None
            bola = "🟢" if status is True else ("🔴" if status is False else "⚪")
            extra = f" 📡 <i>IP atual: {ip_via_mac}</i>" if ip_via_mac else ""
            section += f"• {bola} {nome_atual} — <b>{ramal_com_fila(name, ramal_exibido(idx))}</b> ({local}){extra}\n"
        section += "\n"

        if len(current) + len(section) > TELEGRAM_MSG_LIMIT:
            chunks.append(current.rstrip())
            current = section
        else:
            current += section

    if current.strip():
        chunks.append(current.rstrip())
    return chunks


def build_full_listing_by_location_chunks(ping_status: dict | None = None, apenas_offline: bool = False) -> list:
    """Como build_full_listing_chunks, mas agrupada por prédio e andar em
    vez de por nome (ex.: 'Prédio ADM — 1º Piso' com os ramais embaixo)."""
    ping_status = ping_status or {}
    grouped = {}
    for idx in indices_ativos():
        name, ramal, _building, predio, andar = DATA[idx]
        entry = ping_status.get((name, ramal))
        status = entry[0] if entry else None
        if apenas_offline and status is not False:
            continue
        grouped.setdefault((predio, andar), []).append((idx, name, nome_exibido(idx), ramal))

    if not grouped:
        return []

    predios_ordenados = [PREDIO_CODES[code][0] for code in predios_disponiveis()]
    outros_predios = sorted({p for (p, _a) in grouped.keys() if p not in predios_ordenados})
    ordem_predios = predios_ordenados + outros_predios

    titulo = "🔴 <b>Ramais offline por prédio/andar</b>" if apenas_offline else "📋 <b>Lista por prédio/andar</b>"
    chunks = []
    current = f"{titulo}\n🟢 Online · 🔴 Sem resposta · ⚪ Sem IP cadastrado\n📡 = achado por MAC em outro IP\n\n"

    for predio_nome in ordem_predios:
        andares_presentes = sorted(
            {a for (p, a) in grouped.keys() if p == predio_nome}, key=_andar_sort_key
        )
        for andar in andares_presentes:
            itens = sorted(grouped[(predio_nome, andar)], key=lambda t: normalize(t[2]))
            if predio_nome == "?":
                cabecalho = "<b>Local não confirmado</b>\n"
            else:
                titulo_andar = f"{andar} Piso" if andar != "?" else "andar não confirmado"
                cabecalho = f"<b>Prédio {predio_nome} — {titulo_andar}</b>\n"
            section = cabecalho
            for idx, name, nome_atual, ramal in itens:
                entry = ping_status.get((name, ramal))
                status = entry[0] if entry else None
                ip_via_mac = entry[1] if entry else None
                bola = "🟢" if status is True else ("🔴" if status is False else "⚪")
                extra = f" 📡 <i>IP atual: {ip_via_mac}</i>" if ip_via_mac else ""
                section += f"• {bola} {nome_atual} — <b>{ramal_com_fila(name, ramal_exibido(idx))}</b>{extra}\n"
            section += "\n"

            if len(current) + len(section) > TELEGRAM_MSG_LIMIT:
                chunks.append(current.rstrip())
                current = section
            else:
                current += section

    if current.strip():
        chunks.append(current.rstrip())
    return chunks


def _strip_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(c for c in normalized if not unicodedata.combining(c))


def normalize(text: str) -> str:
    return _strip_accents(text).lower().strip()


# Ramais que aparecem em mais de um setor com nomes diferentes: indicam uma
# fila/linha compartilhada (o mesmo número toca em mais de um lugar), mesmo
# quando a planilha não tem um número de fila separado cadastrado.
_RAMAIS_COMPARTILHADOS = {
    ramal for ramal, qtd in Counter(item[1] for item in DATA).items() if qtd > 1
}


def ramal_com_fila(name: str, ramal: str) -> str:
    """Retorna o texto do ramal, acrescentando a fila ao lado quando existir.
    O rótulo automático "(Fila compartilhada)" — que aparecia sempre que o
    mesmo número de ramal se repetia em mais de um setor — foi removido a
    pedido; números repetidos continuam funcionando normalmente, só não
    mostram mais esse texto ao lado."""
    fila = FILA_INFO.get((name, ramal))
    if fila:
        return f"{ramal} (Fila {fila})"
    return ramal


def search_indices(term: str, limit: int = MAX_RESULTS):
    """Como search(), mas retorna os índices em DATA (para montar botões).
    Casa por nome do setor, número de ramal, número de PAT ou N° Auxiliar
    do aparelho (o N° Auxiliar é tratado como mais um número de ramal)."""
    term_n = normalize(term)
    if not term_n:
        return []
    idxs = []
    for i, (name, ramal, building, predio, andar) in enumerate(DATA):
        if ramal_excluido(i):
            continue
        if term_n in normalize(nome_exibido(i)) or term_n in ramal_exibido(i):
            idxs.append(i)
            continue
        devices = get_effective_devices(name, ramal)
        if any(
            not dev.get("excluido") and (
                (dev.get("pat") and term_n in normalize(str(dev.get("pat"))))
                or (dev.get("aux") and term_n in normalize(str(dev.get("aux"))))
            )
            for dev in devices
        ):
            idxs.append(i)
    return idxs[:limit]


def search_indices_by_pat(term: str, limit: int = MAX_RESULTS):
    """Busca só pelo número de PAT (patrimônio) do aparelho, ignorando
    nome do setor e ramal — usado no botão dedicado 'Buscar por patrimônio'."""
    term_n = normalize(term)
    if not term_n:
        return []
    idxs = []
    for i, (name, ramal, building, predio, andar) in enumerate(DATA):
        if ramal_excluido(i):
            continue
        devices = get_effective_devices(name, ramal)
        if any(term_n in normalize(str(dev.get("pat") or "")) for dev in devices if dev.get("pat") and not dev.get("excluido")):
            idxs.append(i)
    return idxs[:limit]


def encontrar_por_pat_exato(numero: str) -> list:
    """Acha os índices de DATA cujo aparelho tem esse PAT exatamente igual
    (ignorando zeros à esquerda) — usado na leitura de código de barras."""
    alvo = numero.strip().lstrip("0") or "0"
    idxs = []
    for i, (name, ramal, building, predio, andar) in enumerate(DATA):
        if ramal_excluido(i):
            continue
        devices = get_effective_devices(name, ramal)
        for dev in devices:
            if dev.get("excluido"):
                continue
            pat = str(dev.get("pat") or "").strip().lstrip("0") or "0"
            if pat and pat == alvo:
                idxs.append(i)
                break
    return idxs


def search_indices_by_ramal(term: str, limit: int = MAX_RESULTS):
    """Busca só pelo número do ramal (ou N° Auxiliar de algum aparelho da
    linha, que conta como mais um número de ramal), ignorando nome do setor
    e PAT — usado no botão dedicado 'Buscar por ramal'."""
    term_n = normalize(term)
    if not term_n:
        return []
    idxs = []
    for i, (name, ramal, building, predio, andar) in enumerate(DATA):
        if ramal_excluido(i):
            continue
        if term_n in ramal_exibido(i):
            idxs.append(i)
            continue
        devices = get_effective_devices(name, ramal)
        if any(
            not dev.get("excluido") and dev.get("aux") and term_n in normalize(str(dev.get("aux")))
            for dev in devices
        ):
            idxs.append(i)
    return idxs[:limit]


def search(term: str, limit: int = MAX_RESULTS):
    """Busca setores por nome (ignorando acentos/caixa) ou por ramal exato/parcial."""
    return [DATA[i] for i in search_indices(term, limit)]


def format_results(term: str, results) -> str:
    if not results:
        return (
            f"Nenhum ramal encontrado para \"{term}\".\n"
            "Tente outro termo (ex: parte do nome do setor, o número do ramal ou o PAT do aparelho)."
        )
    lines = [f"Resultados para \"{term}\":\n"]
    for name, ramal, building, predio, andar in results:
        local = f"{predio} — {andar} andar" if predio != "?" else "local não confirmado"
        lines.append(f"• {name} — <b>{ramal_com_fila(name, ramal)}</b> ({local})")
    if len(results) == MAX_RESULTS:
        lines.append("\n(mostrando os primeiros resultados — refine sua busca)")
    return "\n".join(lines)


def _sector_button(idx: int, back_token: str) -> "InlineKeyboardButton":
    name, ramal, *_ = DATA[idx]
    label = f"{nome_exibido(idx)} — {ramal_com_fila(name, ramal_exibido(idx))}"
    if len(label) > 64:
        label = label[:61] + "..."
    return InlineKeyboardButton(label, callback_data=f"det:{idx}:{back_token}")


def build_result_keyboard(indices, back_token: str) -> "InlineKeyboardMarkup":
    botoes = [[_sector_button(idx, back_token)] for idx in indices]
    botoes.append(_menu_button_row())
    return InlineKeyboardMarkup(botoes)


def _ping_badge(status) -> str:
    if status is True:
        return " 🟢"
    if status is False:
        return " 🔴"
    return ""


def build_detail_text(idx: int, ping_status: dict | None = None) -> str:
    name, ramal, _building, predio, andar = DATA[idx]
    ping_status = ping_status or {}
    texto = f"📞 <b>{nome_exibido(idx)}</b>\n\n"
    if ramal_excluido(idx):
        texto += "🗑️ <b>Este ramal está excluído</b> — some das buscas/listagens.\n\n"
    texto += f"Ramal: <b>{ramal_com_fila(name, ramal_exibido(idx))}</b>\n"
    if predio != "?":
        texto += f"Local: {predio}\nAndar: {andar}\n"
    else:
        texto += "Local: não confirmado\n"

    info_check = INVENTARIO.get(idx, {})
    if info_check.get("checado"):
        quando = info_check.get("quando", "")
        texto += f"🧾 Inventário: ✅ Conferido{f' em {quando}' if quando else ''}\n"
    else:
        texto += "🧾 Inventário: ⬜ Ainda não conferido\n"

    devices = get_effective_devices(name, ramal)
    visiveis = [d for d in devices if not d.get("excluido")]
    if not visiveis:
        texto += "\n☎️ Equipamento: <i>não cadastrado no inventário</i>\n"
    else:
        texto += "\n☎️ <b>Equipamento</b>"
        if len(visiveis) > 1:
            texto += f" ({len(visiveis)} aparelhos nesta linha)"
        texto += "\n"
        for i, dev in enumerate(devices):
            if dev.get("excluido"):
                continue
            pat = dev["pat"] or "—"
            n_serie = dev["n_serie"] or "—"
            mac = dev["mac"] or "—"
            ip = dev["ip"] or "—"
            tronco = dev.get("tronco") or "—"
            originar = dev.get("originar", True)
            receber = dev.get("receber", True)
            entry = ping_status.get(f"{idx}:{i}")
            status = entry[0] if entry else None
            ip_via_mac = entry[1] if entry else None
            badge = _ping_badge(status) if (dev["ip"] or dev.get("mac")) else ""
            if len(visiveis) > 1:
                texto += f"\n<b>Aparelho {i + 1}</b>{badge}\n"
            linha_ip_mac = f"\n📡 <i>Encontrado por MAC no IP atual: <code>{ip_via_mac}</code></i>" if ip_via_mac else ""
            aux = dev.get("aux") or "—"
            texto += (
                f"PAT: <code>{pat}</code>\n"
                f"N° Série: <code>{n_serie}</code>\n"
                f"MAC: <code>{mac}</code>\n"
                f"IP: <code>{ip}</code>{badge}{linha_ip_mac}\n"
                f"N° Tronco: <code>{tronco}</code>\n"
                f"Originar chamada: {'🟢' if originar else '🔴'}\n"
                f"Receber chamada: {'🟢' if receber else '🔴'}\n"
                f"N° Auxiliar: <code>{aux}</code>\n"
            )

    texto += f"\nTelefone Geral: {GENERAL_PHONE}"
    return texto


def build_detail_keyboard(
    idx: int,
    back_token: str,
    ping_status: dict | None = None,
    user_id: int | None = None,
) -> "InlineKeyboardMarkup":
    name, ramal, _building, _predio, _andar = DATA[idx]
    devices = get_effective_devices(name, ramal)
    ping_status = ping_status or {}

    botoes = []
    for i, dev in enumerate(devices):
        if dev.get("excluido"):
            continue
        ip = dev["ip"]
        if not ip and not dev.get("mac"):
            continue
        entry = ping_status.get(f"{idx}:{i}")
        status = entry[0] if entry else None
        ip_via_mac = entry[1] if entry else None
        prefixo = f"Aparelho {i + 1}: " if len(devices) > 1 else ""
        rotulo_alvo = f"{ip_via_mac} 📡" if ip_via_mac else (ip or "IP via MAC")
        if status is True:
            label = f"🟢 {prefixo}{rotulo_alvo} — Online (testar de novo)"
        elif status is False:
            label = f"🔴 {prefixo}{rotulo_alvo} — Sem resposta (testar de novo)"
        else:
            label = f"📡 Testar ping — {prefixo}{rotulo_alvo}"
        if len(label) > 64:
            label = label[:61] + "..."
        botoes.append([InlineKeyboardButton(label, callback_data=f"ping:{idx}:{i}:{back_token}")])

    if is_admin(user_id):
        conferido = esta_conferido(idx)
        botoes.append(
            [InlineKeyboardButton(
                "✅ Conferido no inventário (toque p/ desmarcar)" if conferido
                else "⬜ Marcar como conferido no inventário",
                callback_data=f"check_inventario:{idx}:{back_token}",
            )]
        )
        botoes.append(
            [InlineKeyboardButton("✏️ Editar nome do setor", callback_data=f"editnome:{idx}:{back_token}")]
        )
        botoes.append(
            [InlineKeyboardButton("✏️ Editar número do ramal", callback_data=f"editramal:{idx}:{back_token}")]
        )

        visiveis = [i for i, dev in enumerate(devices) if not dev.get("excluido")]
        excluidos = [i for i, dev in enumerate(devices) if dev.get("excluido")]

        if not visiveis:
            botoes.append(
                [InlineKeyboardButton("✏️ Cadastrar PAT/N° Série/MAC/IP/Tronco", callback_data=f"edit:{idx}:0:{back_token}")]
            )
        else:
            for i in visiveis:
                prefixo = f"Aparelho {i + 1}: " if len(devices) > 1 else ""
                botoes.append(
                    [InlineKeyboardButton(f"✏️ Editar {prefixo}PAT/N° Série/MAC/IP/Tronco", callback_data=f"edit:{idx}:{i}:{back_token}")]
                )
                originar = devices[i].get("originar", True)
                receber = devices[i].get("receber", True)
                botoes.append(
                    [InlineKeyboardButton(
                        f"🔁 {prefixo}Originar chamada: {'🟢' if originar else '🔴'}",
                        callback_data=f"togcall:{idx}:{i}:originar:{back_token}",
                    )]
                )
                botoes.append(
                    [InlineKeyboardButton(
                        f"🔁 {prefixo}Receber chamada: {'🟢' if receber else '🔴'}",
                        callback_data=f"togcall:{idx}:{i}:receber:{back_token}",
                    )]
                )
                rotulo_excluir = f"🗑️ Excluir Aparelho {i + 1}" if len(devices) > 1 else "🗑️ Excluir aparelho"
                botoes.append(
                    [InlineKeyboardButton(
                        rotulo_excluir,
                        callback_data=f"excluir_aparelho_confirma:{idx}:{i}:{back_token}",
                    )]
                )

            # Ramal já tem pelo menos um aparelho — mesmo assim, permite
            # cadastrar mais um (linha com vários aparelhos/ramais compartilhados).
            botoes.append(
                [InlineKeyboardButton(
                    "➕ Cadastrar novo aparelho",
                    callback_data=f"edit:{idx}:{len(devices)}:{back_token}",
                )]
            )

        for i in excluidos:
            prefixo = f"Aparelho {i + 1}" if len(devices) > 1 else "aparelho"
            botoes.append(
                [InlineKeyboardButton(
                    f"↩️ Restaurar {prefixo} (excluído)",
                    callback_data=f"restaurar_aparelho:{idx}:{i}:{back_token}",
                )]
            )

        if ramal_excluido(idx):
            botoes.append(
                [InlineKeyboardButton("↩️ Restaurar este ramal (excluído)", callback_data=f"restaurar_ramal:{idx}:{back_token}")]
            )
        else:
            botoes.append(
                [InlineKeyboardButton("🗑️ Excluir este ramal", callback_data=f"excluir_ramal_confirma:{idx}:{back_token}")]
            )

    botoes.append([InlineKeyboardButton("🔙 Voltar", callback_data=back_token)])
    botoes.append(_menu_button_row())
    return InlineKeyboardMarkup(botoes)


def build_edit_fields_keyboard(idx: int, dev_i: int, back_token: str) -> "InlineKeyboardMarkup":
    botoes = [
        [InlineKeyboardButton("Editar PAT", callback_data=f"campo:{idx}:{dev_i}:pat:{back_token}")],
        [InlineKeyboardButton("Editar N° Série", callback_data=f"campo:{idx}:{dev_i}:n_serie:{back_token}")],
        [InlineKeyboardButton("Editar MAC", callback_data=f"campo:{idx}:{dev_i}:mac:{back_token}")],
        [InlineKeyboardButton("Editar IP", callback_data=f"campo:{idx}:{dev_i}:ip:{back_token}")],
        [InlineKeyboardButton("Editar N° Tronco", callback_data=f"campo:{idx}:{dev_i}:tronco:{back_token}")],
        [InlineKeyboardButton("Editar N° Auxiliar", callback_data=f"campo:{idx}:{dev_i}:aux:{back_token}")],
        [InlineKeyboardButton("🔙 Voltar", callback_data=f"det:{idx}:{back_token}")],
    ]
    return InlineKeyboardMarkup(botoes)


# ── Inatividade ───────────────────────────────────────────────────────────
INACTIVITY_SECONDS = 900  # 15 minutos


async def _voltar_menu_por_inatividade(context) -> None:
    chat_id = context.job.chat_id
    try:
        await context.bot.send_message(
            chat_id=chat_id,
            text=(
                "⏰ 1 minuto sem uso — voltando ao menu inicial.\n\n"
                "🏢 <b>Navegar por prédio e andar</b>\n\nEscolha um prédio:"
            ),
            parse_mode=ParseMode.HTML,
            reply_markup=build_predio_keyboard(chat_id),
        )
    except Exception:
        logger.warning("Não foi possível enviar o menu de inatividade para o chat %s", chat_id)


def _reiniciar_timer_inatividade(update: Update, context) -> None:
    if not context.job_queue:
        return
    chat_id = update.effective_chat.id
    nome = f"inactivity:{chat_id}"
    for job in context.job_queue.get_jobs_by_name(nome):
        job.schedule_removal()
    context.job_queue.run_once(
        _voltar_menu_por_inatividade, INACTIVITY_SECONDS, chat_id=chat_id, name=nome
    )


# ── Handlers ──────────────────────────────────────────────────────────────

async def start(update: Update, context) -> None:
    _reiniciar_timer_inatividade(update, context)
    user_id = update.effective_user.id if update.effective_user else None
    await update.message.reply_text(
        "📞 <b>Guia de Ramais VOIP</b>\n"
        "Hospital Universitário São Francisco de Paula\n\n"
        f"Telefone Geral: {GENERAL_PHONE}\n\n"
        "/listar — mostra todos os ramais cadastrados\n"
        "/menu — navega por prédio e andar\n\n"
        "Ou escolha um prédio abaixo:",
        parse_mode=ParseMode.HTML,
        reply_markup=build_predio_keyboard(user_id),
    )


async def ajuda(update: Update, context) -> None:
    await start(update, context)


async def meuid(update: Update, context) -> None:
    user = update.effective_user
    await update.message.reply_text(
        f"🆔 Seu ID do Telegram é: <code>{user.id}</code>\n\n"
        "Copie esse número e coloque em ADMIN_IDS_PADRAO (ou na variável de "
        "ambiente TELEGRAM_ADMIN_IDS) no arquivo Bot_VOIP.py para virar administrador.",
        parse_mode=ParseMode.HTML,
    )


async def chatid(update: Update, context) -> None:
    chat = update.effective_chat
    tipo = {"private": "conversa privada", "group": "grupo", "supergroup": "supergrupo", "channel": "canal"}.get(
        chat.type, chat.type
    )
    await update.message.reply_text(
        f"🆔 ID deste chat ({tipo}): <code>{chat.id}</code>",
        parse_mode=ParseMode.HTML,
    )


async def statusagora(update: Update, context) -> None:
    """Testa o ping de todos os ramais na hora e mostra um resumo de
    quantos estão online/offline no momento."""
    user_id = update.effective_user.id if update.effective_user else None
    if not is_admin(user_id):
        await update.message.reply_text("Apenas administradores podem usar esse comando.")
        return

    aviso = await update.message.reply_text("🔎 Testando o ping de todos os ramais agora...")
    ping_status = await get_full_listing_ping_status()
    online = sum(1 for status, _ in ping_status.values() if status is True)
    offline_itens = sorted(
        [(name, ramal) for (name, ramal), (status, _ip) in ping_status.items() if status is False],
        key=lambda item: normalize(item[0]),
    )

    texto = (
        f"📊 <b>Status agora</b>\n\n"
        f"🟢 Online: {online}\n"
        f"🔴 Offline: {len(offline_itens)}\n\n"
    )
    if offline_itens:
        texto += "<b>Offline no momento:</b>\n"
        texto += "\n".join(f"• {name} — {ramal}" for name, ramal in offline_itens)
    else:
        texto += "Nenhum ramal offline no momento."

    await aviso.delete()
    await update.message.reply_text(texto, parse_mode=ParseMode.HTML)


async def listar(update: Update, context) -> None:
    _reiniciar_timer_inatividade(update, context)
    aviso = await update.message.reply_text("🔎 Verificando status dos telefones...")
    ping_status = await get_full_listing_ping_status()
    await aviso.delete()
    chunks = build_full_listing_chunks(ping_status)
    for i, chunk in enumerate(chunks, start=1):
        footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
        is_last = i == len(chunks)
        await update.message.reply_text(
            chunk + footer,
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]) if is_last else None,
        )


async def menu(update: Update, context) -> None:
    _reiniciar_timer_inatividade(update, context)
    user_id = update.effective_user.id if update.effective_user else None
    await update.message.reply_text(
        "🏢 <b>Navegar por prédio e andar</b>\n\nEscolha um prédio:",
        parse_mode=ParseMode.HTML,
        reply_markup=build_predio_keyboard(user_id),
    )


async def _enviar_checklist(query, context, user_id: int | None) -> None:
    """Envia (ou reenvia, editando a mensagem atual quando possível) o
    checklist de inventário, com o botão de zerar pra admins no final."""
    chunks = build_checklist_chunks()
    for i, chunk in enumerate(chunks, start=1):
        footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
        is_last = i == len(chunks)
        markup = None
        if is_last:
            botoes_finais = []
            if is_admin(user_id):
                botoes_finais.append(
                    [InlineKeyboardButton("🔄 Zerar inventário (novo ciclo)", callback_data="zerar_inventario_confirma")]
                )
            botoes_finais.append(_menu_button_row())
            markup = InlineKeyboardMarkup(botoes_finais)
        if i == 1:
            await query.edit_message_text(chunk + footer, parse_mode=ParseMode.HTML, reply_markup=markup)
        else:
            await query.message.reply_text(chunk + footer, parse_mode=ParseMode.HTML, reply_markup=markup)


async def menu_callback(update: Update, context) -> None:
    _reiniciar_timer_inatividade(update, context)
    query = update.callback_query
    await _responder_query_seguro(query)
    data = query.data

    if data == "menu":
        user_id = update.effective_user.id if update.effective_user else None
        await query.edit_message_text(
            "🏢 <b>Navegar por prédio e andar</b>\n\nEscolha um prédio:",
            parse_mode=ParseMode.HTML,
            reply_markup=build_predio_keyboard(user_id),
        )
        return

    if data == "listar_todos":
        await query.edit_message_text("🔎 Verificando status dos telefones...")
        ping_status = await get_full_listing_ping_status()
        chunks = build_full_listing_chunks(ping_status)
        for i, chunk in enumerate(chunks, start=1):
            footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
            is_last = i == len(chunks)
            markup = InlineKeyboardMarkup([_menu_button_row()]) if is_last else None
            if i == 1:
                await query.edit_message_text(
                    chunk + footer,
                    parse_mode=ParseMode.HTML,
                    reply_markup=markup,
                )
            else:
                await query.message.reply_text(
                    chunk + footer,
                    parse_mode=ParseMode.HTML,
                    reply_markup=markup,
                )
        return

    if data == "listar_por_local":
        await query.edit_message_text("🔎 Verificando status dos telefones...")
        ping_status = await get_full_listing_ping_status()
        chunks = build_full_listing_by_location_chunks(ping_status)
        for i, chunk in enumerate(chunks, start=1):
            footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
            is_last = i == len(chunks)
            markup = InlineKeyboardMarkup([_menu_button_row()]) if is_last else None
            if i == 1:
                await query.edit_message_text(
                    chunk + footer,
                    parse_mode=ParseMode.HTML,
                    reply_markup=markup,
                )
            else:
                await query.message.reply_text(
                    chunk + footer,
                    parse_mode=ParseMode.HTML,
                    reply_markup=markup,
                )
        return

    if data == "listar_offline":
        await query.edit_message_text("🔎 Verificando status dos telefones...")
        ping_status = await get_full_listing_ping_status()
        chunks = build_full_listing_chunks(ping_status, apenas_offline=True)
        if not chunks:
            await query.edit_message_text(
                "✅ Nenhum ramal offline no momento.",
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return
        for i, chunk in enumerate(chunks, start=1):
            footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
            is_last = i == len(chunks)
            markup = InlineKeyboardMarkup([_menu_button_row()]) if is_last else None
            if i == 1:
                await query.edit_message_text(
                    chunk + footer,
                    parse_mode=ParseMode.HTML,
                    reply_markup=markup,
                )
            else:
                await query.message.reply_text(
                    chunk + footer,
                    parse_mode=ParseMode.HTML,
                    reply_markup=markup,
                )
        return

    if data == "buscar_pat":
        context.user_data["aguardando_busca_pat"] = True
        await query.edit_message_text(
            "🏷️ <b>Buscar por patrimônio (PAT)</b>\n\n"
            "Digite o número do PAT que deseja buscar:",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
        )
        return

    if data == "buscar_ramal":
        context.user_data["aguardando_busca_ramal"] = True
        await query.edit_message_text(
            "🔢 <b>Buscar por ramal</b>\n\n"
            "Digite o número do ramal que deseja buscar:",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
        )
        return

    if data == "exportar_lista":
        await _responder_query_seguro(query, "Gerando arquivo...")
        conteudo = build_export_csv()
        arquivo = io.BytesIO(conteudo)
        arquivo.name = "ramais_voip.csv"
        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=arquivo,
            filename="ramais_voip.csv",
            caption=f"📤 {len(indices_ativos())} ramais exportados.",
        )
        return

    if data == "checklist_inventario":
        user_id = update.effective_user.id if update.effective_user else None
        await _enviar_checklist(query, context, user_id)
        return

    if data.startswith("check_inventario:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem conferir o inventário.", show_alert=True)
            return

        atual = esta_conferido(idx)
        if atual:
            INVENTARIO.pop(idx, None)
        else:
            INVENTARIO[idx] = {
                "checado": True,
                "quando": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                "por": user_id,
            }
        save_inventario()

        ping_status = context.user_data.get("ping_status", {})
        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    if data == "zerar_inventario_confirma":
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem zerar o inventário.", show_alert=True)
            return
        await query.edit_message_text(
            "⚠️ Isso vai desmarcar <b>todos</b> os ramais como não conferidos, "
            "pra começar um novo ciclo de inventário. Confirma?",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("✅ Sim, zerar tudo", callback_data="zerar_inventario_ok")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data="checklist_inventario")],
                ]
            ),
        )
        return

    if data == "zerar_inventario_ok":
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem zerar o inventário.", show_alert=True)
            return
        INVENTARIO.clear()
        save_inventario()
        await _enviar_checklist(query, context, user_id)
        return

    if data == "criar_ramal":
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem criar ramais.", show_alert=True)
            return
        context.user_data["criando_ramal"] = {"etapa": "nome", "dados": {}}
        await query.edit_message_text(
            "➕ <b>Criar novo ramal</b>\n\nDigite o nome do setor.\n"
            "(Envie <code>cancelar</code> a qualquer momento para desistir.)",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "cancelar_criar_ramal":
        context.user_data.pop("criando_ramal", None)
        await query.edit_message_text(
            "❌ Criação de ramal cancelada. Nada foi salvo.",
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
        )
        return

    if data == "ramais_excluidos":
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem ver ramais excluídos.", show_alert=True)
            return
        idxs = sorted(RAMAIS_EXCLUIDOS, key=lambda i: normalize(nome_exibido(i)))
        if not idxs:
            await query.edit_message_text(
                "🗑️ Nenhum ramal excluído no momento.",
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return
        await query.edit_message_text(
            f"🗑️ <b>Ramais excluídos</b> ({len(idxs)})\n\nToque em um setor pra ver os detalhes e restaurar:",
            parse_mode=ParseMode.HTML,
            reply_markup=build_result_keyboard(idxs, "ramais_excluidos"),
        )
        return

    if data.startswith("excluir_ramal_confirma:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem excluir ramais.", show_alert=True)
            return
        name, ramal, _building, _predio, _andar = DATA[idx]
        await query.edit_message_text(
            f"⚠️ Tem certeza que quer excluir o setor <b>{nome_exibido(idx)}</b> "
            f"(ramal {ramal_com_fila(name, ramal_exibido(idx))})?\n\n"
            "Ele vai sumir das buscas, listagens e exportações — mas não apaga o "
            "cadastro dos aparelhos, dá pra restaurar depois em \"🗑️ Ramais excluídos\" no menu.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("🗑️ Sim, excluir", callback_data=f"excluir_ramal_ok:{idx}:{back_token}")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data=f"det:{idx}:{back_token}")],
                ]
            ),
        )
        return

    if data.startswith("excluir_ramal_ok:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem excluir ramais.", show_alert=True)
            return
        RAMAIS_EXCLUIDOS.add(idx)
        save_ramais_excluidos()
        _recalcular_ramais_compartilhados()
        await query.edit_message_text(
            f"🗑️ Setor <b>{nome_exibido(idx)}</b> excluído. Dá pra restaurar em \"🗑️ Ramais excluídos\" no menu.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
        )
        return

    if data.startswith("restaurar_ramal:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem restaurar ramais.", show_alert=True)
            return
        RAMAIS_EXCLUIDOS.discard(idx)
        save_ramais_excluidos()
        _recalcular_ramais_compartilhados()
        ping_status = context.user_data.get("ping_status", {})
        novo_back_token = "menu" if back_token == "ramais_excluidos" and not RAMAIS_EXCLUIDOS else back_token
        await query.edit_message_text(
            "↩️ Ramal restaurado.\n\n" + build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, novo_back_token, ping_status, user_id),
        )
        return

    if data.startswith("criarramal_predio:"):
        code = data.split(":", 1)[1]
        criando = context.user_data.get("criando_ramal")
        user_id = update.effective_user.id if update.effective_user else None
        if not criando or not is_admin(user_id) or code not in PREDIO_CODES:
            await _responder_query_seguro(query, "Sessão de criação expirada. Toque em ➕ Criar novo ramal de novo.", show_alert=True)
            return
        nome_predio, _label = PREDIO_CODES[code]
        criando["dados"]["predio"] = nome_predio
        criando["dados"]["building"] = nome_predio
        criando["etapa"] = "andar"
        await query.edit_message_text(
            f"Prédio selecionado: <b>{nome_predio}</b>\n\n"
            "Digite o andar (ex.: 1º, 2º, Térreo).\n"
            "(Envie <code>cancelar</code> a qualquer momento para desistir.)",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("pr:"):
        code = data.split(":", 1)[1]
        if code not in PREDIO_CODES:
            return
        _nome, label = PREDIO_CODES[code]
        await query.edit_message_text(
            f"{label}\n\nEscolha um andar:",
            parse_mode=ParseMode.HTML,
            reply_markup=build_andar_keyboard(code),
        )
        return

    if data.startswith("an:"):
        _, code, andar = data.split(":", 2)
        if code not in PREDIO_CODES:
            return
        await query.edit_message_text(
            build_setores_texto(code, andar),
            parse_mode=ParseMode.HTML,
            reply_markup=build_setores_keyboard(code, andar),
        )
        return

    if data.startswith("det:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        name, ramal, _building, _predio, _andar = DATA[idx]
        devices = get_effective_devices(name, ramal)
        devices_to_test = [(i, dev) for i, dev in enumerate(devices) if (dev.get("ip") or dev.get("mac")) and not dev.get("excluido")]

        ping_status = context.user_data.setdefault("ping_status", {})
        user_id = update.effective_user.id if update.effective_user else None

        if devices_to_test:
            await query.edit_message_text(
                build_detail_text(idx, ping_status) + "\n\n🔄 <i>Verificando conectividade...</i>",
                parse_mode=ParseMode.HTML,
            )
            resultados = await asyncio.gather(*[resolve_and_ping(dev) for _, dev in devices_to_test])
            for (i, _dev), ok in zip(devices_to_test, resultados):
                ping_status[f"{idx}:{i}"] = ok
        else:
            await _responder_query_seguro(query)

        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    if data.startswith("ping:"):
        _, idx_str, dev_str, back_token = data.split(":", 3)
        idx = int(idx_str)
        dev_i = int(dev_str)
        name, ramal, _building, _predio, _andar = DATA[idx]
        devices = get_effective_devices(name, ramal)
        if dev_i >= len(devices) or devices[dev_i].get("excluido") or not (devices[dev_i].get("ip") or devices[dev_i].get("mac")):
            await _responder_query_seguro(query, "IP não cadastrado para este aparelho.", show_alert=True)
            return

        await _responder_query_seguro(query, "Testando ping...")
        ok = await resolve_and_ping(devices[dev_i])

        ping_status = context.user_data.setdefault("ping_status", {})
        ping_status[f"{idx}:{dev_i}"] = ok
        user_id = update.effective_user.id if update.effective_user else None

        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    if data.startswith("togcall:"):
        _, idx_str, dev_str, field, back_token = data.split(":", 4)
        idx = int(idx_str)
        dev_i = int(dev_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem editar.", show_alert=True)
            return

        name, ramal, _building, _predio, _andar = DATA[idx]
        devices = get_effective_devices(name, ramal)
        if dev_i >= len(devices):
            await _responder_query_seguro(query, "Aparelho não encontrado.", show_alert=True)
            return

        atual = devices[dev_i].get(field, True)
        novo = not atual
        key = _override_key(name, ramal, dev_i)
        DEVICE_OVERRIDES.setdefault(key, {})[field] = novo
        save_device_overrides()

        ping_status = context.user_data.get("ping_status", {})
        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    if data.startswith("excluir_aparelho_confirma:"):
        _, idx_str, dev_str, back_token = data.split(":", 3)
        idx = int(idx_str)
        dev_i = int(dev_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem excluir aparelhos.", show_alert=True)
            return

        name, ramal, _building, _predio, _andar = DATA[idx]
        devices = get_effective_devices(name, ramal)
        if dev_i >= len(devices) or devices[dev_i].get("excluido"):
            await _responder_query_seguro(query, "Aparelho não encontrado.", show_alert=True)
            return

        rotulo = f"Aparelho {dev_i + 1}" if len(devices) > 1 else "este aparelho"
        pat = devices[dev_i].get("pat") or "sem PAT cadastrado"
        await query.edit_message_text(
            f"⚠️ Tem certeza que quer excluir <b>{rotulo}</b> (PAT: {pat}) deste ramal?\n"
            "Isso não apaga o histórico — dá pra restaurar depois, se precisar.",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("🗑️ Sim, excluir", callback_data=f"excluir_aparelho_ok:{idx}:{dev_i}:{back_token}")],
                    [InlineKeyboardButton("❌ Cancelar", callback_data=f"det:{idx}:{back_token}")],
                ]
            ),
        )
        return

    if data.startswith("excluir_aparelho_ok:"):
        _, idx_str, dev_str, back_token = data.split(":", 3)
        idx = int(idx_str)
        dev_i = int(dev_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem excluir aparelhos.", show_alert=True)
            return

        name, ramal, _building, _predio, _andar = DATA[idx]
        devices = get_effective_devices(name, ramal)
        if dev_i >= len(devices):
            await _responder_query_seguro(query, "Aparelho não encontrado.", show_alert=True)
            return

        key = _override_key(name, ramal, dev_i)
        DEVICE_OVERRIDES.setdefault(key, {})["excluido"] = True
        save_device_overrides()

        ping_status = context.user_data.get("ping_status", {})
        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    if data.startswith("restaurar_aparelho:"):
        _, idx_str, dev_str, back_token = data.split(":", 3)
        idx = int(idx_str)
        dev_i = int(dev_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem restaurar aparelhos.", show_alert=True)
            return

        name, ramal, _building, _predio, _andar = DATA[idx]
        key = _override_key(name, ramal, dev_i)
        if key in DEVICE_OVERRIDES:
            DEVICE_OVERRIDES[key].pop("excluido", None)
            save_device_overrides()

        ping_status = context.user_data.get("ping_status", {})
        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    if data.startswith("edit:"):
        _, idx_str, dev_str, back_token = data.split(":", 3)
        idx = int(idx_str)
        dev_i = int(dev_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem editar.", show_alert=True)
            return
        await query.edit_message_text(
            "O que deseja editar?",
            reply_markup=build_edit_fields_keyboard(idx, dev_i, back_token),
        )
        return

    if data.startswith("editnome:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem editar.", show_alert=True)
            return
        context.user_data["pending_nome"] = {"idx": idx, "back": back_token}
        nome_atual = nome_exibido(idx)
        await query.edit_message_text(
            f"Digite o novo nome para o setor (atual: <b>{nome_atual}</b>).\n"
            "Envie <code>-</code> (hífen) para voltar ao nome original.",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("editramal:"):
        _, idx_str, back_token = data.split(":", 2)
        idx = int(idx_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem editar.", show_alert=True)
            return
        context.user_data["pending_ramal"] = {"idx": idx, "back": back_token}
        ramal_atual = ramal_exibido(idx)
        await query.edit_message_text(
            f"Digite o novo número de ramal para este setor (atual: <b>{ramal_atual}</b>).\n"
            "Isso só muda o número exibido e usado na busca — o cadastro do aparelho "
            "(PAT/MAC/IP/N° Tronco) não é afetado.\n"
            "Envie <code>-</code> (hífen) para voltar ao número original.",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("campo:"):
        _, idx_str, dev_str, field, back_token = data.split(":", 4)
        idx = int(idx_str)
        dev_i = int(dev_str)
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            await _responder_query_seguro(query, "Apenas administradores podem editar.", show_alert=True)
            return
        context.user_data["pending_edit"] = {
            "idx": idx,
            "dev": dev_i,
            "field": field,
            "back": back_token,
        }
        field_label = {"pat": "PAT", "n_serie": "N° Série", "mac": "MAC", "ip": "IP", "tronco": "N° Tronco", "aux": "N° Auxiliar"}.get(field, field.upper())
        await query.edit_message_text(
            f"Digite o novo valor de <b>{field_label}</b> para este aparelho.\n"
            "Envie <code>-</code> (hífen) para limpar o valor.",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "voltar_busca":
        term = context.user_data.get("last_search_term")
        if not term:
            await query.edit_message_text(
                "Faça uma nova busca digitando o nome de um setor.",
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return
        idxs = search_indices(term)
        if not idxs:
            await query.edit_message_text(
                format_results(term, []),
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return
        await query.edit_message_text(
            f"Resultados para \"{term}\":\n\nToque em um setor para ver o ramal:",
            parse_mode=ParseMode.HTML,
            reply_markup=build_result_keyboard(idxs, "voltar_busca"),
        )
        return


async def photo_handler(update: Update, context) -> None:
    """Recebe uma foto (ex.: etiqueta de patrimônio), tenta ler um código
    de barras nela e mostra em qual ramal aquele PAT está cadastrado."""
    _reiniciar_timer_inatividade(update, context)
    user_id = update.effective_user.id if update.effective_user else None

    if not LEITURA_CODIGO_BARRAS_DISPONIVEL:
        await update.message.reply_text(
            "⚠️ A leitura de código de barras não está disponível neste bot ainda.\n"
            "Falta instalar as bibliotecas Pillow e pyzbar "
            "(<code>pip install Pillow pyzbar</code>) e reiniciar o bot.\n\n"
            "Por enquanto, use o botão 🏷️ Buscar por patrimônio (PAT) e digite o número.",
            parse_mode=ParseMode.HTML,
        )
        return

    aviso = await update.message.reply_text("🔎 Lendo o código de barras da foto...")

    try:
        foto = update.message.photo[-1]
        arquivo = await foto.get_file()
        dados_bytes = await arquivo.download_as_bytearray()
        imagem = Image.open(io.BytesIO(bytes(dados_bytes)))
        resultados = decode_barcodes(imagem)
    except Exception:
        logger.exception("Falha ao processar a foto para leitura de código de barras.")
        await aviso.delete()
        await update.message.reply_text(
            "❌ Não consegui processar essa foto. Tenta mandar de novo."
        )
        return

    await aviso.delete()

    if not resultados:
        await update.message.reply_text(
            "❌ Não consegui ler nenhum código de barras nessa foto.\n"
            "Tenta tirar de novo bem de frente, com boa luz e sem reflexo — "
            "ou use o botão 🏷️ Buscar por patrimônio (PAT) e digite o número manualmente."
        )
        return

    numero = resultados[0].data.decode("utf-8", errors="ignore").strip()
    numero_digitos = re.sub(r"\D", "", numero)

    if not numero_digitos:
        await update.message.reply_text(
            f"Li o código <code>{numero}</code>, mas ele não parece ser um número de PAT.\n"
            "Use o botão 🏷️ Buscar por patrimônio (PAT) pra buscar manualmente.",
            parse_mode=ParseMode.HTML,
        )
        return

    idxs = encontrar_por_pat_exato(numero_digitos)

    if not idxs:
        await update.message.reply_text(
            f"🔎 Código lido: <code>{numero_digitos}</code>\n\n"
            "Não encontrei nenhum aparelho cadastrado com esse PAT.\n"
            "Confere se é esse mesmo o número, ou cadastra/edita o PAT no ramal certo primeiro.",
            parse_mode=ParseMode.HTML,
        )
        return

    if len(idxs) > 1:
        await update.message.reply_text(
            f"🔎 Código lido: <code>{numero_digitos}</code>\n\n"
            "Mais de um ramal está com esse PAT cadastrado:",
            parse_mode=ParseMode.HTML,
            reply_markup=build_result_keyboard(idxs, "voltar_busca"),
        )
        return

    idx = idxs[0]
    await update.message.reply_text(
        f"🔎 PAT <code>{numero_digitos}</code> encontrado:",
        parse_mode=ParseMode.HTML,
    )
    await update.message.reply_text(
        build_detail_text(idx),
        parse_mode=ParseMode.HTML,
        reply_markup=build_detail_keyboard(idx, "menu", {}, user_id),
    )


async def text_handler(update: Update, context) -> None:
    _reiniciar_timer_inatividade(update, context)

    pending = context.user_data.get("pending_edit")
    if pending:
        del context.user_data["pending_edit"]
        idx = pending["idx"]
        dev_i = pending["dev"]
        field = pending["field"]
        back_token = pending["back"]
        user_id = update.effective_user.id if update.effective_user else None

        if not is_admin(user_id):
            await update.message.reply_text("Apenas administradores podem editar.")
            return

        name, ramal, *_ = DATA[idx]
        valor = update.message.text.strip()
        valor_final = None if valor == "-" else valor

        key = _override_key(name, ramal, dev_i)
        DEVICE_OVERRIDES.setdefault(key, {})[field] = valor_final
        save_device_overrides()

        ping_status = context.user_data.get("ping_status", {})
        await update.message.reply_text("✅ Valor atualizado.")
        await update.message.reply_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    pending_nome = context.user_data.pop("pending_nome", None)
    if pending_nome:
        idx = pending_nome["idx"]
        back_token = pending_nome["back"]
        user_id = update.effective_user.id if update.effective_user else None

        if not is_admin(user_id):
            await update.message.reply_text("Apenas administradores podem editar.")
            return

        valor = update.message.text.strip()
        if valor == "-":
            NOME_OVERRIDES.pop(idx, None)
        else:
            NOME_OVERRIDES[idx] = valor
        save_nome_overrides()

        ping_status = context.user_data.get("ping_status", {})
        await update.message.reply_text("✅ Nome atualizado.")
        await update.message.reply_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    pending_ramal = context.user_data.pop("pending_ramal", None)
    if pending_ramal:
        idx = pending_ramal["idx"]
        back_token = pending_ramal["back"]
        user_id = update.effective_user.id if update.effective_user else None

        if not is_admin(user_id):
            await update.message.reply_text("Apenas administradores podem editar.")
            return

        valor = update.message.text.strip()
        if valor == "-":
            RAMAL_OVERRIDES.pop(idx, None)
        else:
            RAMAL_OVERRIDES[idx] = valor
        save_ramal_overrides()
        _recalcular_ramais_compartilhados()

        ping_status = context.user_data.get("ping_status", {})
        await update.message.reply_text("✅ Número de ramal atualizado.")
        await update.message.reply_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status, user_id),
        )
        return

    criando = context.user_data.get("criando_ramal")
    if criando:
        user_id = update.effective_user.id if update.effective_user else None
        if not is_admin(user_id):
            del context.user_data["criando_ramal"]
            await update.message.reply_text("Apenas administradores podem criar ramais.")
            return

        etapa = criando["etapa"]
        texto = update.message.text.strip()

        if texto.lower() in ("cancelar", "/cancelar", "-"):
            del context.user_data["criando_ramal"]
            await update.message.reply_text(
                "❌ Criação de ramal cancelada. Nada foi salvo.",
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return

        if etapa == "nome":
            if not texto:
                await update.message.reply_text("Nome inválido. Digite o nome do setor:")
                return
            criando["dados"]["name"] = texto
            criando["etapa"] = "ramal"
            await update.message.reply_text(
                "Digite o número do ramal.\n(Envie <code>cancelar</code> para desistir.)",
                parse_mode=ParseMode.HTML,
            )
            return

        if etapa == "ramal":
            if not texto:
                await update.message.reply_text("Ramal inválido. Digite o número do ramal:")
                return
            criando["dados"]["ramal"] = texto
            criando["etapa"] = "predio"
            botoes = [
                [InlineKeyboardButton(label, callback_data=f"criarramal_predio:{code}")]
                for code, (_nome, label) in PREDIO_CODES.items()
            ]
            botoes.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_criar_ramal")])
            await update.message.reply_text(
                "Escolha o prédio:", reply_markup=InlineKeyboardMarkup(botoes)
            )
            return

        if etapa == "predio":
            await update.message.reply_text("Toque em um dos botões de prédio na mensagem acima pra continuar (ou em ❌ Cancelar).")
            return

        if etapa == "andar":
            if not texto:
                await update.message.reply_text("Andar inválido. Digite o andar (ex.: 1º, 2º, Térreo):")
                return
            dados = criando["dados"]
            dados["andar"] = texto
            del context.user_data["criando_ramal"]

            nova_tupla = (dados["name"], dados["ramal"], dados["building"], dados["predio"], dados["andar"])
            DATA.append(nova_tupla)
            salvar_ramal_criado(*nova_tupla)
            _recalcular_ramais_compartilhados()

            novo_idx = len(DATA) - 1
            await update.message.reply_text("✅ Ramal criado com sucesso!")
            await update.message.reply_text(
                build_detail_text(novo_idx),
                parse_mode=ParseMode.HTML,
                reply_markup=build_detail_keyboard(novo_idx, "menu", {}, user_id),
            )
            return

    if context.user_data.pop("aguardando_busca_pat", False):
        term = update.message.text
        idxs = search_indices_by_pat(term)
        context.user_data["last_search_term"] = term

        if not idxs:
            await update.message.reply_text(
                f'Nenhum ramal encontrado com o PAT "{term}".\n'
                "Tente outro número de patrimônio.",
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return

        await update.message.reply_text(
            f'Resultados para o PAT "{term}":\n\nToque em um setor para ver o ramal:',
            parse_mode=ParseMode.HTML,
            reply_markup=build_result_keyboard(idxs, "voltar_busca"),
        )
        return

    if context.user_data.pop("aguardando_busca_ramal", False):
        term = update.message.text
        idxs = search_indices_by_ramal(term)
        context.user_data["last_search_term"] = term

        if not idxs:
            await update.message.reply_text(
                f'Nenhum ramal encontrado com o número "{term}".\n'
                "Tente outro número de ramal.",
                reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
            )
            return

        await update.message.reply_text(
            f'Resultados para o ramal "{term}":\n\nToque em um setor para ver o ramal:',
            parse_mode=ParseMode.HTML,
            reply_markup=build_result_keyboard(idxs, "voltar_busca"),
        )
        return

    term = update.message.text
    idxs = search_indices(term)
    context.user_data["last_search_term"] = term

    if not idxs:
        await update.message.reply_text(
            format_results(term, []),
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]),
        )
        return

    await update.message.reply_text(
        f"Resultados para \"{term}\":\n\nToque em um setor para ver o ramal:",
        parse_mode=ParseMode.HTML,
        reply_markup=build_result_keyboard(idxs, "voltar_busca"),
    )


async def inline_query(update: Update, context) -> None:
    query = update.inline_query.query
    if not query:
        return
    idxs = search_indices(query, limit=15)
    articles = []
    for idx in idxs:
        name, ramal, building, predio, andar = DATA[idx]
        nome_atual = nome_exibido(idx)
        local = f"{predio} — {andar} andar" if predio != "?" else "local não confirmado"
        rc = ramal_com_fila(name, ramal_exibido(idx))
        articles.append(
            InlineQueryResultArticle(
                id=str(uuid4()),
                title=f"{nome_atual} — {rc}",
                description=local,
                input_message_content=InputTextMessageContent(
                    f"📞 <b>{nome_atual}</b>\nRamal: <b>{rc}</b> ({local})",
                    parse_mode=ParseMode.HTML,
                ),
            )
        )
    await update.inline_query.answer(articles, cache_time=10)


async def erro_handler(update, context) -> None:
    """Handler global de erros: registra o problema no log sem derrubar o
    bot. Evita o aviso 'No error handlers are registered' e trata erros
    conhecidos e inofensivos (ex.: clique em botão expirado) de forma mais
    silenciosa."""
    erro = context.error
    if isinstance(erro, TelegramError) and "too old" in str(erro).lower():
        logger.debug("Callback query expirado, ignorando: %s", erro)
        return
    logger.exception("Erro não tratado ao processar um update.", exc_info=erro)


def main() -> None:
    # Token do bot direto no arquivo (obtido no @BotFather).
    # Usa sempre este valor — propositalmente NÃO lê mais a variável de
    # ambiente TELEGRAM_BOT_TOKEN, porque uma variável antiga/errada nesta
    # máquina ficou sobrescrevendo o token correto e derrubando o bot com
    # "InvalidToken" mesmo depois do valor abaixo já estar certo.
    TOKEN_PADRAO = "8721738671:AAFDCAQYQgOvvlvEH4sYEa23yvoSlrPmhJ0"
    token = TOKEN_PADRAO
    if not token:
        raise SystemExit(
            "Defina a variável de ambiente TELEGRAM_BOT_TOKEN com o token do seu bot "
            "(obtido no @BotFather) antes de rodar este script."
        )

    load_device_overrides()
    load_nome_overrides()
    load_ramal_overrides()
    load_ramais_excluidos()
    load_ramais_criados()
    load_inventario()
    # Recalcula "Fila compartilhada" considerando os números de ramal já
    # editados por admin e os ramais excluídos (load_ramais_criados também
    # chama isso, mas só quando há ramais criados pelo bot — aqui garante
    # que RAMAL_OVERRIDES/RAMAIS_EXCLUIDOS sempre sejam levados em conta,
    # mesmo sem nenhum ramal criado).
    _recalcular_ramais_compartilhados()

    if LEITURA_CODIGO_BARRAS_DISPONIVEL:
        logger.info("Leitura de código de barras (Pillow + pyzbar) disponível.")
    else:
        logger.warning(
            "Pillow/pyzbar não instalados — leitura de código de barras por foto desativada. "
            "Rode 'pip install -r requirements.txt' pra ativar."
        )

    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("ajuda", ajuda))
    app.add_handler(CommandHandler("meuid", meuid))
    app.add_handler(CommandHandler("chatid", chatid))
    app.add_handler(CommandHandler("statusagora", statusagora))
    app.add_handler(CommandHandler("listar", listar))
    app.add_handler(CommandHandler("menu", menu))
    app.add_handler(CallbackQueryHandler(menu_callback, pattern=r"^(menu|pr:|an:|det:|ping:|edit:|campo:|togcall:|editnome:|editramal:|excluir_aparelho_confirma:|excluir_aparelho_ok:|restaurar_aparelho:|ramais_excluidos|excluir_ramal_confirma:|excluir_ramal_ok:|restaurar_ramal:|listar_todos|listar_por_local|listar_offline|buscar_pat|buscar_ramal|criar_ramal|cancelar_criar_ramal|criarramal_predio:|exportar_lista|checklist_inventario|check_inventario:|zerar_inventario_confirma|zerar_inventario_ok|voltar_busca)"))
    app.add_handler(MessageHandler(filters.PHOTO, photo_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    app.add_handler(InlineQueryHandler(inline_query))
    app.add_error_handler(erro_handler)

    logger.info("Bot iniciado. Aguardando mensagens...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()