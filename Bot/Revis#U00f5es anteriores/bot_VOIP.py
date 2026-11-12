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
import logging
import os
import platform
import re
import subprocess
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
    ("Gerência Assistencial", "8314", "HU", "HU", "2º"),
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
    ("Supervisão de Enfermagem (Materno/Infantil)", "8553", "HU", "HU", "4º"),
    ("Supervisão de Hotelaria (3º Piso)", "8491", "HU", "HU", "3º"),
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
    """Lista códigos de prédio que realmente têm setores cadastrados, na ordem de PREDIO_CODES."""
    presentes = {item[3] for item in DATA}
    return [code for code, (nome, _label) in PREDIO_CODES.items() if nome in presentes]


def andares_do_predio(predio_nome: str):
    andares = {item[4] for item in DATA if item[3] == predio_nome}
    return sorted(andares, key=_andar_sort_key)


def setores_do_predio_andar(predio_nome: str, andar: str):
    itens = [item for item in DATA if item[3] == predio_nome and item[4] == andar]
    return sorted(itens, key=lambda x: normalize(x[0]))


def _menu_button_row() -> list:
    return [InlineKeyboardButton("🏠 Menu Inicial", callback_data="menu")]


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


def build_predio_keyboard() -> "InlineKeyboardMarkup":
    botoes = []
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
    idx_by_name = {item: i for i, item in enumerate(DATA)}
    botoes = []
    for item in itens:
        idx = idx_by_name[item]
        botoes.append([_sector_button(idx, back_token)])
    botoes.append([InlineKeyboardButton("🔙 Voltar aos andares", callback_data=f"pr:{code}")])
    botoes.append(_menu_button_row())
    return InlineKeyboardMarkup(botoes)


def build_full_listing_chunks() -> list:
    """Monta a lista completa de ramais agrupada por letra, dividida em blocos
    que respeitam o limite de tamanho de mensagem do Telegram."""
    grouped = {}
    for name, ramal, building, predio, andar in sorted(DATA, key=lambda x: normalize(x[0])):
        letter = _strip_accents(name[0]).upper()
        grouped.setdefault(letter, []).append((name, ramal, building, predio, andar))

    chunks = []
    current = "📋 <b>Lista completa de ramais</b>\n\n"
    for letter in sorted(grouped.keys()):
        section = f"<b>— {letter} —</b>\n"
        for name, ramal, building, predio, andar in grouped[letter]:
            local = f"{predio} — {andar} andar" if predio != "?" else "local não confirmado"
            section += f"• {name} — <b>{ramal_com_fila(name, ramal)}</b> ({local})\n"
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
    """Retorna o texto do ramal, acrescentando a fila ao lado quando existir."""
    fila = FILA_INFO.get((name, ramal))
    if fila:
        return f"{ramal} (Fila {fila})"
    if ramal in _RAMAIS_COMPARTILHADOS:
        return f"{ramal} (Fila compartilhada)"
    return ramal


def search_indices(term: str, limit: int = MAX_RESULTS):
    """Como search(), mas retorna os índices em DATA (para montar botões)."""
    term_n = normalize(term)
    if not term_n:
        return []
    idxs = [
        i
        for i, (name, ramal, building, predio, andar) in enumerate(DATA)
        if term_n in normalize(name) or term_n in ramal
    ]
    return idxs[:limit]


def search(term: str, limit: int = MAX_RESULTS):
    """Busca setores por nome (ignorando acentos/caixa) ou por ramal exato/parcial."""
    return [DATA[i] for i in search_indices(term, limit)]


def format_results(term: str, results) -> str:
    if not results:
        return (
            f"Nenhum ramal encontrado para \"{term}\".\n"
            "Tente outro termo (ex: parte do nome do setor ou o número do ramal)."
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
    label = f"{name} — {ramal_com_fila(name, ramal)}"
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
    texto = (
        f"📞 <b>{name}</b>\n\n"
        f"Ramal: <b>{ramal_com_fila(name, ramal)}</b>\n"
    )
    if predio != "?":
        texto += f"Local: {predio}\nAndar: {andar}\n"
    else:
        texto += "Local: não confirmado\n"

    devices = DEVICE_INFO.get((name, ramal), [])
    if not devices:
        texto += "\n☎️ Equipamento: <i>não cadastrado no inventário</i>\n"
    else:
        texto += "\n☎️ <b>Equipamento</b>"
        if len(devices) > 1:
            texto += f" ({len(devices)} aparelhos nesta linha)"
        texto += "\n"
        for i, dev in enumerate(devices, start=1):
            pat = dev["pat"] or "—"
            n_serie = dev["n_serie"] or "—"
            mac = dev["mac"] or "—"
            ip = dev["ip"] or "—"
            badge = _ping_badge(ping_status.get(f"{idx}:{i - 1}")) if dev["ip"] else ""
            if len(devices) > 1:
                texto += f"\n<b>Aparelho {i}</b>{badge}\n"
            texto += (
                f"PAT: <code>{pat}</code>\n"
                f"N° Série: <code>{n_serie}</code>\n"
                f"MAC: <code>{mac}</code>\n"
                f"IP: <code>{ip}</code>{badge}\n"
            )

    texto += f"\nTelefone Geral: {GENERAL_PHONE}"
    return texto


def build_detail_keyboard(idx: int, back_token: str, ping_status: dict | None = None) -> "InlineKeyboardMarkup":
    name, ramal, _building, _predio, _andar = DATA[idx]
    devices = DEVICE_INFO.get((name, ramal), [])
    ping_status = ping_status or {}

    botoes = []
    for i, dev in enumerate(devices):
        ip = dev["ip"]
        if not ip:
            continue
        status = ping_status.get(f"{idx}:{i}")
        prefixo = f"Aparelho {i + 1}: " if len(devices) > 1 else ""
        if status is True:
            label = f"🟢 {prefixo}{ip} — Online (testar de novo)"
        elif status is False:
            label = f"🔴 {prefixo}{ip} — Sem resposta (testar de novo)"
        else:
            label = f"📡 Testar ping — {prefixo}{ip}"
        if len(label) > 64:
            label = label[:61] + "..."
        botoes.append([InlineKeyboardButton(label, callback_data=f"ping:{idx}:{i}:{back_token}")])

    botoes.append([InlineKeyboardButton("🔙 Voltar", callback_data=back_token)])
    botoes.append(_menu_button_row())
    return InlineKeyboardMarkup(botoes)


# ── Handlers ──────────────────────────────────────────────────────────────

async def start(update: Update, context) -> None:
    await update.message.reply_text(
        "📞 <b>Guia de Ramais VOIP</b>\n"
        "Hospital Universitário São Francisco de Paula\n\n"
        f"Telefone Geral: {GENERAL_PHONE}\n\n"
        "/listar — mostra todos os ramais cadastrados\n"
        "/menu — navega por prédio e andar\n\n"
        "Ou escolha um prédio abaixo:",
        parse_mode=ParseMode.HTML,
        reply_markup=build_predio_keyboard(),
    )


async def ajuda(update: Update, context) -> None:
    await start(update, context)


async def listar(update: Update, context) -> None:
    chunks = build_full_listing_chunks()
    for i, chunk in enumerate(chunks, start=1):
        footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
        is_last = i == len(chunks)
        await update.message.reply_text(
            chunk + footer,
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([_menu_button_row()]) if is_last else None,
        )


async def menu(update: Update, context) -> None:
    await update.message.reply_text(
        "🏢 <b>Navegar por prédio e andar</b>\n\nEscolha um prédio:",
        parse_mode=ParseMode.HTML,
        reply_markup=build_predio_keyboard(),
    )


async def menu_callback(update: Update, context) -> None:
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "menu":
        await query.edit_message_text(
            "🏢 <b>Navegar por prédio e andar</b>\n\nEscolha um prédio:",
            parse_mode=ParseMode.HTML,
            reply_markup=build_predio_keyboard(),
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
        devices = DEVICE_INFO.get((name, ramal), [])
        ips_to_test = [(i, dev["ip"]) for i, dev in enumerate(devices) if dev["ip"]]

        ping_status = context.user_data.setdefault("ping_status", {})

        if ips_to_test:
            await query.edit_message_text(
                build_detail_text(idx, ping_status) + "\n\n🔄 <i>Verificando conectividade...</i>",
                parse_mode=ParseMode.HTML,
            )
            resultados = await asyncio.gather(*[ping_host(ip) for _, ip in ips_to_test])
            for (i, _ip), ok in zip(ips_to_test, resultados):
                ping_status[f"{idx}:{i}"] = ok
        else:
            await query.answer()

        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status),
        )
        return

    if data.startswith("ping:"):
        _, idx_str, dev_str, back_token = data.split(":", 3)
        idx = int(idx_str)
        dev_i = int(dev_str)
        name, ramal, _building, _predio, _andar = DATA[idx]
        devices = DEVICE_INFO.get((name, ramal), [])
        if dev_i >= len(devices) or not devices[dev_i]["ip"]:
            await query.answer("IP não cadastrado para este aparelho.", show_alert=True)
            return

        await query.answer("Testando ping...")
        ok = await ping_host(devices[dev_i]["ip"])

        ping_status = context.user_data.setdefault("ping_status", {})
        ping_status[f"{idx}:{dev_i}"] = ok

        await query.edit_message_text(
            build_detail_text(idx, ping_status),
            parse_mode=ParseMode.HTML,
            reply_markup=build_detail_keyboard(idx, back_token, ping_status),
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


async def text_handler(update: Update, context) -> None:
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
    results = search(query, limit=15)
    articles = []
    for name, ramal, building, predio, andar in results:
        local = f"{predio} — {andar} andar" if predio != "?" else "local não confirmado"
        rc = ramal_com_fila(name, ramal)
        articles.append(
            InlineQueryResultArticle(
                id=str(uuid4()),
                title=f"{name} — {rc}",
                description=local,
                input_message_content=InputTextMessageContent(
                    f"📞 <b>{name}</b>\nRamal: <b>{rc}</b> ({local})",
                    parse_mode=ParseMode.HTML,
                ),
            )
        )
    await update.inline_query.answer(articles, cache_time=10)


def main() -> None:
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise SystemExit(
            "Defina a variável de ambiente TELEGRAM_BOT_TOKEN com o token do seu bot "
            "(obtido no @BotFather) antes de rodar este script."
        )

    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("ajuda", ajuda))
    app.add_handler(CommandHandler("listar", listar))
    app.add_handler(CommandHandler("menu", menu))
    app.add_handler(CallbackQueryHandler(menu_callback, pattern=r"^(menu|pr:|an:|det:|ping:|voltar_busca)"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    app.add_handler(InlineQueryHandler(inline_query))

    logger.info("Bot iniciado. Aguardando mensagens...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()