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

import logging
import os
import unicodedata
from uuid import uuid4

from telegram import (
    InlineQueryResultArticle,
    InputTextMessageContent,
    Update,
)
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
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
# (Nome do setor, Ramal, Prédio/Unidade)
DATA = [
    ("Almoxarifado (Central de Materiais)", "8560", "HU"),
    ("Almoxarifado (Recebimentos)", "8560", "HU"),
    ("Bloco Cirúrgico (Farmácia)", "8350", "HU"),
    ("Bloco Cirúrgico (Marcação)", "8376", "HU"),
    ("Bloco Cirúrgico (Posto de Enfermagem)", "8591", "HU"),
    ("Bloco Cirúrgico (Sala de Recuperação)", "8347", "HU"),
    ("CAF", "8331", "HU"),
    ("Casa da Gestante", "8360", "HU"),
    ("CDI - Administrativo", "8496", "HU"),
    ("CDI (Convênios)", "8310", "HU"),
    ("CDI (Sala 01)", "8449", "HU"),
    ("CDI (Sala 02)", "8563", "HU"),
    ("CDI (Sala 03)", "8380", "HU"),
    ("CDI (SUS)", "8318", "HU"),
    ("Central de Cadastro - MV", "8507", "ADM"),
    ("Central de Cadastro SIGH", "8506", "ADM"),
    ("Centro Acadêmico (Recepção)", "8386", "HU"),
    ("Clínica Cirúrgica (Escriturários)", "8557", "HU"),
    ("Clínica Cirúrgica (Posto de Enfermagem)", "8320", "HU"),
    ("CME", "8336", "HU"),
    ("Comunicação e Marketing", "8368", "PA"),
    ("Compras", "8326", "HU"),
    ("Conceição (Posto de Enfermagem)", "8575", "HU"),
    ("Conceição (Sala de Prescrição)", "8324", "HU"),
    ("Conforto Médico (Clínica Médica)", "8357", "HU"),
    ("Conforto Médico (Ginecologia)", "8361", "HU"),
    ("Conforto Pediatria", "8367", "HU"),
    ("Conforto UTI Neo/Ped", "8585", "HU"),
    ("Consultório 1 Multiprofissional", "8469", "HU"),
    ("Consultório 2 Nefrologia", "8426", "HU"),
    ("Consultório 3 Multiprofissional", "8427", "HU"),
    ("Consultório Infantil (PA)", "8594", "PA"),
    ("Contabilidade", "8504", "ADM"),
    ("Contas Médicas", "8503", "ADM"),
    ("COREME", "8341", "DIR"),
    ("Dietoterapia", "8333", "HU"),
    ("Direção (Recepção)", "8358", "DIR"),
    ("Eletrocardiograma", "8311", "HU"),
    ("Endoscopia (Sala 02)", "8348", "HU"),
    ("Endoscopia Administrativo", "8348", "HU"),
    ("Engenharia Clínica", "8390", "HU"),
    ("Escriturários (4º Piso)", "8329", "HU"),
    ("Farmácia (4º Piso)", "8332", "HU"),
    ("Farmácia de Manipulação", "8343", "HU"),
    ("Fátima (Posto de Enfermagem)", "8302", "HU"),
    ("Fátima (Supervisão)", "8596", "HU"),
    ("Faturamento (Convênios)", "8321", "ADM"),
    ("Faturamento SUS", "8374", "ADM"),
    ("Financeiro (ADM)", "8309", "ADM"),
    ("Financeiro / Caixa (HU)", "8379", "HU"),
    ("Fisioterapia", "8593", "HU"),
    ("Gerência Administrativa", "8559", "ADM"),
    ("Gerência Assistencial", "8314", "HU"),
    ("Gerência de Hotelaria", "8485", "HU"),
    ("Gerência de Suprimentos", "8487", "HU"),
    ("Gestão de Pessoas", "8344", "ADM"),
    ("Guadalupe (Posto de Enfermagem)", "8552", "HU"),
    ("Hemodiálise (Administrativo)", "8364", "HU"),
    ("Hemodiálise (Supervisão)", "8562", "HU"),
    ("Hemodinâmica", "8558", "HU"),
    ("Higienização", "8369", "HU"),
    ("Internação (Convênios)", "8371", "HU"),
    ("Internação (Supervisão)", "8599", "HU"),
    ("Internação SUS", "8391", "HU"),
    ("Jurídico", "8349", "ADM"),
    ("Lavanderia", "8362", "HU"),
    ("LEAC (Administrativo)", "8359", "HU"),
    ("LEAC (Agência Transfusional)", "8330", "HU"),
    ("LEAC (Bioquímica)", "8572", "HU"),
    ("LEAC (Convênios)", "8573", "HU"),
    ("LEAC (Gerência)", "8385", "HU"),
    ("LEAC (Hematologia)", "8572", "HU"),
    ("LEAC (Microbiologia)", "8572", "HU"),
    ("LEAC (Supervisão Administrativa)", "8307", "HU"),
    ("LEAC (Triagem)", "8551", "HU"),
    ("LEAC SUS", "8573", "HU"),
    ("Manutenção", "8339", "MAN"),
    ("Maternidade (Convênios)", "8356", "HU"),
    ("Maternidade (SUS)", "8322", "HU"),
    ("NIR", "8394", "HU"),
    ("Nutrição (Produção)", "8351", "HU"),
    ("OPME", "8466", "HU"),
    ("Ouvidoria", "8375", "HU"),
    ("PA Convênios Maternidade", "8508", "HU"),
    ("Pediatria (4º Piso)", "8366", "HU"),
    ("Portaria", "8319", "HU"),
    ("Posto de Enfermagem (Clínica Médica)", "8301", "HU"),
    ("Projetos", "8564", "ADM"),
    ("Pronto Atendimento (Recepção)", "8378", "PA"),
    ("Pronto Atendimento (Posto de Enfermagem)", "8353", "PA"),
    ("Pronto Socorro (Acolhimento)", "8592", "PS"),
    ("Pronto Socorro (Cirurgia/Bucomaxilofacial)", "8550", "PS"),
    ("Pronto Socorro (SCIH)", "8568", "PS"),
    ("Pronto Socorro (Coleta Laboratorial)", "8567", "PS"),
    ("Pronto Socorro (Compras / Financeiro)", "8576", "PS"),
    ("Pronto Socorro (Consultório Pediátrico)", "8377", "PS"),
    ("Pronto Socorro (Direção)", "8574", "PS"),
    ("Pronto Socorro (Emergência)", "8304", "PS"),
    ("Pronto Socorro (Escriturários)", "8567", "PS"),
    ("Pronto Socorro (Estabilização Adulto)", "8592", "PS"),
    ("Pronto Socorro (Farmácia)", "8582", "PS"),
    ("Pronto Socorro (Portaria)", "8577", "PS"),
    ("Pronto Socorro (Posto de Enfermagem)", "8588", "PS"),
    ("Pronto Socorro (Posto Enfermagem Pediatria)", "8579", "PS"),
    ("Pronto Socorro (Raio-X)", "8303", "PS"),
    ("Pronto Socorro (Recepção)", "8305", "PS"),
    ("Pronto Socorro (Serviço Social)", "8584", "PS"),
    ("Pronto Socorro (Triagem Médica)", "8586", "PS"),
    ("Prontuários (Direção)", "8589", "HU"),
    ("Psicologia", "8396", "HU"),
    ("Qualidade", "8482", "HU"),
    ("Radiologia", "8397", "HU"),
    ("Raio-X (Laudos)", "8384", "HU"),
    ("Recém Nascido", "8325", "HU"),
    ("Recepção (5º Piso)", "8486", "HU"),
    ("Rede Fisio", "8389", "HU"),
    ("RH", "8345", "ADM"),
    ("RT Farmácia", "8335", "HU"),
    ("Sala de Diálise 1", "8313", "HU"),
    ("Sala de Diálise 3", "8312", "HU"),
    ("Sala de Diálise 4", "8328", "HU"),
    ("Sala de Estudos (3º Piso)", "8398", "HU"),
    ("Sala de Estudos (4º Piso - Conceição)", "8422", "HU"),
    ("Sala de Estudos (5º Piso)", "8569", "HU"),
    ("Sala de Laudos", "8317", "HU"),
    ("São Camilo (Posto de Enfermagem)", "8571", "HU"),
    ("Saúde + Fácil (Agendamento de Consultas)", "8383", "PA"),
    ("Saúde + Fácil (Agendamento Reumato SUS)", "8583", "PA"),
    ("Saúde + Fácil (Recepção)", "8570", "PA"),
    ("Saúde + Fácil (Supervisão)", "8488", "PA"),
    ("Saúde Ocupacional (Administrativo)", "8561", "PA"),
    ("Saúde Ocupacional (Recepção)", "8355", "PA"),
    ("SCIH", "8340", "HU"),
    ("SESMT (Consultório 1)", "8433", "PA"),
    ("SESMT Interno", "8352", "ADM"),
    ("Serviço Social (Guadalupe)", "8388", "HU"),
    ("Supervisão de Enfermagem (Materno/Infantil)", "8553", "HU"),
    ("Supervisão de Hotelaria (3º Piso)", "8491", "HU"),
    ("Supervisão de Restaurante", "8501", "HU"),
    ("Supervisão UTIs", "8497", "HU"),
    ("Telefonia", "8327", "HU"),
    ("TI", "8306", "ADM"),
    ("Tomografia", "8393", "HU"),
    ("UTI Geral (Farmácia)", "8363", "HU"),
    ("UTI Geral (Posto de Enfermagem / Área 1)", "8334", "HU"),
    ("UTI Geral (Médicos)", "8392", "HU"),
    ("UTI Neo Natal", "8365", "HU"),
    ("UTI Pediátrica", "8590", "HU"),
    ("UTSI", "8329", "HU"),
    ("UTU Área 2", "8308", "HU"),
]

GENERAL_PHONE = "(53) 2128-8300"
MAX_RESULTS = 20
TELEGRAM_MSG_LIMIT = 4000  # margem de segurança abaixo do limite real (4096)


def build_full_listing_chunks() -> list:
    """Monta a lista completa de ramais agrupada por letra, dividida em blocos
    que respeitam o limite de tamanho de mensagem do Telegram."""
    grouped = {}
    for name, ramal, building in sorted(DATA, key=lambda x: normalize(x[0])):
        letter = _strip_accents(name[0]).upper()
        grouped.setdefault(letter, []).append((name, ramal, building))

    chunks = []
    current = "📋 <b>Lista completa de ramais</b>\n\n"
    for letter in sorted(grouped.keys()):
        section = f"<b>— {letter} —</b>\n"
        for name, ramal, building in grouped[letter]:
            section += f"• {name} — <b>R. {ramal}</b> ({building})\n"
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


def search(term: str, limit: int = MAX_RESULTS):
    """Busca setores por nome (ignorando acentos/caixa) ou por ramal exato/parcial."""
    term_n = normalize(term)
    if not term_n:
        return []
    results = [
        (name, ramal, building)
        for name, ramal, building in DATA
        if term_n in normalize(name) or term_n in ramal
    ]
    return results[:limit]


def format_results(term: str, results) -> str:
    if not results:
        return (
            f"Nenhum ramal encontrado para \"{term}\".\n"
            "Tente outro termo (ex: parte do nome do setor ou o número do ramal)."
        )
    lines = [f"Resultados para \"{term}\":\n"]
    for name, ramal, building in results:
        lines.append(f"• {name} — <b>R. {ramal}</b> ({building})")
    if len(results) == MAX_RESULTS:
        lines.append("\n(mostrando os primeiros resultados — refine sua busca)")
    return "\n".join(lines)


# ── Handlers ──────────────────────────────────────────────────────────────

async def start(update: Update, context) -> None:
    await update.message.reply_text(
        "📞 <b>Guia de Ramais VOIP</b>\n"
        "Hospital Universitário São Francisco de Paula\n\n"
        f"Telefone Geral: {GENERAL_PHONE}\n\n"
        "Digite o nome de um setor (ex: <i>Farmácia</i>) ou um número de ramal "
        "para buscar. Você também pode usar o modo inline: "
        "<code>@seu_bot nome_do_setor</code> em qualquer conversa.\n\n"
        "/listar — mostra todos os ramais cadastrados\n"
        "/ajuda — mostra esta mensagem novamente",
        parse_mode=ParseMode.HTML,
    )


async def ajuda(update: Update, context) -> None:
    await start(update, context)


async def listar(update: Update, context) -> None:
    chunks = build_full_listing_chunks()
    for i, chunk in enumerate(chunks, start=1):
        footer = f"\n\n<i>({i}/{len(chunks)})</i>" if len(chunks) > 1 else ""
        await update.message.reply_text(chunk + footer, parse_mode=ParseMode.HTML)


async def text_handler(update: Update, context) -> None:
    term = update.message.text
    results = search(term)
    await update.message.reply_text(
        format_results(term, results), parse_mode=ParseMode.HTML
    )


async def inline_query(update: Update, context) -> None:
    query = update.inline_query.query
    if not query:
        return
    results = search(query, limit=15)
    articles = [
        InlineQueryResultArticle(
            id=str(uuid4()),
            title=f"{name} — R. {ramal}",
            description=f"Prédio/Unidade: {building}",
            input_message_content=InputTextMessageContent(
                f"📞 <b>{name}</b>\nRamal: <b>{ramal}</b> ({building})",
                parse_mode=ParseMode.HTML,
            ),
        )
        for name, ramal, building in results
    ]
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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    app.add_handler(InlineQueryHandler(inline_query))

    logger.info("Bot iniciado. Aguardando mensagens...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
