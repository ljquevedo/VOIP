# Bot de Ramais VOIP — HUSFP 2026

Bot de Telegram para consultar o guia de ramais do Hospital Universitário
São Francisco de Paula, tanto em mensagem direta quanto em **modo inline**.

## 1. Ativar o modo inline (você já tem o token)

No chat com o **@BotFather** no Telegram:

1. Envie `/setinline`
2. Escolha o seu bot
3. Informe um texto de placeholder, ex: `Digite o nome do setor...`

Sem esse passo, a busca por mensagem direta funciona normalmente, mas
`@seu_bot nome_do_setor` em outros chats não vai aparecer.

## 2. Instalar dependências

```bash
pip install -r requirements.txt --break-system-packages
```

(ou, melhor ainda, em um ambiente virtual: `python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt`)

## 3. Configurar o token

```bash
export TELEGRAM_BOT_TOKEN="seu_token_aqui"
```

No Windows (PowerShell): `$env:TELEGRAM_BOT_TOKEN="seu_token_aqui"`

Nunca coloque o token direto no código-fonte, especialmente se for subir
para um repositório público.

## 4. Rodar

```bash
python bot.py
```

O bot roda em modo *polling* (não precisa de servidor com IP público nem
HTTPS). Ele fica ativo enquanto o processo estiver rodando.

## 5. Usar

- **Mensagem direta:** envie `/start` e depois qualquer termo, ex: `farmácia`
  ou `8560`.
- **Modo inline:** em qualquer chat, digite `@seu_bot farmácia` — o Telegram
  mostra os resultados para você escolher e enviar.

## Sobre os dados

Os 144 ramais estão embutidos diretamente em `bot.py` (lista `DATA`),
extraídos do guia HTML/PDF de 2026. A busca ignora acentos e
maiúsculas/minúsculas, e também casa números de ramal parciais.

### Atualizar os ramais no futuro

Edite a lista `DATA` no início de `bot.py`, seguindo o formato:

```python
("Nome do Setor", "ramal", "SIGLA_DO_PREDIO"),
```

Não é necessário reiniciar nada além do processo do bot (`Ctrl+C` e rodar
`python bot.py` de novo) para os novos dados entrarem em vigor.

## 6. Manter rodando 24/7 (opcional)

Para produção, rode com um gerenciador de processos, por exemplo:

```bash
# com screen
screen -S ramais-bot
python bot.py
# Ctrl+A depois D para sair sem encerrar

# ou com systemd (Linux), criando um serviço que executa:
# ExecStart=/caminho/venv/bin/python /caminho/bot.py
# Environment=TELEGRAM_BOT_TOKEN=seu_token_aqui
```
