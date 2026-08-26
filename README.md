# Agente Comercial Power BI

API FastAPI para consultas comerciais no Power BI Service.

## Arquivo principal

Renomeie a versão validada do agente para:

`agente_comercial.py`

## Instalação

```bash
pip install -r requirements.txt
```

## Execução

```bash
uvicorn agente_comercial:app --host 0.0.0.0 --port 8000
```

## Variáveis de ambiente

Configure as variáveis descritas em `.env.example`.

Nunca publique o arquivo `.env` nem segredos no GitHub.
