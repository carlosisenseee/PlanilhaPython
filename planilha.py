import requests
import json

url = "https://apiintegracao.milvus.com.br/api/chamado/listagem"

payload = {
    "filtro_body": {
        # "tipo_ticket": "Novo Colaborador"
        "tipo_ticket": "Novo Colaborador"
    }
}

headers = {
    "Content-Type": "application/json",
    "Authorization": "7X0CuWWQydCeiBWRpp1fTipDQjC3mWqw7pmUjEdMySCY0iB1KyRm1uWfp4l8BBoWrULb2JvlDskE9YLTd1D8v8x4QhSFlYS5QCknx"
}

response = requests.post(url, headers=headers, json=payload, timeout=30)
response.raise_for_status()

dados = response.json()

for chamado in dados.get("lista", []):
    # Seperar os dados em colunas para enviar para o excel
    print(f'{chamado["codigo"]}')
    print(f'{chamado.get("assunto")}')
    print(f'{chamado.get("descricao")}')