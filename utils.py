import os
import json
import re
import datetime
import zipfile
import io
from pathlib import Path

# Arquivo onde salvaremos o "estado" atual (qual pasta está sendo processada)
ARQUIVO_ESTADO = "sessao_atual.json"

def salvar_estado(chave, valor):
    """Salva uma informação para ser usada pelo próximo script."""
    dados = {}
    if os.path.exists(ARQUIVO_ESTADO):
        try:
            with open(ARQUIVO_ESTADO, 'r') as f:
                dados = json.load(f)
        except: pass
    
    dados[chave] = str(valor) # Converte Path para string
    
    with open(ARQUIVO_ESTADO, 'w') as f:
        json.dump(dados, f, indent=4)
    print(f"💾 Estado salvo: {chave} -> {valor}")

def carregar_estado(chave):
    """Recupera uma informação salva anteriormente."""
    if not os.path.exists(ARQUIVO_ESTADO):
        return None
    try:
        with open(ARQUIVO_ESTADO, 'r') as f:
            dados = json.load(f)
            return dados.get(chave)
    except:
        return None