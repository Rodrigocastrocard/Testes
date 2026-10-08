import os


def getEnv(nome: str) -> str:
    valor = os.getenv(nome)
    if valor is None or str(valor).strip() == "":
        raise ValueError(f"Variável de ambiente '{nome}' não definida.")
    return valor