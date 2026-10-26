"""Gera um CSV de vendas fictício (com algumas sujeiras de propósito) para demonstração."""
import random
from datetime import date, timedelta
from pathlib import Path

PRODUTOS = {
    "Extintor ABC 4kg": 120.0,
    "Extintor CO2 6kg": 480.0,
    "Placa de Saída": 18.0,
    "Mangueira de Incêndio": 210.0,
    "Recarga Extintor": 65.0,
}
VENDEDORES = ["Ana", "Bruno", "Carla", "Diego"]


def gerar(caminho: str = "data/vendas_exemplo.csv", n: int = 300, seed: int = 42) -> None:
    random.seed(seed)
    inicio = date(2026, 1, 1)
    linhas = ["data,produto,vendedor,quantidade,preco_unitario"]
    for _ in range(n):
        dia = inicio + timedelta(days=random.randint(0, 270))
        produto = random.choice(list(PRODUTOS))
        vendedor = random.choice(VENDEDORES)
        qtd = random.randint(1, 8)
        preco = PRODUTOS[produto]
        linhas.append(f"{dia.isoformat()},{produto},{vendedor},{qtd},{preco}")

    # Sujeiras realistas: duplicata, linha sem quantidade e vendedor em minúsculo
    linhas.append(linhas[1])
    linhas.append("2026-03-10,Placa de Saída,Ana,,18.0")
    linhas.append("2026-03-11,Placa de Saída,  bruno ,2,18.0")

    Path(caminho).parent.mkdir(parents=True, exist_ok=True)
    Path(caminho).write_text("\n".join(linhas) + "\n", encoding="utf-8")


if __name__ == "__main__":
    gerar()
    print("Arquivo data/vendas_exemplo.csv gerado.")
