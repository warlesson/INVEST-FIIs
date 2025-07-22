import flet as ft
import requests
from flet import Colors

# Funções para buscar dados reais
def buscar_dados_bcb(codigo_serie):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo_serie}/dados/ultimos/1?formato=json"
    try:
        resposta = requests.get(url, timeout=10)
        if resposta.status_code == 200:
            return resposta.json()[0]["valor"]
    except:
        pass
    return "Erro"

def buscar_cripto(nome):
    url = f"https://api.coingecko.com/api/v3/simple/price?ids={nome}&vs_currencies=brl"
    try:
        resposta = requests.get(url, timeout=10)
        if resposta.status_code == 200:
            return resposta.json()[nome]["brl"]
    except:
        pass
    return "Erro"

# App Flet
def main(page: ft.Page):
    page.title = "Painel Econômico - Investimentos"
    page.scroll = ft.ScrollMode.ALWAYS
    page.theme_mode = ft.ThemeMode.LIGHT

    titulo = ft.Text("📊 Painel Econômico Atualizado", size=30, weight="bold")

    nomes = [
        ("CDI Anual", "12"),      # CDI
        ("SELIC", "11"),          # SELIC
        ("IPCA", "433"),          # IPCA
        ("IBOVESPA", None),       # Valor fixo
        ("Dólar", "1"),           # Dólar
        ("Euro", "21619"),        # Euro
    ]

    criptoativos = [
        ("Bitcoin", "bitcoin"),
        ("Ethereum", "ethereum"),
    ]

    def criar_cartao(nome, valor):
        return ft.Container(
            content=ft.Column([
                ft.Text(nome, size=16, weight="bold"),
                ft.Text(valor, size=20, color="green")
            ]),
            padding=15,
            width=200,
            bgcolor=Colors.BLUE_100,
            border_radius=10
        )

    def atualizar_dados(e):
        cards.controls.clear()

        for nome, codigo in nomes:
            if codigo:
                valor = buscar_dados_bcb(codigo)
                try:
                    texto_valor = f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " %"
                except:
                    texto_valor = "Indisponível"
            else:
                texto_valor = "126.785 pts"
            cards.controls.append(criar_cartao(nome, texto_valor))

        for nome, cripto in criptoativos:
            preco = buscar_cripto(cripto)
            try:
                valor = f"R$ {float(preco):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except:
                valor = "Indisponível"
            cards.controls.append(criar_cartao(nome, valor))

        page.update()

    cards = ft.Row(wrap=True, spacing=20)
    botao_atualizar = ft.ElevatedButton("🔄 Atualizar Agora", on_click=atualizar_dados)
    page.add(titulo, botao_atualizar, cards)

    atualizar_dados(None)

ft.app(target=main)
