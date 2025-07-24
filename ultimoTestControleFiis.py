import flet as ft
import pandas as pd
from pathlib import Path
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from io import BytesIO
import base64

EXCEL_PATH = Path(__file__).parent / "controle_fiis_exportado.xlsx"

class ControleFIIsApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.aportes = []
        self.proventos = []
        self.vendas = []

        self.page.title = "Controle Avançado de FIIs"
        self.page.scroll = "auto"
        self.page.window_width = 1400
        self.page.window_height = 900
        self.page.theme_mode = ft.ThemeMode.LIGHT

        # Campos de aporte para adicionar (formulário) - organizados em grupos
        self.fundo_aporte = ft.TextField(label="FII", width=120, hint_text="Ex: HGLG11")
        self.tipo_aporte = ft.TextField(label="Tipo", width=100, hint_text="Ex: Compra")
        self.qtd_aporte = ft.TextField(label="Nº Cotas", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 100")
        self.preco_aporte = ft.TextField(label="Valor Cota (R$)", width=120, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 10.50")
        
        # Novos campos para análise
        self.setor_aporte = ft.Dropdown(
            label="Setor",
            width=150,
            options=[
                ft.dropdown.Option("Logística"),
                ft.dropdown.Option("Shoppings"),
                ft.dropdown.Option("Lajes Corporativas"),
                ft.dropdown.Option("Híbrido"),
                ft.dropdown.Option("Papel"),
                ft.dropdown.Option("Hospitalar"),
                ft.dropdown.Option("Educacional"),
                ft.dropdown.Option("Residencial"),
                ft.dropdown.Option("Outros"),
            ]
        )
        
        self.pvp_aporte = ft.TextField(label="P/VP", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 0.95")
        self.liquidez_aporte = ft.TextField(label="Liquidez Diária", width=120, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 1000000")
        self.vacancia_aporte = ft.TextField(label="Vacância (%)", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 5.2")
        
        self.dy_mes_aporte = ft.TextField(label="DY Mês", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 0.85")
        self.dy_ano_aporte = ft.TextField(label="DY Ano", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 10.20")
        self.dy_percentual_aporte = ft.TextField(label="DY %", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 8.5")
        
        self.dv_ano_aporte = ft.TextField(label="Valor DV Ano", width=120, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 1020.00")
        self.dv_mes_aporte = ft.TextField(label="Valor DV Mês", width=120, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 85.00")
        
        self.data_com_aporte = ft.TextField(label="Data COM", width=140, hint_text="dd/mm/aaaa")
        self.data_cadastrado_aporte = ft.TextField(label="Data Cadastrado", width=140, hint_text="dd/mm/aaaa", value=datetime.now().strftime("%d/%m/%Y"))

        # Campos de provento para adicionar
        self.fundo_provento = ft.TextField(label="Fundo", width=200, hint_text="Ex: HGLG11")
        self.valor_provento = ft.TextField(label="Rendimento por cota (R$)", width=180, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 0.85")

        # Campos de venda
        self.fundo_venda = ft.TextField(label="FII", width=120, hint_text="Ex: HGLG11")
        self.qtd_venda = ft.TextField(label="Nº Cotas", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 50")
        self.preco_venda = ft.TextField(label="Valor Cota (R$)", width=120, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 11.50")
        self.data_venda = ft.TextField(label="Data Venda", width=140, hint_text="dd/mm/aaaa", value=datetime.now().strftime("%d/%m/%Y"))

        # Tabelas
        self.tabela_aportes = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII", weight="bold")),
                ft.DataColumn(ft.Text("Setor", weight="bold")),
                ft.DataColumn(ft.Text("Tipo", weight="bold")),
                ft.DataColumn(ft.Text("Nº Cotas", weight="bold")),
                ft.DataColumn(ft.Text("Valor Cota", weight="bold")),
                ft.DataColumn(ft.Text("Valor Investido", weight="bold")),
                ft.DataColumn(ft.Text("P/VP", weight="bold")),
                ft.DataColumn(ft.Text("Vacância %", weight="bold")),
                ft.DataColumn(ft.Text("DY %", weight="bold")),
                ft.DataColumn(ft.Text("Data COM", weight="bold")),
                ft.DataColumn(ft.Text("Ações", weight="bold")),
            ], 
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
            horizontal_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
        )

        self.tabela_proventos = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Fundo", weight="bold")),
                ft.DataColumn(ft.Text("Valor R$", weight="bold")),
                ft.DataColumn(ft.Text("Data", weight="bold")),
                ft.DataColumn(ft.Text("Ações", weight="bold")),
            ], 
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
            horizontal_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
        )

        self.tabela_vendas = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII", weight="bold")),
                ft.DataColumn(ft.Text("Nº Cotas", weight="bold")),
                ft.DataColumn(ft.Text("Valor Cota", weight="bold")),
                ft.DataColumn(ft.Text("Valor Total", weight="bold")),
                ft.DataColumn(ft.Text("Data Venda", weight="bold")),
                ft.DataColumn(ft.Text("Lucro/Prejuízo", weight="bold")),
                ft.DataColumn(ft.Text("Ações", weight="bold")),
            ], 
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
            horizontal_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
        )

        self.tabela_rendimentos = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII", weight="bold")),
                ft.DataColumn(ft.Text("Setor", weight="bold")),
                ft.DataColumn(ft.Text("QTDE COTAS", weight="bold")),
                ft.DataColumn(ft.Text("PREÇO MÉDIO", weight="bold")),
                ft.DataColumn(ft.Text("VALOR ATUAL", weight="bold")),
                ft.DataColumn(ft.Text("P/VP", weight="bold")),
                ft.DataColumn(ft.Text("VACÂNCIA %", weight="bold")),
                ft.DataColumn(ft.Text("PROVENTOS", weight="bold")),
                ft.DataColumn(ft.Text("RENDIMENTO MÊS", weight="bold")),
                ft.DataColumn(ft.Text("RENDIMENTO ANO APROXIMADO", weight="bold")),
                ft.DataColumn(ft.Text("LUCRO/PREJUÍZO", weight="bold")),
                ft.DataColumn(ft.Text("DATA COM", weight="bold")),
            ], 
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
            horizontal_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
        )

        # Resumo detalhado em cards aprimorados
        self.card_total_cotas = ft.Container(
            content=ft.Column([
                ft.Text("Total de Cotas de FIIs", size=14, weight="bold", color=ft.Colors.WHITE),
                ft.Text("0", size=24, weight="bold", color=ft.Colors.WHITE),
                ft.Text("cotas", size=12, color=ft.Colors.WHITE70),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=20,
            bgcolor=ft.Colors.BLUE_600,
            border_radius=12,
            expand=True,
            alignment=ft.alignment.center
        )

        self.card_rendimento_mes = ft.Container(
            content=ft.Column([
                ft.Text("Rendimento Acumulado Mês", size=14, weight="bold", color=ft.Colors.WHITE),
                ft.Text("R$ 0,00", size=24, weight="bold", color=ft.Colors.WHITE),
                ft.Text("todos os FIIs", size=12, color=ft.Colors.WHITE70),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=20,
            bgcolor=ft.Colors.GREEN_600,
            border_radius=12,
            expand=True,
            alignment=ft.alignment.center
        )

        self.card_rendimento_ano = ft.Container(
            content=ft.Column([
                ft.Text("Rendimento Acumulado Ano", size=14, weight="bold", color=ft.Colors.WHITE),
                ft.Text("R$ 0,00", size=24, weight="bold", color=ft.Colors.WHITE),
                ft.Text("aproximado", size=12, color=ft.Colors.WHITE70),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=20,
            bgcolor=ft.Colors.PURPLE_600,
            border_radius=12,
            expand=True,
            alignment=ft.alignment.center
        )

        self.card_total_investido = ft.Container(
            content=ft.Column([
                ft.Text("Total Investido", size=14, weight="bold", color=ft.Colors.WHITE),
                ft.Text("R$ 0,00", size=24, weight="bold", color=ft.Colors.WHITE),
                ft.Text("patrimônio", size=12, color=ft.Colors.WHITE70),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=20,
            bgcolor=ft.Colors.ORANGE_600,
            border_radius=12,
            expand=True,
            alignment=ft.alignment.center
        )

        # Novos cards para análise
        self.card_dy_medio = ft.Container(
            content=ft.Column([
                ft.Text("DY Médio da Carteira", size=14, weight="bold", color=ft.Colors.WHITE),
                ft.Text("0,00%", size=24, weight="bold", color=ft.Colors.WHITE),
                ft.Text("dividend yield", size=12, color=ft.Colors.WHITE70),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=20,
            bgcolor=ft.Colors.TEAL_600,
            border_radius=12,
            expand=True,
            alignment=ft.alignment.center
        )

        self.card_pvp_medio = ft.Container(
            content=ft.Column([
                ft.Text("P/VP Médio da Carteira", size=14, weight="bold", color=ft.Colors.WHITE),
                ft.Text("0,00", size=24, weight="bold", color=ft.Colors.WHITE),
                ft.Text("preço/valor patrimonial", size=12, color=ft.Colors.WHITE70),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=20,
            bgcolor=ft.Colors.INDIGO_600,
            border_radius=12,
            expand=True,
            alignment=ft.alignment.center
        )

        # Diálogos
        self.dialog_edit_aporte = ft.AlertDialog(
            title=ft.Text("Editar Aporte"),
            modal=True,
            actions=[],
            content=ft.Column(width=700, height=500, scroll="auto")
        )
        
        self.dialog_edit_provento = ft.AlertDialog(
            title=ft.Text("Editar Provento"),
            modal=True,
            actions=[],
            content=ft.Column(width=400, height=200)
        )

        self.dialog_edit_venda = ft.AlertDialog(
            title=ft.Text("Editar Venda"),
            modal=True,
            actions=[],
            content=ft.Column(width=400, height=300)
        )

        self.dialog_confirm = ft.AlertDialog(
            title=ft.Text("Confirmar Exclusão"),
            modal=True,
            content=ft.Text(""),
            actions=[]
        )
        
        self.index_para_excluir = None
        self.excluir_tipo = None
        self.editando_aporte = None
        self.editando_provento = None
        self.editando_venda = None

        self.carregar_dados_excel()
        self.construir_interface()
        self.atualizar_tabelas()

    def construir_interface(self):
        # Formulário para adicionar aporte - reorganizado com novos campos
        grupo_basico = ft.Container(
            content=ft.Column([
                ft.Text("Informações Básicas", size=16, weight="bold", color=ft.Colors.BLUE_700),
                ft.Row([
                    self.fundo_aporte,
                    self.setor_aporte,
                    self.tipo_aporte,
                ], spacing=10, wrap=True),
                ft.Row([
                    self.qtd_aporte,
                    self.preco_aporte,
                ], spacing=10, wrap=True)
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50,
            expand=True
        )

        grupo_indicadores = ft.Container(
            content=ft.Column([
                ft.Text("Indicadores de Análise", size=16, weight="bold", color=ft.Colors.PURPLE_700),
                ft.Row([
                    self.pvp_aporte,
                    self.liquidez_aporte,
                    self.vacancia_aporte,
                ], spacing=10, wrap=True),
                ft.Row([
                    self.dy_percentual_aporte,
                    self.dy_mes_aporte,
                    self.dy_ano_aporte,
                ], spacing=10, wrap=True)
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50,
            expand=True
        )

        grupo_datas = ft.Container(
            content=ft.Column([
                ft.Text("Datas", size=16, weight="bold", color=ft.Colors.ORANGE_700),
                ft.Row([
                    self.data_com_aporte,
                    self.data_cadastrado_aporte,
                ], spacing=10, wrap=True)
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50,
            expand=True
        )

        # Organizando os grupos em linha horizontal
        grupos_horizontais = ft.Row([
            grupo_basico,
            grupo_indicadores,
            grupo_datas,
        ], spacing=15, expand=True)

        botao_adicionar = ft.Container(
            content=ft.ElevatedButton(
                "Adicionar Aporte", 
                on_click=self.adicionar_aporte,
                bgcolor=ft.Colors.BLUE_600,
                color=ft.Colors.WHITE,
                width=200,
                height=40
            ),
            alignment=ft.alignment.center,
            padding=10
        )

        formulario_aporte = ft.Column([
            ft.Text("Novo Aporte", size=20, weight="bold", color=ft.Colors.BLUE_800),
            ft.Divider(),
            grupos_horizontais,
            botao_adicionar,
        ], spacing=15)

        tab_aportes = ft.Column([
            formulario_aporte,
            ft.Divider(height=20),
            ft.Text("Lista de Aportes", size=18, weight="bold"),
            ft.Container(
                content=self.tabela_aportes,
                border=ft.border.all(1, ft.Colors.GREY_300),
                border_radius=10,
                padding=10
            ),
        ], scroll="auto", expand=True, spacing=10)

        # Tab de proventos
        formulario_provento = ft.Container(
            content=ft.Column([
                ft.Text("Novo Provento", size=20, weight="bold", color=ft.Colors.GREEN_800),
                ft.Divider(),
                ft.Row([
                    self.fundo_provento, 
                    self.valor_provento, 
                    ft.ElevatedButton(
                        "Adicionar Provento", 
                        on_click=self.adicionar_provento,
                        bgcolor=ft.Colors.GREEN_600,
                        color=ft.Colors.WHITE,
                        height=40
                    )
                ], spacing=10, wrap=True)
            ]),
            padding=15,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            bgcolor=ft.Colors.GREY_50
        )

        tab_proventos = ft.Column([
            formulario_provento,
            ft.Divider(height=20),
            ft.Text("Lista de Proventos", size=18, weight="bold"),
            ft.Container(
                content=self.tabela_proventos,
                border=ft.border.all(1, ft.Colors.GREY_300),
                border_radius=10,
                padding=10
            ),
        ], scroll="auto", expand=True, spacing=10)

        # Nova tab de vendas
        formulario_venda = ft.Container(
            content=ft.Column([
                ft.Text("Nova Venda", size=20, weight="bold", color=ft.Colors.RED_800),
                ft.Divider(),
                ft.Row([
                    self.fundo_venda,
                    self.qtd_venda,
                    self.preco_venda,
                    self.data_venda,
                    ft.ElevatedButton(
                        "Registrar Venda", 
                        on_click=self.adicionar_venda,
                        bgcolor=ft.Colors.RED_600,
                        color=ft.Colors.WHITE,
                        height=40
                    )
                ], spacing=10, wrap=True)
            ]),
            padding=15,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            bgcolor=ft.Colors.GREY_50
        )

        tab_vendas = ft.Column([
            formulario_venda,
            ft.Divider(height=20),
            ft.Text("Lista de Vendas", size=18, weight="bold"),
            ft.Container(
                content=self.tabela_vendas,
                border=ft.border.all(1, ft.Colors.GREY_300),
                border_radius=10,
                padding=10
            ),
        ], scroll="auto", expand=True, spacing=10)

        # Tab de rendimentos aprimorada
        tab_rendimentos = ft.Column([
            ft.Text("Resumo de Rendimentos", size=20, weight="bold", color=ft.Colors.PURPLE_800),
            ft.Divider(),
            # Primeira linha de cards
            ft.Row([
                self.card_total_cotas,
                self.card_rendimento_mes,
                self.card_rendimento_ano,
                self.card_total_investido,
            ], spacing=15, expand=True),
            # Segunda linha de cards
            ft.Row([
                self.card_dy_medio,
                self.card_pvp_medio,
            ], spacing=15, expand=True),
            ft.Divider(height=20),
            ft.Container(
                content=self.tabela_rendimentos,
                border=ft.border.all(1, ft.Colors.GREY_300),
                border_radius=10,
                padding=10
            ),
        ], scroll="auto", expand=True, spacing=15)

        self.tabs = ft.Tabs(
            selected_index=0, 
            tabs=[
                ft.Tab(text="📊 Aportes", content=tab_aportes),
                ft.Tab(text="💰 Proventos", content=tab_proventos),
                ft.Tab(text="💸 Vendas", content=tab_vendas),
                ft.Tab(text="📈 Rendimentos", content=tab_rendimentos),
            ], 
            expand=True,
            tab_alignment=ft.TabAlignment.CENTER
        )

        self.page.add(self.tabs, self.dialog_edit_aporte, self.dialog_edit_provento, self.dialog_edit_venda, self.dialog_confirm)

    def carregar_dados_excel(self):
        self.aportes.clear()
        self.proventos.clear()
        self.vendas.clear()
        if EXCEL_PATH.exists():
            try:
                with pd.ExcelFile(EXCEL_PATH) as reader:
                    try:
                        df_ap = pd.read_excel(reader, sheet_name="Aportes")
                        for _, row in df_ap.iterrows():
                            self.aportes.append({
                                "fundo": str(row["fundo"]),
                                "setor": str(row.get("setor", "")),
                                "tipo": str(row.get("tipo", "")),
                                "quantidade": int(row["quantidade"]),
                                "preco": float(row["preco"]),
                                "pvp": float(row.get("pvp", 0)),
                                "liquidez": float(row.get("liquidez", 0)),
                                "vacancia": float(row.get("vacancia", 0)),
                                "dy_mes": float(row.get("dy_mes", 0)),
                                "dy_ano": float(row.get("dy_ano", 0)),
                                "dy_percentual": float(row.get("dy_percentual", 0)),
                                "dv_ano": float(row.get("dv_ano", 0)),
                                "dv_mes": float(row.get("dv_mes", 0)),
                                "data_com": str(row.get("data_com", "")),
                                "data_cadastrado": str(row.get("data_cadastrado", "")),
                            })
                    except Exception:
                        pass
                    try:
                        df_pr = pd.read_excel(reader, sheet_name="Proventos")
                        for _, row in df_pr.iterrows():
                            self.proventos.append({
                                "fundo": str(row["fundo"]),
                                "valor": float(row["valor"]),
                                "data": str(row.get("data", "Desconhecida"))
                            })
                    except Exception:
                        pass
                    try:
                        df_vd = pd.read_excel(reader, sheet_name="Vendas")
                        for _, row in df_vd.iterrows():
                            self.vendas.append({
                                "fundo": str(row["fundo"]),
                                "quantidade": int(row["quantidade"]),
                                "preco": float(row["preco"]),
                                "data": str(row.get("data", "Desconhecida"))
                            })
                    except Exception:
                        pass
            except Exception as e:
                print("Erro ao carregar Excel:", e)

    def calcular_preco_medio(self, fundo):
        """Calcula o preço médio de aquisição de um FII específico"""
        aportes_fundo = [a for a in self.aportes if a["fundo"] == fundo]
        vendas_fundo = [v for v in self.vendas if v["fundo"] == fundo]
        
        if not aportes_fundo:
            return 0
        
        total_cotas_compradas = sum(a["quantidade"] for a in aportes_fundo)
        total_valor_investido = sum(a["quantidade"] * a["preco"] for a in aportes_fundo)
        
        if total_cotas_compradas == 0:
            return 0
        
        return total_valor_investido / total_cotas_compradas

    def calcular_cotas_atuais(self, fundo):
        """Calcula a quantidade atual de cotas de um FII (aportes - vendas)"""
        aportes_fundo = [a for a in self.aportes if a["fundo"] == fundo]
        vendas_fundo = [v for v in self.vendas if v["fundo"] == fundo]
        
        total_compradas = sum(a["quantidade"] for a in aportes_fundo)
        total_vendidas = sum(v["quantidade"] for v in vendas_fundo)
        
        return total_compradas - total_vendidas

    def salvar_excel(self):
        try:
            df_ap = pd.DataFrame(self.aportes)
            df_pr = pd.DataFrame(self.proventos)
            df_vd = pd.DataFrame(self.vendas)

            # Preparar dados para a aba de Rendimentos com novos cálculos
            df_rendimentos = pd.DataFrame(columns=["FII", "SETOR", "QTDE COTAS", "PREÇO MÉDIO", "VALOR ATUAL", "P/VP", "VACÂNCIA %", "PROVENTOS", "RENDIMENTO MÊS", "RENDIMENTO ANO APROXIMADO", "LUCRO/PREJUÍZO", "DATA COM"])
            
            if self.aportes:
                fundos_unicos = list(set(a["fundo"] for a in self.aportes))
                dados_rendimentos = []
                
                for fundo in fundos_unicos:
                    cotas_atuais = self.calcular_cotas_atuais(fundo)
                    if cotas_atuais <= 0:
                        continue
                        
                    preco_medio = self.calcular_preco_medio(fundo)
                    
                    # Pegar dados mais recentes do fundo
                    aportes_fundo = [a for a in self.aportes if a["fundo"] == fundo]
                    ultimo_aporte = max(aportes_fundo, key=lambda x: x.get("data_cadastrado", ""))
                    
                    # Calcular proventos
                    proventos_fundo = [p for p in self.proventos if p["fundo"] == fundo]
                    total_proventos = sum(p["valor"] for p in proventos_fundo)
                    
                    # Cálculos
                    rendimento_mes = cotas_atuais * total_proventos
                    rendimento_ano = rendimento_mes * 12
                    valor_atual = cotas_atuais * preco_medio  # Aqui seria o preço atual se tivéssemos API
                    lucro_prejuizo = valor_atual - (cotas_atuais * preco_medio)  # Seria diferente com preço atual
                    
                    dados_rendimentos.append({
                        "FII": fundo,
                        "SETOR": ultimo_aporte.get("setor", ""),
                        "QTDE COTAS": cotas_atuais,
                        "PREÇO MÉDIO": preco_medio,
                        "VALOR ATUAL": valor_atual,
                        "P/VP": ultimo_aporte.get("pvp", 0),
                        "VACÂNCIA %": ultimo_aporte.get("vacancia", 0),
                        "PROVENTOS": total_proventos,
                        "RENDIMENTO MÊS": rendimento_mes,
                        "RENDIMENTO ANO APROXIMADO": rendimento_ano,
                        "LUCRO/PREJUÍZO": lucro_prejuizo,
                        "DATA COM": ultimo_aporte.get("data_com", "")
                    })
                
                df_rendimentos = pd.DataFrame(dados_rendimentos)

            with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
                df_ap.to_excel(writer, sheet_name="Aportes", index=False)
                df_pr.to_excel(writer, sheet_name="Proventos", index=False)
                df_vd.to_excel(writer, sheet_name="Vendas", index=False)
                df_rendimentos.to_excel(writer, sheet_name="Rendimentos", index=False)
        except Exception as e:
            self.show_snack(f"Erro ao salvar Excel: {e}", ft.Colors.RED)

    def show_snack(self, msg, color):
        self.page.snack_bar = ft.SnackBar(
            ft.Text(msg, color=ft.Colors.WHITE), 
            bgcolor=color, 
            open=True,
            duration=3000
        )
        self.page.update()

    def limpar_campos_aporte(self):
        """Função para limpar todos os campos do formulário de aporte"""
        self.fundo_aporte.value = ""
        self.setor_aporte.value = None
        self.tipo_aporte.value = ""
        self.qtd_aporte.value = ""
        self.preco_aporte.value = ""
        self.pvp_aporte.value = ""
        self.liquidez_aporte.value = ""
        self.vacancia_aporte.value = ""
        self.dy_mes_aporte.value = ""
        self.dy_ano_aporte.value = ""
        self.dy_percentual_aporte.value = ""
        self.dv_ano_aporte.value = ""
        self.dv_mes_aporte.value = ""
        self.data_com_aporte.value = ""
        self.data_cadastrado_aporte.value = datetime.now().strftime("%d/%m/%Y")
        self.page.update()

    def validar_campos_aporte(self):
        """Função para validar os campos obrigatórios do aporte"""
        erros = []
        
        if not self.fundo_aporte.value or not self.fundo_aporte.value.strip():
            erros.append("FII é obrigatório")
        
        try:
            qtd = int(self.qtd_aporte.value) if self.qtd_aporte.value else 0
            if qtd <= 0:
                erros.append("Número de cotas deve ser maior que zero")
        except ValueError:
            erros.append("Número de cotas deve ser um número válido")
        
        try:
            preco = float(self.preco_aporte.value) if self.preco_aporte.value else 0
            if preco <= 0:
                erros.append("Valor da cota deve ser maior que zero")
        except ValueError:
            erros.append("Valor da cota deve ser um número válido")
        
        return erros

    def atualizar_tabelas(self):
        # Atualiza tabela de aportes
        self.tabela_aportes.rows.clear()
        for i, a in enumerate(self.aportes):
            valor_investido = a["quantidade"] * a["preco"]
            
            acoes = ft.Row([
                ft.IconButton(
                    ft.Icons.EDIT, 
                    tooltip="Editar", 
                    data=i, 
                    on_click=self.abrir_edicao_aporte,
                    icon_color=ft.Colors.BLUE_600
                ),
                ft.IconButton(
                    ft.Icons.DELETE, 
                    tooltip="Excluir", 
                    data=i, 
                    on_click=self.abrir_confirmacao_exclusao_aporte,
                    icon_color=ft.Colors.RED_600
                ),
            ], spacing=5)
            
            self.tabela_aportes.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(a["fundo"], weight="bold")),
                ft.DataCell(ft.Text(a.get("setor", ""))),
                ft.DataCell(ft.Text(a["tipo"])),
                ft.DataCell(ft.Text(str(a["quantidade"]))),
                ft.DataCell(ft.Text(f'R$ {a["preco"]:.2f}')),
                ft.DataCell(ft.Text(f'R$ {valor_investido:.2f}', weight="bold")),
                ft.DataCell(ft.Text(f'{a.get("pvp", 0):.2f}')),
                ft.DataCell(ft.Text(f'{a.get("vacancia", 0):.1f}%')),
                ft.DataCell(ft.Text(f'{a.get("dy_percentual", 0):.2f}%')),
                ft.DataCell(ft.Text(a.get("data_com", ""))),
                ft.DataCell(acoes),
            ]))

        # Atualiza tabela de proventos
        self.tabela_proventos.rows.clear()
        for i, p in enumerate(self.proventos):
            acoes = ft.Row([
                ft.IconButton(
                    ft.Icons.EDIT, 
                    tooltip="Editar", 
                    data=i, 
                    on_click=self.abrir_edicao_provento,
                    icon_color=ft.Colors.BLUE_600
                ),
                ft.IconButton(
                    ft.Icons.DELETE, 
                    tooltip="Excluir", 
                    data=i, 
                    on_click=self.abrir_confirmacao_exclusao_provento,
                    icon_color=ft.Colors.RED_600
                ),
            ], spacing=5)
            
            self.tabela_proventos.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(p["fundo"], weight="bold")),
                ft.DataCell(ft.Text(f'R$ {p["valor"]:.2f}')),
                ft.DataCell(ft.Text(p["data"])),
                ft.DataCell(acoes),
            ]))

        # Atualiza tabela de vendas
        self.tabela_vendas.rows.clear()
        for i, v in enumerate(self.vendas):
            valor_total = v["quantidade"] * v["preco"]
            preco_medio = self.calcular_preco_medio(v["fundo"])
            lucro_prejuizo = (v["preco"] - preco_medio) * v["quantidade"]
            
            acoes = ft.Row([
                ft.IconButton(
                    ft.Icons.EDIT, 
                    tooltip="Editar", 
                    data=i, 
                    on_click=self.abrir_edicao_venda,
                    icon_color=ft.Colors.BLUE_600
                ),
                ft.IconButton(
                    ft.Icons.DELETE, 
                    tooltip="Excluir", 
                    data=i, 
                    on_click=self.abrir_confirmacao_exclusao_venda,
                    icon_color=ft.Colors.RED_600
                ),
            ], spacing=5)
            
            cor_lucro = ft.Colors.GREEN if lucro_prejuizo >= 0 else ft.Colors.RED
            
            self.tabela_vendas.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(v["fundo"], weight="bold")),
                ft.DataCell(ft.Text(str(v["quantidade"]))),
                ft.DataCell(ft.Text(f'R$ {v["preco"]:.2f}')),
                ft.DataCell(ft.Text(f'R$ {valor_total:.2f}', weight="bold")),
                ft.DataCell(ft.Text(v["data"])),
                ft.DataCell(ft.Text(f'R$ {lucro_prejuizo:.2f}', color=cor_lucro, weight="bold")),
                ft.DataCell(acoes),
            ]))

        # Atualiza tabela rendimentos com novos cálculos
        self.tabela_rendimentos.rows.clear()

        if self.aportes:
            fundos_unicos = list(set(a["fundo"] for a in self.aportes))
            dados_consolidados = []
            
            for fundo in fundos_unicos:
                cotas_atuais = self.calcular_cotas_atuais(fundo)
                if cotas_atuais <= 0:
                    continue
                    
                preco_medio = self.calcular_preco_medio(fundo)
                
                # Pegar dados mais recentes do fundo
                aportes_fundo = [a for a in self.aportes if a["fundo"] == fundo]
                ultimo_aporte = max(aportes_fundo, key=lambda x: x.get("data_cadastrado", ""))
                
                # Calcular proventos
                proventos_fundo = [p for p in self.proventos if p["fundo"] == fundo]
                total_proventos = sum(p["valor"] for p in proventos_fundo)
                
                # Cálculos
                rendimento_mes = cotas_atuais * total_proventos
                rendimento_ano = rendimento_mes * 12
                valor_atual = cotas_atuais * preco_medio
                lucro_prejuizo = 0  # Seria calculado com preço atual vs preço médio
                
                dados_consolidados.append({
                    "fundo": fundo,
                    "setor": ultimo_aporte.get("setor", ""),
                    "cotas_atuais": cotas_atuais,
                    "preco_medio": preco_medio,
                    "valor_atual": valor_atual,
                    "pvp": ultimo_aporte.get("pvp", 0),
                    "vacancia": ultimo_aporte.get("vacancia", 0),
                    "total_proventos": total_proventos,
                    "rendimento_mes": rendimento_mes,
                    "rendimento_ano": rendimento_ano,
                    "lucro_prejuizo": lucro_prejuizo,
                    "data_com": ultimo_aporte.get("data_com", ""),
                    "dy_percentual": ultimo_aporte.get("dy_percentual", 0)
                })

            # Calcular totais para os cards
            total_cotas = sum(d["cotas_atuais"] for d in dados_consolidados)
            total_rendimento_mes = sum(d["rendimento_mes"] for d in dados_consolidados)
            total_rendimento_ano = sum(d["rendimento_ano"] for d in dados_consolidados)
            valor_total_investido = sum(d["valor_atual"] for d in dados_consolidados)
            
            # Calcular médias ponderadas
            if valor_total_investido > 0:
                dy_medio = sum(d["dy_percentual"] * d["valor_atual"] for d in dados_consolidados) / valor_total_investido
                pvp_medio = sum(d["pvp"] * d["valor_atual"] for d in dados_consolidados if d["pvp"] > 0) / valor_total_investido
            else:
                dy_medio = 0
                pvp_medio = 0

            # Atualizar os cards com os valores calculados
            self.card_total_cotas.content.controls[1].value = f"{int(total_cotas):,}".replace(",", ".")
            self.card_rendimento_mes.content.controls[1].value = f"R$ {total_rendimento_mes:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.card_rendimento_ano.content.controls[1].value = f"R$ {total_rendimento_ano:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.card_total_investido.content.controls[1].value = f"R$ {valor_total_investido:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.card_dy_medio.content.controls[1].value = f"{dy_medio:.2f}%"
            self.card_pvp_medio.content.controls[1].value = f"{pvp_medio:.2f}"

            # Preencher tabela
            for dados in dados_consolidados:
                cor_lucro = ft.Colors.GREEN if dados["lucro_prejuizo"] >= 0 else ft.Colors.RED
                
                self.tabela_rendimentos.rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(dados["fundo"], weight="bold")),
                    ft.DataCell(ft.Text(dados["setor"])),
                    ft.DataCell(ft.Text(str(int(dados["cotas_atuais"])))),
                    ft.DataCell(ft.Text(f'R$ {dados["preco_medio"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {dados["valor_atual"]:.2f}')),
                    ft.DataCell(ft.Text(f'{dados["pvp"]:.2f}')),
                    ft.DataCell(ft.Text(f'{dados["vacancia"]:.1f}%')),
                    ft.DataCell(ft.Text(f'R$ {dados["total_proventos"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {dados["rendimento_mes"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {dados["rendimento_ano"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {dados["lucro_prejuizo"]:.2f}', color=cor_lucro)),
                    ft.DataCell(ft.Text(dados["data_com"])),
                ]))

        else:
            # Resetar os cards quando não há dados
            self.card_total_cotas.content.controls[1].value = "0"
            self.card_rendimento_mes.content.controls[1].value = "R$ 0,00"
            self.card_rendimento_ano.content.controls[1].value = "R$ 0,00"
            self.card_total_investido.content.controls[1].value = "R$ 0,00"
            self.card_dy_medio.content.controls[1].value = "0,00%"
            self.card_pvp_medio.content.controls[1].value = "0,00"

        self.page.update()

    def adicionar_aporte(self, e):
        # Validar campos primeiro
        erros = self.validar_campos_aporte()
        if erros:
            self.show_snack(f"Erro: {'; '.join(erros)}", ft.Colors.RED)
            return

        try:
            fundo = self.fundo_aporte.value.strip().upper()
            setor = self.setor_aporte.value or ""
            tipo = self.tipo_aporte.value.strip()
            quantidade = int(self.qtd_aporte.value)
            preco = float(self.preco_aporte.value)
            pvp = float(self.pvp_aporte.value or 0)
            liquidez = float(self.liquidez_aporte.value or 0)
            vacancia = float(self.vacancia_aporte.value or 0)
            dy_mes = float(self.dy_mes_aporte.value or 0)
            dy_ano = float(self.dy_ano_aporte.value or 0)
            dy_percentual = float(self.dy_percentual_aporte.value or 0)
            dv_ano = float(self.dv_ano_aporte.value or 0)
            dv_mes = float(self.dv_mes_aporte.value or 0)
            data_com = self.data_com_aporte.value.strip()
            data_cadastrado = self.data_cadastrado_aporte.value.strip()

            self.aportes.append({
                "fundo": fundo,
                "setor": setor,
                "tipo": tipo,
                "quantidade": quantidade,
                "preco": preco,
                "pvp": pvp,
                "liquidez": liquidez,
                "vacancia": vacancia,
                "dy_mes": dy_mes,
                "dy_ano": dy_ano,
                "dy_percentual": dy_percentual,
                "dv_ano": dv_ano,
                "dv_mes": dv_mes,
                "data_com": data_com,
                "data_cadastrado": data_cadastrado,
            })

            # Limpar campos
            self.limpar_campos_aporte()
            
            self.salvar_excel()
            self.atualizar_tabelas()
            self.show_snack(f"Aporte adicionado com sucesso: {fundo}", ft.Colors.GREEN)
            
        except Exception as ex:
            self.show_snack(f"Erro ao adicionar aporte: {ex}", ft.Colors.RED)

    def adicionar_provento(self, e):
        try:
            fundo = self.fundo_provento.value.strip().upper()
            valor = float(self.valor_provento.value)
            if not fundo or valor <= 0:
                raise ValueError("Preencha todos os campos corretamente")
            data_hoje = datetime.now().strftime("%d/%m/%Y")
            self.proventos.append({"fundo": fundo, "valor": valor, "data": data_hoje})
            self.fundo_provento.value = ""
            self.valor_provento.value = ""
            self.salvar_excel()
            self.atualizar_tabelas()
            self.show_snack(f"Provento adicionado: {fundo}", ft.Colors.GREEN)
        except Exception as ex:
            self.show_snack(f"Erro ao adicionar provento: {ex}", ft.Colors.RED)

    def adicionar_venda(self, e):
        try:
            fundo = self.fundo_venda.value.strip().upper()
            quantidade = int(self.qtd_venda.value)
            preco = float(self.preco_venda.value)
            data = self.data_venda.value.strip()
            
            if not fundo or quantidade <= 0 or preco <= 0:
                raise ValueError("Preencha todos os campos corretamente")
            
            # Verificar se há cotas suficientes
            cotas_atuais = self.calcular_cotas_atuais(fundo)
            if quantidade > cotas_atuais:
                raise ValueError(f"Quantidade insuficiente. Você possui apenas {cotas_atuais} cotas de {fundo}")
            
            self.vendas.append({
                "fundo": fundo, 
                "quantidade": quantidade, 
                "preco": preco, 
                "data": data
            })
            
            self.fundo_venda.value = ""
            self.qtd_venda.value = ""
            self.preco_venda.value = ""
            self.data_venda.value = datetime.now().strftime("%d/%m/%Y")
            
            self.salvar_excel()
            self.atualizar_tabelas()
            self.show_snack(f"Venda registrada: {fundo}", ft.Colors.GREEN)
        except Exception as ex:
            self.show_snack(f"Erro ao registrar venda: {ex}", ft.Colors.RED)

    # Métodos de edição e exclusão (simplificados para o exemplo)
    def abrir_edicao_aporte(self, e):
        # Implementação similar ao código original, mas com novos campos
        pass

    def abrir_edicao_provento(self, e):
        # Implementação similar ao código original
        pass

    def abrir_edicao_venda(self, e):
        # Nova implementação para edição de vendas
        pass

    def abrir_confirmacao_exclusao_aporte(self, e):
        # Implementação similar ao código original
        pass

    def abrir_confirmacao_exclusao_provento(self, e):
        # Implementação similar ao código original
        pass

    def abrir_confirmacao_exclusao_venda(self, e):
        # Nova implementação para exclusão de vendas
        pass

def main(page: ft.Page):
    ControleFIIsApp(page)

if __name__ == "__main__":
    ft.app(target=main, view=ft.FLET_APP)