import flet as ft
import pandas as pd
from pathlib import Path
from datetime import datetime

EXCEL_PATH = Path(__file__).parent / "controle_fiis_exportado.xlsx"

class ControleFIIsApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.aportes = []
        self.proventos = []

        self.page.title = "Controle de FIIs"
        self.page.scroll = "auto"
        self.page.window_width = 1200
        self.page.window_height = 800
        self.page.theme_mode = ft.ThemeMode.LIGHT

        # Campos de aporte para adicionar (formulário) - organizados em grupos
        self.fundo_aporte = ft.TextField(label="FII", width=120, hint_text="Ex: HGLG11")
        self.tipo_aporte = ft.TextField(label="Tipo", width=100, hint_text="Ex: Compra")
        self.qtd_aporte = ft.TextField(label="Nº Cotas", width=100, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 100")
        self.preco_aporte = ft.TextField(label="Valor Cota (R$)", width=120, keyboard_type=ft.KeyboardType.NUMBER, hint_text="Ex: 10.50")
        
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

        # Tabelas
        self.tabela_aportes = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII", weight="bold")),
                ft.DataColumn(ft.Text("Tipo", weight="bold")),
                ft.DataColumn(ft.Text("Nº Cotas", weight="bold")),
                ft.DataColumn(ft.Text("Valor Cota", weight="bold")),
                ft.DataColumn(ft.Text("Valor Investido", weight="bold")),
                ft.DataColumn(ft.Text("DY Mês", weight="bold")),
                ft.DataColumn(ft.Text("DY Ano", weight="bold")),
                ft.DataColumn(ft.Text("DY %", weight="bold")),
                ft.DataColumn(ft.Text("Valor DV Ano", weight="bold")),
                ft.DataColumn(ft.Text("Valor DV Mês", weight="bold")),
                ft.DataColumn(ft.Text("Data COM", weight="bold")),
                ft.DataColumn(ft.Text("Data Cadastrado", weight="bold")),
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

        self.tabela_rendimentos = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("FII", weight="bold")),
                ft.DataColumn(ft.Text("QTDE COTAS", weight="bold")),
                ft.DataColumn(ft.Text("ULTIMO PREÇO", weight="bold")),
                ft.DataColumn(ft.Text("PROVENTOS", weight="bold")),
                ft.DataColumn(ft.Text("RENDIMENTO MÊS", weight="bold")),
                ft.DataColumn(ft.Text("RENDIMENTO ANO", weight="bold")),
                ft.DataColumn(ft.Text("DATA COM", weight="bold")),
                ft.DataColumn(ft.Text("DATA QUE FOI GERADA", weight="bold")),
            ], 
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
            horizontal_lines=ft.border.BorderSide(1, ft.Colors.GREY_300),
        )

        self.texto_resumo = ft.Text(
            value="Total de cotas: 0 | Rendimento acumulado: R$ 0.00 | Total investido: R$ 0.00", 
            size=18, 
            weight="bold",
            color=ft.Colors.BLUE_700
        )

        # Diálogo para editar aporte - melhor organizado
        self.dialog_edit_aporte = ft.AlertDialog(
            title=ft.Text("Editar Aporte"),
            modal=True,
            actions=[],
            content=ft.Column(width=600, height=400, scroll="auto")
        )
        # Diálogo para editar provento
        self.dialog_edit_provento = ft.AlertDialog(
            title=ft.Text("Editar Provento"),
            modal=True,
            actions=[],
            content=ft.Column(width=400, height=200)
        )

        # Diálogo para confirmação de exclusão
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

        self.carregar_dados_excel()
        self.construir_interface()
        self.atualizar_tabelas()

    def construir_interface(self):
        # Formulário para adicionar aporte - melhor organizado em grupos
        grupo_basico = ft.Container(
            content=ft.Column([
                ft.Text("Informações Básicas", size=16, weight="bold", color=ft.Colors.BLUE_700),
                ft.Row([
                    self.fundo_aporte,
                    self.tipo_aporte,
                    self.qtd_aporte,
                    self.preco_aporte,
                ], spacing=10, wrap=True)
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50
        )

        grupo_dividendos = ft.Container(
            content=ft.Column([
                ft.Text("Informações de Dividendos", size=16, weight="bold", color=ft.Colors.GREEN_700),
                ft.Row([
                    self.dy_mes_aporte,
                    self.dy_ano_aporte,
                    self.dy_percentual_aporte,
                ], spacing=10, wrap=True),
                ft.Row([
                    self.dv_ano_aporte,
                    self.dv_mes_aporte,
                ], spacing=10, wrap=True)
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50
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
            bgcolor=ft.Colors.GREY_50
        )

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
            grupo_basico,
            grupo_dividendos,
            grupo_datas,
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

        # Tab de proventos melhorada
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

        tab_rendimentos = ft.Column([
            ft.Text("Resumo de Rendimentos", size=20, weight="bold", color=ft.Colors.PURPLE_800),
            ft.Divider(),
            ft.Container(
                content=self.tabela_rendimentos,
                border=ft.border.all(1, ft.Colors.GREY_300),
                border_radius=10,
                padding=10
            ),
            ft.Container(
                content=self.texto_resumo,
                padding=15,
                border=ft.border.all(2, ft.Colors.BLUE_300),
                border_radius=10,
                bgcolor=ft.Colors.BLUE_50,
                alignment=ft.alignment.center
            )
        ], scroll="auto", expand=True, spacing=15)

        self.tabs = ft.Tabs(
            selected_index=0, 
            tabs=[
                ft.Tab(text="📊 Aportes", content=tab_aportes),
                ft.Tab(text="💰 Proventos", content=tab_proventos),
                ft.Tab(text="📈 Rendimentos", content=tab_rendimentos),
            ], 
            expand=True,
            tab_alignment=ft.TabAlignment.CENTER
        )

        self.page.add(self.tabs, self.dialog_edit_aporte, self.dialog_edit_provento, self.dialog_confirm)

    def carregar_dados_excel(self):
        self.aportes.clear()
        self.proventos.clear()
        if EXCEL_PATH.exists():
            try:
                with pd.ExcelFile(EXCEL_PATH) as reader:
                    try:
                        df_ap = pd.read_excel(reader, sheet_name="Aportes")
                        for _, row in df_ap.iterrows():
                            self.aportes.append({
                                "fundo": str(row["fundo"]),
                                "tipo": str(row.get("tipo", "")),
                                "quantidade": int(row["quantidade"]),
                                "preco": float(row["preco"]),
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
            except Exception as e:
                print("Erro ao carregar Excel:", e)

    def salvar_excel(self):
        try:
            df_ap = pd.DataFrame(self.aportes)
            df_pr = pd.DataFrame(self.proventos)

            # Preparar dados para a aba de Rendimentos
            df_rendimentos = pd.DataFrame(columns=["FII", "QTDE COTAS", "ULTIMO PREÇO", "PROVENTOS", "RENDIMENTO MÊS", "RENDIMENTO ANO", "DATA COM", "DATA QUE FOI GERADA"])
            if self.aportes:
                df_ap_agg = df_ap.groupby("fundo").agg(
                    qtd_total=("quantidade", "sum"),
                    preco_medio=("preco", "last"),
                    dy_mes=("dy_mes", "last"),
                    dy_ano=("dy_ano", "last"),
                    data_com=("data_com", "last")
                ).reset_index()

                df_pr_agg = df_pr.groupby("fundo").agg(
                    provento_total=("valor", "sum")
                ).reset_index()

                df_merged = pd.merge(df_ap_agg, df_pr_agg, on="fundo", how="left")
                df_merged["provento_total"] = df_merged["provento_total"].fillna(0)
                df_merged["rendimento_mes"] = df_merged["qtd_total"] * df_merged["dy_mes"]
                df_merged["rendimento_ano"] = df_merged["qtd_total"] * df_merged["dy_ano"]
                df_merged["data_gerada"] = datetime.now().strftime("%d/%m/%Y")

                df_rendimentos = df_merged[["fundo", "qtd_total", "preco_medio", "provento_total", "rendimento_mes", "rendimento_ano", "data_com", "data_gerada"]]
                df_rendimentos.columns = ["FII", "QTDE COTAS", "ULTIMO PREÇO", "PROVENTOS", "RENDIMENTO MÊS", "RENDIMENTO ANO", "DATA COM", "DATA QUE FOI GERADA"]

            with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
                df_ap.to_excel(writer, sheet_name="Aportes", index=False)
                df_pr.to_excel(writer, sheet_name="Proventos", index=False)
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
        self.tipo_aporte.value = ""
        self.qtd_aporte.value = ""
        self.preco_aporte.value = ""
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
        
        # Validar datas se preenchidas
        if self.data_com_aporte.value and self.data_com_aporte.value.strip():
            try:
                datetime.strptime(self.data_com_aporte.value.strip(), "%d/%m/%Y")
            except ValueError:
                erros.append("Data COM deve estar no formato dd/mm/aaaa")
        
        if self.data_cadastrado_aporte.value and self.data_cadastrado_aporte.value.strip():
            try:
                datetime.strptime(self.data_cadastrado_aporte.value.strip(), "%d/%m/%Y")
            except ValueError:
                erros.append("Data Cadastrado deve estar no formato dd/mm/aaaa")
        
        return erros

    def atualizar_tabelas(self):
        # Atualiza tabela de aportes
        self.tabela_aportes.rows.clear()
        for i, a in enumerate(self.aportes):
            valor_investido = a["quantidade"] * a["preco"]
            
            # Botões de ação organizados
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
                ft.DataCell(ft.Text(a["tipo"])),
                ft.DataCell(ft.Text(str(a["quantidade"]))),
                ft.DataCell(ft.Text(f'R$ {a["preco"]:.2f}')),
                ft.DataCell(ft.Text(f'R$ {valor_investido:.2f}', weight="bold")),
                ft.DataCell(ft.Text(f'{a.get("dy_mes", 0):.2f}')),
                ft.DataCell(ft.Text(f'{a.get("dy_ano", 0):.2f}')),
                ft.DataCell(ft.Text(f'{a.get("dy_percentual", 0):.2f}%')),
                ft.DataCell(ft.Text(f'R$ {a.get("dv_ano", 0):.2f}')),
                ft.DataCell(ft.Text(f'R$ {a.get("dv_mes", 0):.2f}')),
                ft.DataCell(ft.Text(a.get("data_com", ""))),
                ft.DataCell(ft.Text(a.get("data_cadastrado", ""))),
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

        # Atualiza tabela rendimentos
        self.tabela_rendimentos.rows.clear()

        if self.aportes:
            df_ap = pd.DataFrame(self.aportes)
            df_pr = pd.DataFrame(self.proventos) if self.proventos else pd.DataFrame(columns=["fundo", "valor"])

            df_ap_agg = df_ap.groupby("fundo").agg(
                qtd_total=("quantidade", "sum"),
                preco_medio=("preco", "last"),
                dy_mes=("dy_mes", "last"),
                dy_ano=("dy_ano", "last"),
                data_com=("data_com", "last")
            ).reset_index()

            df_pr_agg = df_pr.groupby("fundo").agg(
                provento_total=("valor", "sum")
            ).reset_index()

            df_merged = pd.merge(df_ap_agg, df_pr_agg, on="fundo", how="left")
            df_merged["provento_total"] = df_merged["provento_total"].fillna(0)
            df_merged["rendimento_mes"] = df_merged["qtd_total"] * df_merged["dy_mes"]
            df_merged["rendimento_ano"] = df_merged["qtd_total"] * df_merged["dy_ano"]
            df_merged["data_gerada"] = datetime.now().strftime("%d/%m/%Y")

            total_cotas = df_merged["qtd_total"].sum()
            total_rendimento_acumulado = df_merged["rendimento_mes"].sum() # Usando rendimento do mês para o resumo
            valor_total_investido = (df_merged["qtd_total"] * df_merged["preco_medio"]).sum()

            for _, row in df_merged.iterrows():
                self.tabela_rendimentos.rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(row["fundo"], weight="bold")),
                    ft.DataCell(ft.Text(str(int(row["qtd_total"])))),
                    ft.DataCell(ft.Text(f'R$ {row["preco_medio"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {row["provento_total"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {row["rendimento_mes"]:.2f}')),
                    ft.DataCell(ft.Text(f'R$ {row["rendimento_ano"]:.2f}')),
                    ft.DataCell(ft.Text(row["data_com"])),
                    ft.DataCell(ft.Text(row["data_gerada"])),
                ]))

            self.texto_resumo.value = (
                f"Total de cotas: {int(total_cotas)} | "
                f"Rendimento acumulado: R$ {total_rendimento_acumulado:.2f} | "
                f"Total investido: R$ {valor_total_investido:.2f}"
            )
        else:
            self.texto_resumo.value = "Total de cotas: 0 | Rendimento acumulado: R$ 0.00 | Total investido: R$ 0.00"

        self.page.update()

    def adicionar_aporte(self, e):
        # Validar campos primeiro
        erros = self.validar_campos_aporte()
        if erros:
            self.show_snack(f"Erro: {'; '.join(erros)}", ft.Colors.RED)
            return

        try:
            fundo = self.fundo_aporte.value.strip().upper()
            tipo = self.tipo_aporte.value.strip()
            quantidade = int(self.qtd_aporte.value)
            preco = float(self.preco_aporte.value)
            dy_mes = float(self.dy_mes_aporte.value or 0)
            dy_ano = float(self.dy_ano_aporte.value or 0)
            dy_percentual = float(self.dy_percentual_aporte.value or 0)
            dv_ano = float(self.dv_ano_aporte.value or 0)
            dv_mes = float(self.dv_mes_aporte.value or 0)
            data_com = self.data_com_aporte.value.strip()
            data_cadastrado = self.data_cadastrado_aporte.value.strip()

            self.aportes.append({
                "fundo": fundo,
                "tipo": tipo,
                "quantidade": quantidade,
                "preco": preco,
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

    def abrir_edicao_aporte(self, e):
        i = e.control.data
        self.editando_aporte = i
        ap = self.aportes[i]

        # Campos organizados em grupos para edição
        self.edit_fundo = ft.TextField(label="FII", value=ap["fundo"], expand=True)
        self.edit_tipo = ft.TextField(label="Tipo", value=ap["tipo"], expand=True)
        self.edit_qtd = ft.TextField(label="Nº Cotas", value=str(ap["quantidade"]), keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        self.edit_preco = ft.TextField(label="Valor Cota (R$)", value=f"{ap["preco"]:.2f}", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        
        self.edit_dy_mes = ft.TextField(label="DY Mês", value=f"{ap.get("dy_mes", 0):.2f}", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        self.edit_dy_ano = ft.TextField(label="DY Ano", value=f"{ap.get("dy_ano", 0):.2f}", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        self.edit_dy_percentual = ft.TextField(label="DY %", value=f"{ap.get("dy_percentual", 0):.2f}", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        
        self.edit_dv_ano = ft.TextField(label="Valor DV Ano", value=f"{ap.get("dv_ano", 0):.2f}", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        self.edit_dv_mes = ft.TextField(label="Valor DV Mês", value=f"{ap.get("dv_mes", 0):.2f}", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        
        self.edit_data_com = ft.TextField(label="Data COM", value=ap.get("data_com", ""), hint_text="dd/mm/aaaa", expand=True)
        self.edit_data_cadastrado = ft.TextField(label="Data Cadastrado", value=ap.get("data_cadastrado", ""), hint_text="dd/mm/aaaa", expand=True)

        # Organizar campos em grupos visuais
        grupo_basico_edit = ft.Container(
            content=ft.Column([
                ft.Text("Informações Básicas", size=14, weight="bold", color=ft.Colors.BLUE_700),
                ft.Row([self.edit_fundo, self.edit_tipo], spacing=10),
                ft.Row([self.edit_qtd, self.edit_preco], spacing=10),
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50
        )

        grupo_dividendos_edit = ft.Container(
            content=ft.Column([
                ft.Text("Dividendos", size=14, weight="bold", color=ft.Colors.GREEN_700),
                ft.Row([self.edit_dy_mes, self.edit_dy_ano, self.edit_dy_percentual], spacing=5, expand=True),
                ft.Row([self.edit_dv_ano, self.edit_dv_mes], spacing=5, expand=True),
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50
        )

        grupo_datas_edit = ft.Container(
            content=ft.Column([
                ft.Text("Datas", size=14, weight="bold", color=ft.Colors.ORANGE_700),
                ft.Row([self.edit_data_com, self.edit_data_cadastrado], spacing=10),
            ]),
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=8,
            bgcolor=ft.Colors.GREY_50
        )

        content = ft.Column([
            grupo_basico_edit,
            grupo_dividendos_edit,
            grupo_datas_edit,
        ], spacing=15, tight=True)

        self.dialog_edit_aporte.title.value = f"Editar Aporte: {ap['fundo']}"
        self.dialog_edit_aporte.content = content
        self.dialog_edit_aporte.actions = [
            ft.ElevatedButton(
                "Salvar", 
                on_click=self.salvar_edicao_aporte,
                bgcolor=ft.Colors.GREEN_600,
                color=ft.Colors.WHITE
            ),
            ft.ElevatedButton(
                "Cancelar", 
                on_click=self.cancelar_edicao,
                bgcolor=ft.Colors.GREY_600,
                color=ft.Colors.WHITE
            ),
        ]
        self.dialog_edit_aporte.open = True
        self.page.update()

    def salvar_edicao_aporte(self, e):
        try:
            nova_fundo = self.edit_fundo.value.strip().upper()
            novo_tipo = self.edit_tipo.value.strip()
            nova_qtd = int(self.edit_qtd.value)
            novo_preco = float(self.edit_preco.value)
            novo_dy_mes = float(self.edit_dy_mes.value or 0)
            novo_dy_ano = float(self.edit_dy_ano.value or 0)
            novo_dy_percentual = float(self.edit_dy_percentual.value or 0)
            novo_dv_ano = float(self.edit_dv_ano.value or 0)
            novo_dv_mes = float(self.edit_dv_mes.value or 0)
            nova_data_com = self.edit_data_com.value.strip()
            nova_data_cadastrado = self.edit_data_cadastrado.value.strip()

            if not nova_fundo or nova_qtd <= 0 or novo_preco <= 0:
                raise ValueError("FII, Nº Cotas e Valor Cota são obrigatórios e maiores que zero")

            if nova_data_com:
                try:
                    datetime.strptime(nova_data_com, "%d/%m/%Y")
                except:
                    raise ValueError("Data COM deve estar no formato dd/mm/aaaa")
            if nova_data_cadastrado:
                try:
                    datetime.strptime(nova_data_cadastrado, "%d/%m/%Y")
                except:
                    raise ValueError("Data Cadastrado deve estar no formato dd/mm/aaaa")

            idx = self.editando_aporte
            self.aportes[idx] = {
                "fundo": nova_fundo,
                "tipo": novo_tipo,
                "quantidade": nova_qtd,
                "preco": novo_preco,
                "dy_mes": novo_dy_mes,
                "dy_ano": novo_dy_ano,
                "dy_percentual": novo_dy_percentual,
                "dv_ano": novo_dv_ano,
                "dv_mes": novo_dv_mes,
                "data_com": nova_data_com,
                "data_cadastrado": nova_data_cadastrado,
            }
            self.editando_aporte = None
            self.dialog_edit_aporte.open = False
            self.salvar_excel()
            self.atualizar_tabelas()
            self.show_snack("Aporte editado com sucesso!", ft.Colors.GREEN)
        except Exception as ex:
            self.show_snack(f"Erro ao editar aporte: {ex}", ft.Colors.RED)
        self.page.update()

    def abrir_confirmacao_exclusao_aporte(self, e):
        i = e.control.data
        self.index_para_excluir = i
        self.excluir_tipo = "aporte"
        self.dialog_confirm.content = ft.Text(f"Confirma exclusão do aporte: {self.aportes[i]['fundo']}?")
        self.dialog_confirm.actions = [
            ft.ElevatedButton(
                "Sim", 
                on_click=self.confirmar_exclusao,
                bgcolor=ft.Colors.RED_600,
                color=ft.Colors.WHITE
            ),
            ft.ElevatedButton(
                "Não", 
                on_click=self.cancelar_exclusao,
                bgcolor=ft.Colors.GREY_600,
                color=ft.Colors.WHITE
            ),
        ]
        self.dialog_confirm.open = True
        self.page.update()

    def abrir_confirmacao_exclusao_provento(self, e):
        i = e.control.data
        self.index_para_excluir = i
        self.excluir_tipo = "provento"
        self.dialog_confirm.content = ft.Text(f"Confirma exclusão do provento: {self.proventos[i]['fundo']}?")
        self.dialog_confirm.actions = [
            ft.ElevatedButton(
                "Sim", 
                on_click=self.confirmar_exclusao,
                bgcolor=ft.Colors.RED_600,
                color=ft.Colors.WHITE
            ),
            ft.ElevatedButton(
                "Não", 
                on_click=self.cancelar_exclusao,
                bgcolor=ft.Colors.GREY_600,
                color=ft.Colors.WHITE
            ),
        ]
        self.dialog_confirm.open = True
        self.page.update()

    def confirmar_exclusao(self, e):
        if self.excluir_tipo == "aporte":
            del self.aportes[self.index_para_excluir]
        elif self.excluir_tipo == "provento":
            del self.proventos[self.index_para_excluir]
        self.index_para_excluir = None
        self.excluir_tipo = None
        self.dialog_confirm.open = False
        self.salvar_excel()
        self.atualizar_tabelas()
        self.show_snack("Registro excluído com sucesso!", ft.Colors.ORANGE)
        self.page.update()

    def cancelar_exclusao(self, e):
        self.index_para_excluir = None
        self.excluir_tipo = None
        self.dialog_confirm.open = False
        self.page.update()

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

    def abrir_edicao_provento(self, e):
        i = e.control.data
        self.editando_provento = i
        p = self.proventos[i]
        self.edit_valor = ft.TextField(label="Valor R$", value=f"{p['valor']:.2f}", width=200, keyboard_type=ft.KeyboardType.NUMBER)
        self.edit_data = ft.TextField(label="Data (dd/mm/aaaa)", value=p["data"], width=200, hint_text="dd/mm/aaaa")

        self.dialog_edit_provento.title.value = f"Editar Provento: {p['fundo']}"
        self.dialog_edit_provento.content = ft.Column([
            ft.Text(f"Fundo: {p['fundo']}", size=16, weight="bold"),
            ft.Divider(),
            self.edit_valor,
            self.edit_data,
        ], spacing=10)
        self.dialog_edit_provento.actions = [
            ft.ElevatedButton(
                "Salvar", 
                on_click=self.salvar_edicao_provento,
                bgcolor=ft.Colors.GREEN_600,
                color=ft.Colors.WHITE
            ),
            ft.ElevatedButton(
                "Cancelar", 
                on_click=self.cancelar_edicao,
                bgcolor=ft.Colors.GREY_600,
                color=ft.Colors.WHITE
            ),
        ]
        self.dialog_edit_provento.open = True
        self.page.update()

    def salvar_edicao_provento(self, e):
        try:
            novo_valor = float(self.edit_valor.value)
            nova_data = self.edit_data.value.strip()
            if novo_valor <= 0:
                raise ValueError("Valor deve ser maior que zero")
            try:
                datetime.strptime(nova_data, "%d/%m/%Y")
            except:
                raise ValueError("Data deve estar no formato dd/mm/aaaa")
            idx = self.editando_provento
            self.proventos[idx]["valor"] = novo_valor
            self.proventos[idx]["data"] = nova_data
            self.editando_provento = None
            self.dialog_edit_provento.open = False
            self.salvar_excel()
            self.atualizar_tabelas()
            self.show_snack("Provento editado com sucesso!", ft.Colors.GREEN)
        except Exception as ex:
            self.show_snack(f"Erro ao editar provento: {ex}", ft.Colors.RED)
        self.page.update()

    def cancelar_edicao(self, e):
        self.editando_aporte = None
        self.editando_provento = None
        self.dialog_edit_aporte.open = False
        self.dialog_edit_provento.open = False
        self.page.update()

def main(page: ft.Page):
    ControleFIIsApp(page)

if __name__ == "__main__":
    ft.app(target=main, view=ft.FLET_APP)

