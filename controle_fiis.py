# APP faz o controle dos meus fundos, add ou exclui e cria excel atualizado 
# controle_fiis
# APP funcional OK

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
        self.page.window_width = 900
        self.page.window_height = 700

        # Campos de aporte
        self.fundo_aporte = ft.TextField(label="Fundo (ex: KNCR11)", width=200)
        self.qtd_aporte = ft.TextField(label="Quantidade", width=120, keyboard_type=ft.KeyboardType.NUMBER)
        self.preco_aporte = ft.TextField(label="Preço por cota (R$)", width=150, keyboard_type=ft.KeyboardType.NUMBER)

        # Campos de provento
        self.fundo_provento = ft.TextField(label="Fundo", width=200)
        self.valor_provento = ft.TextField(label="Rendimento por cota (R$)", width=180, keyboard_type=ft.KeyboardType.NUMBER)

        # Tabelas
        self.tabela_aportes = ft.DataTable(columns=[
            ft.DataColumn(ft.Text("Fundo")),
            ft.DataColumn(ft.Text("Quantidade")),
            ft.DataColumn(ft.Text("Preço Unitário")),
            ft.DataColumn(ft.Text("Total Investido")),
            ft.DataColumn(ft.Text("Editar")),
            ft.DataColumn(ft.Text("Excluir")),
        ], rows=[])

        self.tabela_proventos = ft.DataTable(columns=[
            ft.DataColumn(ft.Text("Fundo")),
            ft.DataColumn(ft.Text("Valor R$")),
            ft.DataColumn(ft.Text("Data")),
            ft.DataColumn(ft.Text("Editar")),
            ft.DataColumn(ft.Text("Excluir")),
        ], rows=[])

        # Rendimento tabela e texto resumo
        self.tabela_rendimentos = ft.DataTable(columns=[
            ft.DataColumn(ft.Text("Fundo")),
            ft.DataColumn(ft.Text("Qtde Cotas")),
            ft.DataColumn(ft.Text("Último Preço")),
            ft.DataColumn(ft.Text("Proventos (R$)")),
            ft.DataColumn(ft.Text("Rendimento Total (R$)")),
        ], rows=[])
        self.texto_resumo = ft.Text(value="Total de cotas: 0 | Rendimento acumulado: R$ 0.00", size=16, weight="bold")

        # Dialog edição
        self.dialog = ft.AlertDialog(title=ft.Text("Editar Registro"), modal=True, actions=[
            ft.ElevatedButton("Salvar", on_click=self.salvar_edicao),
            ft.ElevatedButton("Cancelar", on_click=self.cancelar_edicao),
        ])

        # Dialog confirmação exclusão
        self.dialog_confirm = ft.AlertDialog(
            title=ft.Text("Confirmação"),
            content=ft.Text("Tem certeza que deseja excluir este registro?"),
            modal=True,
            actions=[
                ft.ElevatedButton("Sim", on_click=self.confirmar_exclusao),
                ft.ElevatedButton("Não", on_click=self.cancelar_exclusao),
            ],
        )

        self.editando_aporte = None
        self.editando_provento = None
        self.excluir_aporte_index = None
        self.excluir_provento_index = None

        # SnackBar inicial
        self.page.snack_bar = ft.SnackBar(content=ft.Text(""))

        self.carregar_dados_excel()
        self.construir_interface()
        self.atualizar_tabelas()

    def construir_interface(self):
        layout = ft.Column([
            ft.Text("Cadastro de Aportes", size=20, weight="bold"),
            ft.Row([
                self.fundo_aporte,
                self.qtd_aporte,
                self.preco_aporte,
                ft.ElevatedButton("Adicionar Aporte", on_click=self.adicionar_aporte)
            ]),
            self.tabela_aportes,
            ft.Divider(height=2, thickness=2, color=ft.Colors.BLACK),

            ft.Text("Cadastro de Proventos", size=20, weight="bold"),
            ft.Row([
                self.fundo_provento,
                self.valor_provento,
                ft.ElevatedButton("Adicionar Provento", on_click=self.adicionar_provento)
            ]),
            self.tabela_proventos,
            ft.Divider(height=2, thickness=2, color=ft.Colors.BLACK),

            ft.Text("Rendimentos Acumulados", size=20, weight="bold"),
            self.tabela_rendimentos,
            self.texto_resumo,
        ], scroll="auto", expand=True)

        self.page.add(layout, self.dialog, self.dialog_confirm)

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
                                "quantidade": int(row["quantidade"]),
                                "preco": float(row["preco"])
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
            with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
                df_ap.to_excel(writer, sheet_name="Aportes", index=False)
                df_pr.to_excel(writer, sheet_name="Proventos", index=False)
        except Exception as e:
            self.show_snack(f"Erro ao salvar Excel: {e}", ft.Colors.RED)

    def show_snack(self, msg, color):
        self.page.snack_bar.content.value = msg
        self.page.snack_bar.bgcolor = color
        self.page.snack_bar.open = True
        self.page.update()

    def atualizar_tabelas(self):
        self.tabela_aportes.rows.clear()
        for i, a in enumerate(self.aportes):
            total = a["quantidade"] * a["preco"]
            self.tabela_aportes.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(a["fundo"])),
                ft.DataCell(ft.Text(str(a["quantidade"]))),
                ft.DataCell(ft.Text(f'R$ {a["preco"]:.2f}')),
                ft.DataCell(ft.Text(f'R$ {total:.2f}')),
                ft.DataCell(ft.IconButton(ft.Icons.EDIT, tooltip="Editar", data=i, on_click=self.abrir_edicao_aporte)),
                ft.DataCell(ft.IconButton(ft.Icons.DELETE, tooltip="Excluir", data=i, on_click=self.confirmar_delecao_aporte)),
            ]))

        self.tabela_proventos.rows.clear()
        for i, p in enumerate(self.proventos):
            self.tabela_proventos.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(p["fundo"])),
                ft.DataCell(ft.Text(f'R$ {p["valor"]:.2f}')),
                ft.DataCell(ft.Text(p["data"])),
                ft.DataCell(ft.IconButton(ft.Icons.EDIT, tooltip="Editar", data=i, on_click=self.abrir_edicao_provento)),
                ft.DataCell(ft.IconButton(ft.Icons.DELETE, tooltip="Excluir", data=i, on_click=self.confirmar_delecao_provento)),
            ]))

        # Atualiza rendimentos
        df_ap = pd.DataFrame(self.aportes)
        df_pr = pd.DataFrame(self.proventos)

        self.tabela_rendimentos.rows.clear()

        if not df_ap.empty:
            df_ap_agg = df_ap.groupby("fundo").agg(
                qtd_total=("quantidade", "sum"),
                preco_medio=("preco", "last")
            ).reset_index()

            if not df_pr.empty:
                df_pr_agg = df_pr.groupby("fundo").agg(
                    provento_total=("valor", "sum")
                ).reset_index()
            else:
                df_pr_agg = pd.DataFrame(columns=["fundo", "provento_total"])

            df_merged = pd.merge(df_ap_agg, df_pr_agg, on="fundo", how="left")
            df_merged["provento_total"] = df_merged["provento_total"].fillna(0)
            df_merged["rendimento_total"] = df_merged["qtd_total"] * df_merged["provento_total"]

            total_cotas = df_merged["qtd_total"].sum()
            total_rendimento = df_merged["rendimento_total"].sum()

            for _, row in df_merged.iterrows():
                self.tabela_rendimentos.rows.append(
                    ft.DataRow(cells=[
                        ft.DataCell(ft.Text(row["fundo"])),
                        ft.DataCell(ft.Text(str(row["qtd_total"]))),
                        ft.DataCell(ft.Text(f'R$ {row["preco_medio"]:.2f}')),
                        ft.DataCell(ft.Text(f'R$ {row["provento_total"]:.2f}')),
                        ft.DataCell(ft.Text(f'R$ {row["rendimento_total"]:.2f}')),
                    ])
                )
            valor_total_investido = (df_merged["qtd_total"] * df_merged["preco_medio"]).sum()
            self.texto_resumo.value = (
                f"Total de cotas: {total_cotas} | "
                f"Rendimento acumulado: R$ {total_rendimento:.2f} | "
                f"Total investido: R$ {valor_total_investido:.2f}"
                    )
        else:
            self.texto_resumo.value = "Total de cotas: 0 | Rendimento acumulado: R$ 0.00 | Total investido: R$ 0.00"

        self.page.update()

    def adicionar_aporte(self, e):
        try:
            fundo = self.fundo_aporte.value.strip().upper()
            quantidade = int(self.qtd_aporte.value)
            preco = float(self.preco_aporte.value)
            if not fundo or quantidade <= 0 or preco <= 0:
                raise ValueError("Preencha todos os campos corretamente")
            self.aportes.append({"fundo": fundo, "quantidade": quantidade, "preco": preco})
            self.fundo_aporte.value = ""
            self.qtd_aporte.value = ""
            self.preco_aporte.value = ""
            self.salvar_excel()
            self.atualizar_tabelas()
            self.show_snack(f"Aporte adicionado: {fundo}", ft.Colors.GREEN)
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

    # ... O restante do código com edição, exclusão e confirmações permanece igual ao último que te enviei ...

    def abrir_edicao_aporte(self, e):
        i = e.control.data
        self.editando_aporte = i
        a = self.aportes[i]
        self.edit_qtd = ft.TextField(label="Quantidade", value=str(a["quantidade"]), width=200)
        self.edit_preco = ft.TextField(label="Preço por cota (R$)", value=f"{a['preco']:.2f}", width=200)

        self.dialog.title.value = f"Editar Aporte: {a['fundo']}"
        self.dialog.content = ft.Column([
            ft.Text(f"Fundo: {a['fundo']}"),
            self.edit_qtd,
            self.edit_preco,
        ])
        self.dialog.open = True
        self.page.update()

    def abrir_edicao_provento(self, e):
        i = e.control.data
        self.editando_provento = i
        p = self.proventos[i]
        self.edit_valor = ft.TextField(label="Valor R$ ", value=f"{p['valor']:.2f}", width=200)
        self.edit_data = ft.TextField(label="Data (dd/mm/aaaa)", value=p["data"], width=200)

        self.dialog.title.value = f"Editar Provento: {p['fundo']}"
        self.dialog.content = ft.Column([
            ft.Text(f"Fundo: {p['fundo']}"),
            self.edit_valor,
            self.edit_data,
        ])
        self.dialog.open = True
        self.page.update()

    def salvar_edicao(self, e):
        if self.editando_aporte is not None:
            try:
                nova_qtd = int(self.edit_qtd.value)
                novo_preco = float(self.edit_preco.value)
                if nova_qtd <= 0 or novo_preco <= 0:
                    raise ValueError("Valores devem ser maiores que zero")
                self.aportes[self.editando_aporte]["quantidade"] = nova_qtd
                self.aportes[self.editando_aporte]["preco"] = novo_preco
                self.editando_aporte = None
                self.dialog.open = False
                self.salvar_excel()
                self.atualizar_tabelas()
                self.show_snack("Aporte editado com sucesso!", ft.Colors.GREEN)
            except Exception as ex:
                self.show_snack(f"Erro ao editar aporte: {ex}", ft.Colors.RED)

        elif self.editando_provento is not None:
            try:
                novo_valor = float(self.edit_valor.value)
                nova_data = self.edit_data.value.strip()
                if novo_valor <= 0:
                    raise ValueError("Valor deve ser maior que zero")
                try:
                    datetime.strptime(nova_data, "%d/%m/%Y")
                except:
                    raise ValueError("Data deve estar no formato dd/mm/aaaa")
                self.proventos[self.editando_provento]["valor"] = novo_valor
                self.proventos[self.editando_provento]["data"] = nova_data
                self.editando_provento = None
                self.dialog.open = False
                self.salvar_excel()
                self.atualizar_tabelas()
                self.show_snack("Provento editado com sucesso!", ft.Colors.GREEN)
            except Exception as ex:
                self.show_snack(f"Erro ao editar provento: {ex}", ft.Colors.RED)

        self.page.update()

    def cancelar_edicao(self, e):
        self.editando_aporte = None
        self.editando_provento = None
        self.dialog.open = False
        self.page.update()

    def confirmar_delecao_aporte(self, e):
        self.excluir_aporte_index = e.control.data
        self.dialog_confirm.title.value = "Confirmar exclusão de aporte"
        self.dialog_confirm.content.value = "Tem certeza que deseja excluir este aporte?"
        self.dialog_confirm.open = True
        self.page.update()

    def confirmar_delecao_provento(self, e):
        self.excluir_provento_index = e.control.data
        self.dialog_confirm.title.value = "Confirmar exclusão de provento"
        self.dialog_confirm.content.value = "Tem certeza que deseja excluir este provento?"
        self.dialog_confirm.open = True
        self.page.update()

    def confirmar_exclusao(self, e):
        if self.excluir_aporte_index is not None:
            del self.aportes[self.excluir_aporte_index]
            self.excluir_aporte_index = None
            self.show_snack("Aporte removido", ft.Colors.RED)
        elif self.excluir_provento_index is not None:
            del self.proventos[self.excluir_provento_index]
            self.excluir_provento_index = None
            self.show_snack("Provento removido", ft.Colors.RED)
        self.dialog_confirm.open = False
        self.salvar_excel()
        self.atualizar_tabelas()
        self.page.update()

    def cancelar_exclusao(self, e):
        self.excluir_aporte_index = None
        self.excluir_provento_index = None
        self.dialog_confirm.open = False
        self.page.update()


def main(page: ft.Page):
    ControleFIIsApp(page)

ft.app(target=main, view=ft.FLET_APP)
