import os
import threading
from datetime import datetime

import pandas as pd
import numpy as np
import customtkinter as ctk
from tkinter import filedialog

try:
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


# =========================================================================
# ETAPA 2 - CLASSIFICAÇÃO DE CONTAS
# =========================================================================

def _validar_colunas(df: pd.DataFrame, obrigatorias: list, nome_arquivo: str):
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' está sem as colunas obrigatórias: {faltando}"
        )


def _aplicar_cores(caminho_arquivo: str, col_acao_nome: str):
    """Aplica cores condicionais na coluna de Ação para facilitar leitura visual."""
    if not OPENPYXL_OK:
        return
    import openpyxl

    wb = openpyxl.load_workbook(caminho_arquivo)
    ws = wb.active

    # Descobre em que coluna está "Ação"
    header = [cell.value for cell in ws[1]]
    if col_acao_nome not in header:
        wb.close()
        return
    col_idx = header.index(col_acao_nome) + 1
    col_letra = get_column_letter(col_idx)

    fill_bloqueio = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")   # vermelho claro
    fill_alerta = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")     # amarelo
    fill_manter = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")     # verde claro

    for row in range(2, ws.max_row + 1):
        valor = ws[f"{col_letra}{row}"].value or ""
        if "BLOQUEAR" in valor:
            ws[f"{col_letra}{row}"].fill = fill_bloqueio
        elif "ALERTA" in valor:
            ws[f"{col_letra}{row}"].fill = fill_alerta
        elif valor == "Manter":
            ws[f"{col_letra}{row}"].fill = fill_manter

    ws.freeze_panes = "A2"
    wb.save(caminho_arquivo)
    wb.close()


def processar_etapa2_core(path_contas: str, path_transferencias: str) -> str:

    # ================= COLUNAS =================
    COL_ID = 'IGA ID'
    COL_CATEGORIA = 'Categoria'
    COL_STATUS = 'Status IGA'
    COL_CRIACAO = 'WhenCreated'
    COL_LOGON = 'LastLogonTimeStamp'

    COL_TRANSFERIDO = 'Transferido?'
    COL_DATA_TRANSF = 'Data da transferência'
    COL_ACAO = 'Ação'
    COL_MOTIVO = 'Motivo Ação'
    COL_OBS = 'Observações Adicionais'

    COL_TRANSF_DATE = 'Transfer Date'
    COL_TRANSF_ID = 'Network ID'

    data_str = datetime.now().strftime('%Y-%m-%d')
    ARQ_SAIDA = f'Relatorio_Contas_FINAL_{data_str}.xlsx'

    # ================= CATEGORIAS =================
    EXCECOES = {'Chave de Processo', 'Conta PAM'}
    DXC = {'Conta A- DXC', 'Conta D- DXC'}
    OUTROS = {'EMPREGADO', 'CONTRATADO', 'Chave A- VALE ou C0', 'Usuário Não Identificado'}

    # ================= CARGA =================
    df = pd.read_excel(path_contas)
    df_transf = pd.read_excel(path_transferencias)

    _validar_colunas(df, [COL_ID, COL_CATEGORIA, COL_STATUS, COL_CRIACAO, COL_LOGON],
                      os.path.basename(path_contas))

    linhas_originais = len(df)

    # ================= SAÍDA PADRÃO =================
    for col, padrao in [
        (COL_TRANSFERIDO, 'Não'),
        (COL_DATA_TRANSF, pd.NaT),
        (COL_ACAO, 'Manter'),
        (COL_MOTIVO, ''),
        (COL_OBS, ''),
    ]:
        if col not in df.columns:
            df[col] = padrao

    # ================= DATAS =================
    df[COL_CRIACAO] = pd.to_datetime(df[COL_CRIACAO], errors='coerce')
    df[COL_LOGON] = pd.to_datetime(df[COL_LOGON], errors='coerce')

    # ================= MAPA DE TRANSFERÊNCIAS (SEM MERGE) =================
    if COL_TRANSF_ID in df_transf.columns and COL_ID not in df_transf.columns:
        df_transf = df_transf.rename(columns={COL_TRANSF_ID: COL_ID})

    _validar_colunas(df_transf, [COL_ID, COL_TRANSF_DATE], os.path.basename(path_transferencias))

    df_transf['DataTransferencia'] = pd.to_datetime(df_transf[COL_TRANSF_DATE], errors='coerce')

    mapa_transf = (
        df_transf
        .dropna(subset=[COL_ID])
        .sort_values('DataTransferencia')
        .groupby(COL_ID)['DataTransferencia']
        .last()
    )

    df[COL_DATA_TRANSF] = df[COL_ID].map(mapa_transf)
    df[COL_TRANSFERIDO] = df[COL_DATA_TRANSF].notna().map({True: 'Sim', False: 'Não'})

    # ================= MÁSCARAS BASE =================
    mask_excecao = df[COL_CATEGORIA].isin(EXCECOES)
    mask_dxc = df[COL_CATEGORIA].isin(DXC)
    mask_outros = df[COL_CATEGORIA].isin(OUTROS)
    mask_avaliar = ~mask_excecao

    hoje = datetime.now()

    # ================= CONDIÇÕES CALCULADAS DE FORMA INDEPENDENTE =================
    # Importante: nenhuma condição abaixo depende do resultado de outra.
    # Isso evita o bug em que uma conta com ALERTA de transferência nunca
    # era reavaliada pelo critério de tempo (e vice-versa).

    # --- B) Inatividade IGA ---
    cond_iga = mask_avaliar & (df[COL_STATUS] == 'INACTIVE')

    # --- C) Transferência (bloqueio e alerta) ---
    cond_transfer_base = (
        mask_avaliar &
        df[COL_TRANSFERIDO].eq('Sim') &
        df[COL_DATA_TRANSF].notna() &
        df[COL_CRIACAO].notna()
    )
    dias_ct = (df[COL_DATA_TRANSF] - df[COL_CRIACAO]).dt.days

    cond_bloq_transf = cond_transfer_base & (df[COL_CRIACAO] <= df[COL_DATA_TRANSF]) & (dias_ct > 7)
    cond_alerta_transf = cond_transfer_base & (df[COL_CRIACAO] <= df[COL_DATA_TRANSF]) & (dias_ct >= 0) & (dias_ct <= 7)

    # --- D) Tempo (LastLogon ou Criação) ---
    data_base = df[COL_LOGON].combine_first(df[COL_CRIACAO])
    dias = (hoje - data_base).dt.days

    cond_dxc_tempo = mask_avaliar & mask_dxc & data_base.notna() & (dias > 180)
    cond_outros_tempo = mask_avaliar & mask_outros & data_base.notna() & (dias > 90)
    cond_tempo = cond_dxc_tempo | cond_outros_tempo

    # ================= PRIORIDADE FINAL =================
    # Ordem (da mais fraca pra mais forte, quem vem depois sobrescreve):
    # Manter -> Alerta Transferência -> Tempo -> Inatividade IGA -> Bloqueio Transferência
    # O bloqueio por Transferência é SEMPRE a prioridade máxima: se ele se aplica,
    # a Ação final é sempre "BLOQUEAR por Transferência", não importa o que mais bateu.
    df[COL_ACAO] = 'Manter'
    df[COL_MOTIVO] = ''

    # 1) Alerta de transferência
    df.loc[cond_alerta_transf, COL_ACAO] = 'ALERTA: Verificar Transferência'
    df.loc[cond_alerta_transf, COL_MOTIVO] = 'Criação próxima à data de transferência (possível delay)'

    # 2) Tempo (sobrescreve alerta, se também se aplicar)
    df.loc[cond_dxc_tempo, COL_ACAO] = 'BLOQUEAR por TEMPO (DXC > 180 dias)'
    df.loc[cond_dxc_tempo & df[COL_LOGON].isna(), COL_MOTIVO] = 'Nunca logado. Conta criada há ' + dias.astype(str) + ' dias'
    df.loc[cond_dxc_tempo & df[COL_LOGON].notna(), COL_MOTIVO] = dias.astype(str) + ' dias sem logon'

    df.loc[cond_outros_tempo, COL_ACAO] = 'BLOQUEAR por TEMPO (Outros > 90 dias)'
    df.loc[cond_outros_tempo & df[COL_LOGON].isna(), COL_MOTIVO] = 'Nunca logado. Conta criada há ' + dias.astype(str) + ' dias'
    df.loc[cond_outros_tempo & df[COL_LOGON].notna(), COL_MOTIVO] = dias.astype(str) + ' dias sem logon'

    # 3) Inatividade IGA (sobrescreve tempo/alerta, se também se aplicar)
    df.loc[cond_iga, COL_ACAO] = 'BLOQUEAR por Inatividade'
    df.loc[cond_iga, COL_MOTIVO] = 'Status IGA = INACTIVE'

    # 4) Bloqueio por transferência: PRIORIDADE MÁXIMA, aplicada sempre por último
    df.loc[cond_bloq_transf, COL_ACAO] = 'BLOQUEAR por Transferência'
    df.loc[cond_bloq_transf, COL_MOTIVO] = 'Conta criada antes da transferência'

    # ================= OBSERVAÇÕES ADICIONAIS (SINALIZAÇÃO) =================
    # Aqui não mudamos a Ação final, só avisamos quando a linha também batia
    # em outro critério que acabou sendo "engolido" pela prioridade acima.
    obs = pd.Series([''] * len(df), index=df.index)

    def _add_obs(mask, texto):
        nonlocal obs
        obs = obs.where(~mask, (obs + np.where(obs.eq(''), '', ' | ') + texto))

    e_tempo_final = df[COL_ACAO].isin(['BLOQUEAR por TEMPO (DXC > 180 dias)',
                                        'BLOQUEAR por TEMPO (Outros > 90 dias)'])
    e_iga_final = df[COL_ACAO] == 'BLOQUEAR por Inatividade'
    e_bloq_transf_final = df[COL_ACAO] == 'BLOQUEAR por Transferência'
    e_alerta_final = df[COL_ACAO] == 'ALERTA: Verificar Transferência'

    # Caso específico pedido: a conta tinha alerta de transferência (delay <=7 dias),
    # mas NÃO virou bloqueio por transferência -> se o motivo final for Tempo ou IGA,
    # deixa claro que ela será bloqueada mesmo assim, só que por outro critério.
    _add_obs(
        cond_alerta_transf & e_tempo_final,
        'Havia ALERTA de transferência (delay ≤ 7 dias), mas a conta será bloqueada por TEMPO, não por transferência'
    )
    _add_obs(
        cond_alerta_transf & e_iga_final,
        'Havia ALERTA de transferência (delay ≤ 7 dias), mas a conta será bloqueada por INATIVIDADE (Status IGA), não por transferência'
    )

    # Demais sobreposições (informativas)
    _add_obs(cond_bloq_transf & ~e_bloq_transf_final,
             'Também se enquadra em BLOQUEIO por Transferência (criada antes, >7 dias)')  # não deve ocorrer, mantido por segurança

    _add_obs(cond_tempo & ~e_tempo_final & ~e_bloq_transf_final,
             'Também se enquadra em BLOQUEIO por TEMPO')

    _add_obs(cond_iga & ~e_iga_final & ~e_bloq_transf_final,
             'Também está INACTIVE no IGA')

    _add_obs(cond_alerta_transf & ~e_alerta_final & ~e_tempo_final & ~e_iga_final & ~e_bloq_transf_final,
             'Também se enquadra em ALERTA de transferência (delay ≤ 7 dias)')

    df[COL_OBS] = obs

    # ================= GARANTIA FINAL =================
    assert len(df) == linhas_originais, 'ERRO: quantidade de linhas alterada'

    # ================= SALVAR =================
    caminho = os.path.join(os.path.dirname(path_contas), ARQ_SAIDA)
    df.to_excel(caminho, index=False)
    _aplicar_cores(caminho, COL_ACAO)

    resumo = df[COL_ACAO].value_counts(dropna=False)
    qtd_sinalizadas = int((df[COL_OBS] != '').sum())

    return (
        f'✅ SUCESSO!\n\n'
        f'Arquivo gerado:\n{caminho}\n\n'
        f'Resumo das Ações:\n{resumo.to_string()}\n\n'
        f'⚠️ Linhas com observação adicional (mais de um critério bateu): {qtd_sinalizadas}'
    )


# =========================================================================
# ETAPA 3 - COMPARAÇÃO COM O MÊS ANTERIOR
# =========================================================================

def comparar_etapa3_core(path_atual: str, path_anterior: str) -> str:

    COL_ID = 'IGA ID'
    COL_ACAO = 'Ação'
    COL_ULTIMA_ACAO = 'Última Ação'
    COL_MUDANCA = 'Status da Mudança'

    data_str = datetime.now().strftime('%Y-%m-%d')
    ARQ_SAIDA = f'Relatorio_Comparativo_Etapa3_{data_str}.xlsx'

    df_atual = pd.read_excel(path_atual)
    df_anterior = pd.read_excel(path_anterior)

    _validar_colunas(df_atual, [COL_ID, COL_ACAO], os.path.basename(path_atual))
    _validar_colunas(df_anterior, [COL_ID, COL_ACAO], os.path.basename(path_anterior))

    linhas_originais = len(df_atual)

    # Mapa ID -> Ação do mês anterior (se houver duplicidade de ID, fica a última linha)
    mapa_acao_anterior = (
        df_anterior
        .dropna(subset=[COL_ID])
        .drop_duplicates(subset=[COL_ID], keep='last')
        .set_index(COL_ID)[COL_ACAO]
    )

    df_atual[COL_ULTIMA_ACAO] = df_atual[COL_ID].map(mapa_acao_anterior)

    estava_na_ultima = df_atual[COL_ULTIMA_ACAO].notna()
    df_atual.loc[~estava_na_ultima, COL_ULTIMA_ACAO] = 'Não estava na última'

    # Sinaliza se a ação mudou de um mês pro outro
    mudou = estava_na_ultima & (df_atual[COL_ACAO] != df_atual[COL_ULTIMA_ACAO])
    igual = estava_na_ultima & (df_atual[COL_ACAO] == df_atual[COL_ULTIMA_ACAO])

    df_atual[COL_MUDANCA] = 'Nova conta (não estava na última)'
    df_atual.loc[igual, COL_MUDANCA] = 'Sem alteração'
    df_atual.loc[mudou, COL_MUDANCA] = (
        df_atual.loc[mudou, COL_ULTIMA_ACAO] + ' → ' + df_atual.loc[mudou, COL_ACAO]
    )

    assert len(df_atual) == linhas_originais, 'ERRO: quantidade de linhas alterada'

    caminho = os.path.join(os.path.dirname(path_atual), ARQ_SAIDA)
    df_atual.to_excel(caminho, index=False)
    _aplicar_cores(caminho, COL_ACAO)

    novas = int((~estava_na_ultima).sum())
    mudaram = int(mudou.sum())
    sem_alteracao = int(igual.sum())

    return (
        f'✅ SUCESSO!\n\n'
        f'Arquivo gerado:\n{caminho}\n\n'
        f'Total de linhas: {linhas_originais}\n'
        f'Novas contas (não estavam na última): {novas}\n'
        f'Mudaram de ação: {mudaram}\n'
        f'Sem alteração: {sem_alteracao}'
    )


# =========================================================================
# GUI
# =========================================================================
APP_TITLE = "Processador de Contas - Etapa 2 e Etapa 3"
DEFAULT_GEOMETRY = "760x520"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry(DEFAULT_GEOMETRY)
        self.minsize(700, 460)

        self.tabview = ctk.CTkTabview(self, corner_radius=16)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=20)

        self.tab_etapa2 = self.tabview.add("Etapa 2 - Classificação")
        self.tab_etapa3 = self.tabview.add("Etapa 3 - Comparativo")

        self._build_etapa2(self.tab_etapa2)
        self._build_etapa3(self.tab_etapa3)

    # ---------------- ETAPA 2 ----------------
    def _build_etapa2(self, parent):
        self.path_contas = ctk.StringVar(value="")
        self.path_transferencias = ctk.StringVar(value="")

        title = ctk.CTkLabel(parent, text="Processar Relatório (Etapa 2)",
                              font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(anchor="w", padx=6, pady=(10, 6))

        subtitle = ctk.CTkLabel(parent, text="Escolha os arquivos .xlsx de Contas e Transferências.",
                                 font=ctk.CTkFont(size=13))
        subtitle.pack(anchor="w", padx=6, pady=(0, 14))

        self._file_picker(parent, "Arquivo de Contas (.xlsx)", self.path_contas,
                           lambda: self._pick_file(self.path_contas, "Contas"))
        self._file_picker(parent, "Arquivo de Transferências (.xlsx)", self.path_transferencias,
                           lambda: self._pick_file(self.path_transferencias, "Transferências"))

        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.pack(fill="x", padx=6, pady=(14, 8))

        self.btn_processar = ctk.CTkButton(actions, text="Processar", height=42,
                                            font=ctk.CTkFont(size=14, weight="bold"),
                                            command=self._on_processar_etapa2)
        self.btn_processar.pack(side="left")

        self.progress2 = ctk.CTkProgressBar(actions, mode="indeterminate", height=12)
        self.progress2.pack(side="left", fill="x", expand=True, padx=(14, 0))
        self.progress2.stop()

        self.status_box2 = self._status_box(parent)
        self._set_status(self.status_box2, "Selecione os dois arquivos e clique em Processar.")

    # ---------------- ETAPA 3 ----------------
    def _build_etapa3(self, parent):
        self.path_atual = ctk.StringVar(value="")
        self.path_anterior = ctk.StringVar(value="")

        title = ctk.CTkLabel(parent, text="Comparar com o mês anterior (Etapa 3)",
                              font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(anchor="w", padx=6, pady=(10, 6))

        subtitle = ctk.CTkLabel(
            parent,
            text="Escolha o relatório do mês ATUAL e o do mês ANTERIOR (ambos já processados na Etapa 2).",
            font=ctk.CTkFont(size=13), wraplength=680, justify="left")
        subtitle.pack(anchor="w", padx=6, pady=(0, 14))

        self._file_picker(parent, "Relatório do mês ATUAL (.xlsx)", self.path_atual,
                           lambda: self._pick_file(self.path_atual, "Mês Atual"))
        self._file_picker(parent, "Relatório do mês ANTERIOR (.xlsx)", self.path_anterior,
                           lambda: self._pick_file(self.path_anterior, "Mês Anterior"))

        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.pack(fill="x", padx=6, pady=(14, 8))

        self.btn_comparar = ctk.CTkButton(actions, text="Comparar", height=42,
                                           font=ctk.CTkFont(size=14, weight="bold"),
                                           command=self._on_comparar_etapa3)
        self.btn_comparar.pack(side="left")

        self.progress3 = ctk.CTkProgressBar(actions, mode="indeterminate", height=12)
        self.progress3.pack(side="left", fill="x", expand=True, padx=(14, 0))
        self.progress3.stop()

        self.status_box3 = self._status_box(parent)
        self._set_status(self.status_box3, "Selecione os dois arquivos e clique em Comparar.")

    # ---------------- HELPERS DE UI ----------------
    def _file_picker(self, parent, label, var, command):
        row = ctk.CTkFrame(parent, corner_radius=12)
        row.pack(fill="x", padx=6, pady=6)

        lbl = ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=13, weight="bold"))
        lbl.pack(anchor="w", padx=12, pady=(10, 4))

        inner = ctk.CTkFrame(row, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=(0, 12))

        entry = ctk.CTkEntry(inner, textvariable=var, placeholder_text="Selecione um arquivo...")
        entry.pack(side="left", fill="x", expand=True)

        btn = ctk.CTkButton(inner, text="Selecionar", width=120, command=command)
        btn.pack(side="left", padx=(10, 0))

    def _status_box(self, parent):
        status_frame = ctk.CTkFrame(parent, corner_radius=12)
        status_frame.pack(fill="both", expand=True, padx=6, pady=(10, 6))

        status_label = ctk.CTkLabel(status_frame, text="Status", font=ctk.CTkFont(size=14, weight="bold"))
        status_label.pack(anchor="w", padx=14, pady=(12, 6))

        box = ctk.CTkTextbox(status_frame, corner_radius=10)
        box.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        return box

    def _set_status(self, box, text):
        box.delete("1.0", "end")
        box.insert("1.0", text)

    def _pick_file(self, var, label):
        path = filedialog.askopenfilename(title=f"Selecione o arquivo de {label}",
                                           filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            var.set(path)

    # ---------------- AÇÕES ETAPA 2 ----------------
    def _on_processar_etapa2(self):
        pc = self.path_contas.get().strip()
        pt = self.path_transferencias.get().strip()
        if not pc or not os.path.exists(pc):
            self._set_status(self.status_box2, "❌ Selecione um arquivo válido de Contas.")
            return
        if not pt or not os.path.exists(pt):
            self._set_status(self.status_box2, "❌ Selecione um arquivo válido de Transferências.")
            return

        self.btn_processar.configure(state="disabled", text="Processando...")
        self.progress2.start()
        self._set_status(self.status_box2, "🔄 Processando...")

        threading.Thread(target=self._run_etapa2, args=(pc, pt), daemon=True).start()

    def _run_etapa2(self, pc, pt):
        try:
            resultado = processar_etapa2_core(pc, pt)
        except Exception as e:
            resultado = f"❌ ERRO inesperado:\n{e}"
        self.after(0, lambda: self._finish_etapa2(resultado))

    def _finish_etapa2(self, resultado):
        self.progress2.stop()
        self.btn_processar.configure(state="normal", text="Processar")
        self._set_status(self.status_box2, resultado)

    # ---------------- AÇÕES ETAPA 3 ----------------
    def _on_comparar_etapa3(self):
        pa = self.path_atual.get().strip()
        pant = self.path_anterior.get().strip()
        if not pa or not os.path.exists(pa):
            self._set_status(self.status_box3, "❌ Selecione um arquivo válido do mês Atual.")
            return
        if not pant or not os.path.exists(pant):
            self._set_status(self.status_box3, "❌ Selecione um arquivo válido do mês Anterior.")
            return

        self.btn_comparar.configure(state="disabled", text="Comparando...")
        self.progress3.start()
        self._set_status(self.status_box3, "🔄 Comparando...")

        threading.Thread(target=self._run_etapa3, args=(pa, pant), daemon=True).start()

    def _run_etapa3(self, pa, pant):
        try:
            resultado = comparar_etapa3_core(pa, pant)
        except Exception as e:
            resultado = f"❌ ERRO inesperado:\n{e}"
        self.after(0, lambda: self._finish_etapa3(resultado))

    def _finish_etapa3(self, resultado):
        self.progress3.stop()
        self.btn_comparar.configure(state="normal", text="Comparar")
        self._set_status(self.status_box3, resultado)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
