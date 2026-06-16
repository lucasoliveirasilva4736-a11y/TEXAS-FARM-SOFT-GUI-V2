import traceback
import tkinter as tk
import tkinter.messagebox as tkmb
import os
import re
import math

# Bloco Try-Except global para capturar erros caso falte alguma biblioteca
try:
    import customtkinter as ctk
    from tkinter import messagebox, ttk, filedialog
    import pandas as pd
    from openpyxl import load_workbook

    # Configuração visual do sistema
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    # =======================================================
    # MÓDULO 1: SETOR DE PCP (CALCULADORA DE GRADE NA TELA)
    # =======================================================
    class JanelaPCP(ctk.CTkToplevel):
        def __init__(self, master):
            super().__init__(master)
            self.title("TEXAS FARM - Setor PCP")
            self.geometry("1300x950")
            self.protocol("WM_DELETE_WINDOW", self.fechar_e_voltar)

            # Variáveis de controle de estado
            self.caminho_arquivo = None
            self.abas_disponiveis = {}
            self.lotes_na_aba = {}
            self.lista_cores_atual = []
            self.fixado = False
            
            # Grade padrão de tamanhos da fábrica
            self.tamanhos_alvo = ['PP', 'P', 'M', 'G', 'GG', 'XG', 'G1', 'G2', 'G3', 'G4', '2', '4', '6', '8', '10', '12', '14']

            # --- CONSTRUÇÃO DA INTERFACE ---
            self.frame_topo = ctk.CTkFrame(self, fg_color="transparent")
            self.frame_topo.pack(pady=10, fill="x")
            
            self.btn_fixar = ctk.CTkButton(self.frame_topo, text="📌 PINAR JANELA", fg_color="#4a4a4a", command=self.alternar_fixar, width=150)
            self.btn_fixar.pack(side="left", padx=20)
            
            ctk.CTkLabel(self.frame_topo, text="📊 PAINEL DO PCP (CALCULADORA DE GRADE)", font=("Roboto", 20, "bold"), text_color="#3b8ed0").pack(side="left", expand=True)

            self.frame_main = ctk.CTkScrollableFrame(self)
            self.frame_main.pack(fill="both", expand=True, padx=20, pady=10)

            self.frame_excel = ctk.CTkFrame(self.frame_main)
            self.frame_excel.pack(pady=5, fill="x", padx=5)

            self.btn_excel = ctk.CTkButton(self.frame_excel, text="📂 1. LER PLANILHA DE PRODUÇÃO", fg_color="#1D6F42", hover_color="#144d2e", command=self.importar_excel)
            self.btn_excel.pack(pady=10)

            # Seletores de Filtro e Abas
            self.frame_seletores = ctk.CTkFrame(self.frame_excel, fg_color="transparent")
            self.frame_seletores.pack(fill="x", padx=10, pady=5)

            self.menu_abas = ctk.CTkOptionMenu(self.frame_seletores, values=["-"], command=self.processar_aba_selecionada)
            self.menu_abas.grid(row=0, column=1, padx=5)

            self.entry_busca = ctk.CTkEntry(self.frame_seletores, placeholder_text="🔎 Buscar Lote...", width=150)
            self.entry_busca.grid(row=0, column=2, padx=5)
            self.entry_busca.bind("<KeyRelease>", self.filtrar_lotes)

            self.menu_lotes = ctk.CTkOptionMenu(self.frame_seletores, values=["-"], command=self.carregar_dados_lote, width=200)
            self.menu_lotes.grid(row=0, column=3, padx=5)

            # Entradas de peso da grade
            self.frame_grades = ctk.CTkFrame(self.frame_main, fg_color="transparent")
            self.frame_grades.pack(pady=10)
            self.entradas_grade = {}
            self.setup_grade_ui()

            self.btn_calc = ctk.CTkButton(self.frame_main, text="2. CALCULAR GRADE EXATA", height=50, command=self.executar_grade_pcp, fg_color="#3b8ed0")
            self.btn_calc.pack(pady=20)

            # Tabela de Resultados
            self.frame_res = ctk.CTkFrame(self.frame_main)
            self.frame_res.pack(pady=5, fill="both", expand=True)

            self.label_soma_total = ctk.CTkLabel(self.frame_res, text="TOTAL DO LOTE: 0 PEÇAS", font=("Roboto", 18, "bold"))
            self.label_soma_total.pack(pady=10)

            # --- FORMATAÇÃO DA TABELA (FONTE MAIOR) ---
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview", font=("Roboto", 12), rowheight=30, background="#2b2b2b", foreground="white")
            style.configure("Treeview.Heading", font=("Roboto", 13, "bold"), background="#3b8ed0", foreground="white")
            style.map("Treeview", background=[('selected', '#1f538d')])

            self.colunas = ("REF", "COR") + tuple(self.tamanhos_alvo) + ("TOTAL",)
            self.tabela = ttk.Treeview(self.frame_res, columns=self.colunas, show="headings")
            for col in self.colunas:
                self.tabela.heading(col, text=col)
                largura = 120 if col in ["REF", "COR"] else 55
                self.tabela.column(col, width=largura, anchor="center")
            self.tabela.pack(fill="both", expand=True)

        # --- FUNÇÕES DO PCP ---
        def alternar_fixar(self): 
            self.fixado = not self.fixado
            self.attributes("-topmost", self.fixado)
            self.btn_fixar.configure(fg_color="#2fa572" if self.fixado else "#4a4a4a")

        def fechar_e_voltar(self): 
            self.master.deiconify()
            self.destroy()

        def setup_grade_ui(self):
            ctk.CTkLabel(self.frame_grades, text="ADULTO / PLUS", font=("Roboto", 12, "bold"), text_color="#3b8ed0").grid(row=0, column=0, columnspan=10, sticky="w", padx=5, pady=(5,2))
            for i, tam in enumerate(self.tamanhos_alvo[:10]):
                ctk.CTkLabel(self.frame_grades, text=tam, font=("Roboto", 12)).grid(row=1, column=i, padx=6)
                self.entradas_grade[tam] = ctk.CTkEntry(self.frame_grades, width=55, justify="center")
                self.entradas_grade[tam].insert(0, "0")
                self.entradas_grade[tam].grid(row=2, column=i, padx=2, pady=2)

            ctk.CTkLabel(self.frame_grades, text="INFANTIL (2-14)", font=("Roboto", 12, "bold"), text_color="#3b8ed0").grid(row=3, column=0, columnspan=10, sticky="w", padx=5, pady=(10,2))
            for i, tam in enumerate(self.tamanhos_alvo[10:]):
                ctk.CTkLabel(self.frame_grades, text=tam, font=("Roboto", 12)).grid(row=4, column=i, padx=6)
                self.entradas_grade[tam] = ctk.CTkEntry(self.frame_grades, width=55, justify="center")
                self.entradas_grade[tam].insert(0, "0")
                self.entradas_grade[tam].grid(row=5, column=i, padx=2, pady=2)

        def importar_excel(self):
            c = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
            if not c: return
            self.caminho_arquivo = c
            try:
                self.abas_disponiveis = pd.read_excel(c, sheet_name=None, header=None)
                n = list(self.abas_disponiveis.keys())
                self.menu_abas.configure(values=n if n else ["-"])
                if n:
                    self.menu_abas.set(n[0])
                    self.processar_lotes_da_planilha(n[0])
                messagebox.showinfo("Sucesso", "Planilha carregada e dados lidos com sucesso!")
            except Exception as e: 
                messagebox.showerror("Erro", f"Erro ao ler arquivo: {e}")

        def processar_aba_selecionada(self, n_aba):
            self.processar_lotes_da_planilha(n_aba)
            
        def processar_lotes_da_planilha(self, n_aba):
            df = self.abas_disponiveis[n_aba]
            self.lotes_na_aba = {}
            lote_atual = None
            dados_l = []

            for i, linha in df.iterrows():
                l_str = " ".join([str(val) for val in linha.values if pd.notna(val)]).upper()
                
                if "LOTE:" in l_str:
                    lote_atual = str(linha[0]).strip().upper() if "LOTE:" in str(linha[0]).upper() else str(linha[1]).strip().upper()
                    
                    # Inicia a grade zerada
                    grade_lote = {t: 0 for t in self.tamanhos_alvo}
                    
                    # SCANNER VERTICAL TURBINADO: Vasculha as próximas 20 linhas e várias colunas
                    for r in range(i, min(i + 20, len(df))):
                        for col_idx in range(2, min(7, len(df.columns))): # Lê as colunas C, D, E, F
                            celula = str(df.iloc[r, col_idx]).upper()
                            for t in self.tamanhos_alvo:
                                # Aceita PP=2, PP-2, PP:2 com ou sem espaço
                                m = re.search(rf"\b{t}\b\s*[=:\-]\s*(\d+)", celula)
                                if m and grade_lote[t] == 0: 
                                    grade_lote[t] = int(m.group(1))

                    self.lotes_na_aba[lote_atual] = {"itens": [], "aba": n_aba, "grade": grade_lote}
                    dados_l = self.lotes_na_aba[lote_atual]["itens"]
                
                # Coleta dados das cores (ignora cabeçalhos internos)
                elif lote_atual and not pd.isna(linha[1]) and not pd.isna(linha[2]):
                    cor = str(linha[1]).upper()
                    if "COR" in cor or "TOTAL" in cor or "QUANT" in str(linha[2]).upper(): continue
                    try:
                        try:
                            r_val = str(linha[0]).replace(',', '.').strip()
                            rolos_qte = int(round(float(r_val))) if r_val.upper() not in ["NAN", "NONE", ""] else 0
                        except:
                            rolos_qte = 0
                            
                        dados_l.append({
                            'rolos': rolos_qte,
                            'cor': str(linha[1]).strip(),
                            'ref': str(linha[3]).strip() if pd.notna(linha[3]) else "Sem Ref",
                            
                            # O SEGREDO REVELADO: O round() arredonda os decimais escondidos do Excel!
                            'total': int(round(float(str(linha[2]).replace(',','.'))))
                        })
                    except: continue

            n = list(self.lotes_na_aba.keys())
            self.menu_lotes.configure(values=n if n else ["-"])
            if n:
                self.menu_lotes.set(n[0])
                self.carregar_dados_lote(n[0])
            else:
                self.menu_lotes.set("-")
                
       #-----------------------------------------
        def carregar_dados_lote(self, n):
            if n == "-": return
            
            lote = self.lotes_na_aba.get(n)
            if not lote: return

            self.lista_cores_atual = lote["itens"]
            
            # NOVO: Injeta os valores escaneados nas caixinhas da tela automaticamente
            for tam in self.tamanhos_alvo:
                valor = lote.get("grade", {}).get(tam, 0)
                self.entradas_grade[tam].delete(0, 'end')
                self.entradas_grade[tam].insert(0, str(valor))
                
            self.tabela.delete(*self.tabela.get_children())
            soma = sum([i['total'] for i in self.lista_cores_atual])
            self.label_soma_total.configure(text=f"TOTAL DO LOTE: {soma} PEÇAS")

        def filtrar_lotes(self, e):
            t = self.entry_busca.get().upper()
            f = [l for l in self.lotes_na_aba.keys() if t in l]
            self.menu_lotes.configure(values=f if f else ["-"])
            if f: 
                self.menu_lotes.set(f[0])
                self.carregar_dados_lote(f[0])
            else:
                self.menu_lotes.set("-")

        def executar_grade_pcp(self):
            if not self.lista_cores_atual: return
            try:
                pesos = {k: int(v.get().strip() or 0) for k, v in self.entradas_grade.items()}
                soma_pesos = sum(pesos.values())
                if soma_pesos == 0: return

                self.tabela.delete(*self.tabela.get_children())
                soma_geral = 0

                for item in self.lista_cores_atual:
                    total_cor = item['total']
                    refs = [r.strip() for r in str(item['ref']).split('/') if r.strip()]
                    refs = refs if refs else ["Sem Ref"]
                    
                    grade_cor = {t: 0 for t in self.tamanhos_alvo}
                    fracoes = {t: 0.0 for t in self.tamanhos_alvo}
                    
                    # 1. Calcula a distribuição proporcional apenas para tamanhos que possuem peso
                    for t, p in pesos.items():
                        if p > 0:
                            valor_exato = (total_cor * p) / soma_pesos
                            grade_cor[t] = math.floor(valor_exato)
                            fracoes[t] = valor_exato - math.floor(valor_exato)
                    
                    # 2. SOLUÇÃO MATEMÁTICA CORRIGIDA: Loop para não perder peças
                    sobra = total_cor - sum(grade_cor.values())
                    if sobra > 0:
                        tams_ativos = [t for t, p in pesos.items() if p > 0]
                        if tams_ativos:
                            tams_ord = sorted(tams_ativos, key=lambda x: fracoes[x], reverse=True)
                            for i in range(sobra): 
                                grade_cor[tams_ord[i % len(tams_ord)]] += 1
                    
                    # 3. Faz o Rateio (Split) exato caso a cor atenda a múltiplas referências
                    for idx, ref in enumerate(refs):
                        div = len(refs)
                        tot_ref = {t: (v // div) + (1 if idx < (v % div) else 0) for t, v in grade_cor.items()}
                        soma_ref_final = sum(tot_ref.values())
                        
                        self.tabela.insert("", "end", values=[ref, item['cor']] + [tot_ref[t] for t in self.tamanhos_alvo] + [soma_ref_final])
                        soma_geral += soma_ref_final
                
                self.label_soma_total.configure(text=f"TOTAL CALCULADO: {soma_geral} PEÇAS")
            except Exception as e: 
                messagebox.showerror("Erro de Cálculo", traceback.format_exc())


    # =======================================================
    # MÓDULO 2: SETOR ALMOXARIFADO (ESTOQUE)
    # =======================================================
    class JanelaEstoque(ctk.CTkToplevel):
        def __init__(self, master):
            super().__init__(master)
            self.title("TEXAS FARM - Setor Almoxarifado/Estoque")
            self.geometry("1100x800")
            self.protocol("WM_DELETE_WINDOW", self.fechar_e_voltar)
            
            # Diretório de gravação do histórico da rede Texas Farm
            self.arquivo_memoria = r"Z:\TXF\Software_PCP_TexasFarm\historico_lotes_mt.txt"
            self.caminho_arquivo = None
            self.lotes_processados_agora = set()
            self.novo_estoque_calculado = {}

            # --- CONSTRUÇÃO DA INTERFACE ---
            self.frame_topo = ctk.CTkFrame(self, fg_color="transparent")
            self.frame_topo.pack(pady=10, fill="x")
            ctk.CTkLabel(self.frame_topo, text="📦 PAINEL DO ALMOXARIFADO", font=("Roboto", 20, "bold"), text_color="#d08d3b").pack()

            self.frame_controles = ctk.CTkFrame(self)
            self.frame_controles.pack(pady=5, padx=20, fill="x")
            ctk.CTkButton(self.frame_controles, text="📂 1. ABRIR PLANILHA", command=self.importar_excel, fg_color="#1D6F42", font=("Roboto", 13, "bold")).pack(side="left", padx=5, pady=10)
            ctk.CTkButton(self.frame_controles, text="🔄 2. ESTOQUE ATUAL", command=self.ver_estoque, fg_color="#2fa572", font=("Roboto", 13, "bold")).pack(side="left", padx=5, pady=10)
            ctk.CTkButton(self.frame_controles, text="🧵 3. ATUALIZAR ESTOQUE", command=self.calcular_saldo, fg_color="#1f538d", font=("Roboto", 13, "bold")).pack(side="left", padx=5, pady=10)
            self.btn_confirmar = ctk.CTkButton(self.frame_controles, text="✔️ 4. CONFIRMAR E SALVAR NO EXCEL", command=self.salvar_historico_e_excel, fg_color="#8b0000", state="disabled", font=("Roboto", 13, "bold"))
            self.btn_confirmar.pack(side="left", padx=5, pady=10)

            self.frame_seletores_est = ctk.CTkFrame(self)
            self.frame_seletores_est.pack(pady=5, padx=20, fill="x")
            
            ctk.CTkLabel(self.frame_seletores_est, text="Aba do Lote:", font=("Roboto", 12)).grid(row=0, column=0, padx=5, pady=5)
            self.menu_aba_lote = ctk.CTkOptionMenu(self.frame_seletores_est, values=["-"])
            self.menu_aba_lote.grid(row=0, column=1, padx=5, pady=5)
            
            ctk.CTkLabel(self.frame_seletores_est, text="Aba do Estoque:", font=("Roboto", 12)).grid(row=0, column=2, padx=5, pady=5)
            self.menu_aba_est = ctk.CTkOptionMenu(self.frame_seletores_est, values=["-"])
            self.menu_aba_est.grid(row=0, column=3, padx=5, pady=5)

            self.lbl_status = ctk.CTkLabel(self, text="Selecione a planilha...", text_color="gray", font=("Roboto", 14))
            self.lbl_status.pack(pady=5)

            self.frame_tabela = ctk.CTkFrame(self)
            self.frame_tabela.pack(fill="both", expand=True, padx=20, pady=10)

            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview", font=("Roboto", 13), rowheight=30, background="#2b2b2b", foreground="white")
            style.configure("Treeview.Heading", font=("Roboto", 14, "bold"), background="#d08d3b", foreground="white")

            self.tabela = ttk.Treeview(self.frame_tabela, show="headings")
            self.tabela.pack(fill="both", expand=True)

        # --- FUNÇÕES DO ALMOXARIFADO ---
        def fechar_e_voltar(self): 
            self.master.deiconify()
            self.destroy()

        def importar_excel(self):
            c = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
            self.caminho_arquivo = c
            if c: 
                self.lbl_status.configure(text=f"Carregado: {c.split('/')[-1]}", text_color="white")
                try:
                    xl = pd.ExcelFile(c)
                    nomes_abas = xl.sheet_names
                    self.menu_aba_lote.configure(values=nomes_abas if nomes_abas else ["-"])
                    self.menu_aba_est.configure(values=nomes_abas if nomes_abas else ["-"])
                    
                    # Seleção inteligente das abas por palavra-chave
                    sel_lote = next((s for s in nomes_abas if any(x in s.upper() for x in ["LOTE", "PROGRAMA", "PRODU", "CORTES"])), nomes_abas[0])
                    sel_est = next((s for s in nomes_abas if any(x in s.upper() for x in ["REFER", "ROLO", "ESTOQUE", "CONTROLE", "MOLETOM", "MALHA"])), nomes_abas[0])
                    
                    self.menu_aba_lote.set(sel_lote)
                    self.menu_aba_est.set(sel_est)
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao ler as abas do arquivo: {e}")

        def extrair_estoque(self):
            if not self.caminho_arquivo: return None
            aba_est = self.menu_aba_est.get()
            if aba_est == "-": return None
            try:
                df = pd.read_excel(self.caminho_arquivo, sheet_name=aba_est, header=None)
                estoque = {}
                cc, cq = -1, -1
                
                # Detecta as colunas de cor e quantidade
                for r_idx, row in df.iterrows():
                    for c_idx, val in enumerate(row):
                        v = str(val).strip().upper()
                        if v == "COR" or v == "CORES": cc = c_idx
                        if "QTE" in v or "ROLO" in v or "QTD" in v or "QUANT" in v: cq = c_idx
                    if cc != -1 and cq != -1: break
                    
                for _, row in df.iterrows():
                    cor_bruta = str(row[cc]).upper()
                    cor_limpa = re.sub(r'\(.*?\)', '', cor_bruta).strip()
                    if cor_limpa in ["COR", "NAN", "NONE", "", "TOTAL"] or pd.isna(row[cc]): continue
                    try:
                        v = str(row[cq]).replace(',', '.').strip()
                        if v.upper() not in ["NAN", "NONE", ""]: estoque[cor_limpa] = float(v)
                    except: pass
                return estoque
            except PermissionError:
                messagebox.showerror("Erro", "O Excel está aberto! Feche o arquivo para prosseguir.")
                return None
            except Exception as e: 
                messagebox.showerror("Erro", f"Falha ao ler estoque: {e}")
                return None

        def ver_estoque(self):
            est = self.extrair_estoque()
            if not est: return
            
            self.tabela.delete(*self.tabela.get_children())
            self.tabela["columns"] = ("COR", "QTD SALVA NO EXCEL")
            for c in self.tabela["columns"]: 
                self.tabela.heading(c, text=c)
                self.tabela.column(c, anchor="center")
                
            total = 0
            for cor, q in est.items(): 
                self.tabela.insert("", "end", values=(cor, q))
                total += q
            self.tabela.insert("", "end", values=("TOTAL GERAL", total))
            self.lbl_status.configure(text="Estoque atualizado em tempo real da planilha!", text_color="#2fa572")

        def calcular_saldo(self):
            est = self.extrair_estoque()
            if not est: return
            
            aba_lote = self.menu_aba_lote.get()
            if aba_lote == "-": return
                
            df_lote = pd.read_excel(self.caminho_arquivo, sheet_name=aba_lote, header=None)
            cons, ign = {}, set()
            
            # Verifica o txt de memória para evitar dupla baixa de estoque
            try:
                os.makedirs(os.path.dirname(self.arquivo_memoria), exist_ok=True)
                if os.path.exists(self.arquivo_memoria):
                    with open(self.arquivo_memoria, "r", encoding="utf-8") as f: ign = set(f.read().splitlines())
            except: pass
            
            l_at, self.lotes_processados_agora, cores_erro = None, set(), []
            for _, linha in df_lote.iterrows():
                l_str = " ".join([str(val).upper() for val in linha.values if pd.notna(val)]).upper()
                if "LOTE:" in l_str:
                    for cel in linha.values:
                        if pd.notna(cel) and "LOTE:" in str(cel).upper(): 
                            l_at = str(cel).strip().upper()
                            break
                    continue
                if "TOTAL" in l_str: l_at = None; continue
                
                # Consumo
                if l_at and l_at.startswith("MT") and l_at not in ign:
                    cor_bruta = str(linha[1]).upper()
                    cor_limpa = re.sub(r'\(.*?\)', '', cor_bruta).strip()
                    if cor_limpa in ["COR", "NAN", "NONE", "", "TOTAL"] or pd.isna(linha[1]): continue
                    if cor_limpa not in est:
                        if cor_limpa not in cores_erro: cores_erro.append(cor_limpa)
                        continue

                    self.lotes_processados_agora.add(l_at)
                    try:
                        r = float(str(linha[0]).replace(',', '.').strip())
                        if not pd.isna(r) and r > 0: cons[cor_limpa] = cons.get(cor_limpa, 0.0) + r
                    except: pass

            if cores_erro: messagebox.showwarning("Cores Ausentes", f"Ignoradas por não estarem no Estoque:\n\n" + "\n".join(cores_erro))

            self.tabela.delete(*self.tabela.get_children())
            self.tabela["columns"] = ("COR", "INI", "CONS", "NOVO SALDO")
            for c in self.tabela["columns"]: 
                self.tabela.heading(c, text=c)
                self.tabela.column(c, anchor="center")
                
            d_t = []
            self.novo_estoque_calculado = {}
            for cor in sorted(set(list(est.keys()) + list(cons.keys()))):
                q_i, q_c = est.get(cor, 0.0), cons.get(cor, 0.0)
                if q_i > 0 or q_c > 0:
                    novo_saldo = q_i - q_c
                    self.tabela.insert("", "end", values=(cor, q_i, q_c, novo_saldo))
                    d_t.append(novo_saldo)
                    self.novo_estoque_calculado[cor] = novo_saldo
            
            if self.novo_estoque_calculado:
                self.tabela.insert("", "end", values=("TOTAL GERAL", "", "", sum(d_t)))
                self.btn_confirmar.configure(state="normal")
                self.lbl_status.configure(text="Saldos prontos! Clique em CONFIRMAR para gravar no Excel.", text_color="#d08d3b")

        def salvar_historico_e_excel(self):
            if not self.lotes_processados_agora or not self.novo_estoque_calculado: return
            if messagebox.askyesno("Confirmar e Salvar", "Deseja gravar o novo estoque na planilha e salvar o histórico TXT?"):
                try:
                    wb = load_workbook(self.caminho_arquivo)
                    aba_est = self.menu_aba_est.get()
                    if aba_est not in wb.sheetnames: raise Exception("Aba de Estoque selecionada inválida.")
                    
                    ws = wb[aba_est]
                    cc, cq, header_row = -1, -1, 1
                    
                    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
                        for c_idx, val in enumerate(row):
                            v = str(val).strip().upper() if val else ""
                            if v == "COR" or v == "CORES": cc = c_idx
                            if "QTE" in v or "ROLO" in v or "QTD" in v or "QUANT" in v: cq = c_idx
                        if cc != -1 and cq != -1: header_row = r_idx + 1; break
                            
                    if cc == -1 or cq == -1: raise Exception("Colunas COR e QTD não identificadas.")

                    soma_total_excel = 0
                    linha_total = None

                    # Atualização das Células físicas
                    for row in ws.iter_rows(min_row=header_row+1):
                        cor_bruta = str(row[cc].value).upper() if row[cc].value else ""
                        cor_limpa = re.sub(r'\(.*?\)', '', cor_bruta).strip()

                        if cor_limpa in self.novo_estoque_calculado: row[cq].value = self.novo_estoque_calculado[cor_limpa]
                        if "TOTAL" in cor_limpa: linha_total = row
                        elif row[cq].value is not None and str(row[cq].value).replace('.','',1).isdigit(): soma_total_excel += float(row[cq].value)

                    if linha_total: linha_total[cq].value = soma_total_excel

                    wb.save(self.caminho_arquivo)

                    # Salva no arquivo de registro (Memória)
                    caminho_dir = os.path.dirname(self.arquivo_memoria)
                    if not os.path.exists(caminho_dir):
                        try: os.makedirs(caminho_dir, exist_ok=True)
                        except Exception: self.arquivo_memoria = "historico_lotes_mt_local.txt"

                    with open(self.arquivo_memoria, "a", encoding="utf-8") as f:
                        for l in self.lotes_processados_agora: f.write(l + "\n")
                    
                    self.btn_confirmar.configure(state="disabled")
                    self.lotes_processados_agora.clear()
                    self.novo_estoque_calculado.clear()
                    
                    self.lbl_status.configure(text="Estoque, Total e Histórico salvos com sucesso!", text_color="#1D6F42")
                    messagebox.showinfo("Sucesso", "Estoque gravado no Excel com TOTAL atualizado!")
                    self.ver_estoque()

                except PermissionError: messagebox.showerror("Erro Crítico", "O Excel está ABERTO! Você precisa fechar o arquivo antes de tentar gravar.")
                except Exception as e: messagebox.showerror("Erro de Gravação", traceback.format_exc())

    # =======================================================
    # LAUNCHER HUB
    # =======================================================
    class LauncherTexasFarm(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("TEXAS FARM - ERP Hub")
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"500x400+{(sw-500)//2}+{(sh-400)//2}")
            ctk.CTkLabel(self, text="TEXAS FARM", font=("Roboto", 28, "bold")).pack(pady=(40, 5))
            ctk.CTkButton(self, text="⚙️ PCP", height=70, width=300, font=("Roboto", 15, "bold"), command=self.abrir_pcp).pack(pady=10)
            ctk.CTkButton(self, text="📦 ESTOQUE DE ROLOS", height=70, width=300, font=("Roboto", 15, "bold"), fg_color="#d08d3b", command=self.abrir_estoque).pack(pady=10)
        
        def abrir_pcp(self): self.withdraw(); JanelaPCP(self)
        def abrir_estoque(self): self.withdraw(); JanelaEstoque(self)

    if __name__ == "__main__": 
        LauncherTexasFarm().mainloop()

# Tratamento geral caso as bibliotecas não estejam instaladas
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    tkmb.showerror("Erro Fatal", f"Ocorreu um erro ao iniciar o sistema.\nCertifique-se de que as bibliotecas estão instaladas.\n\nDetalhes do erro:\n{traceback.format_exc()}")
