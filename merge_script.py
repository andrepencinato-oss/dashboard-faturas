import pandas as pd
import numpy as np
import os
import glob

def get_uf_from_cep(cep_str):
    if pd.isna(cep_str): return 'NÃO ENCONTRADO'
    cep = str(cep_str).strip()
    import re
    cep = re.sub(r'\D', '', cep)
    if len(cep) < 8:
        cep = cep.zfill(8)
    if not cep: return 'NÃO ENCONTRADO'
    
    prefix = int(cep[:2])
    if 1 <= prefix <= 19: return 'SP'
    if 20 <= prefix <= 28: return 'RJ'
    if prefix == 29: return 'ES'
    if 30 <= prefix <= 39: return 'MG'
    if 40 <= prefix <= 48: return 'BA'
    if prefix == 49: return 'SE'
    if 50 <= prefix <= 56: return 'PE'
    if prefix == 57: return 'AL'
    if prefix == 58: return 'PB'
    if prefix == 59: return 'RN'
    if 60 <= prefix <= 63: return 'CE'
    if prefix == 64: return 'PI'
    if prefix == 65: return 'MA'
    if 66 <= prefix <= 68:
        if cep.startswith('689'): return 'AP'
        return 'PA'
    if prefix == 69:
        if cep.startswith('693'): return 'RR'
        if cep.startswith('699'): return 'AC'
        return 'AM'
    if 70 <= prefix <= 72: return 'DF'
    if prefix == 73:
        if int(cep[:3]) <= 733: return 'DF'
        return 'GO'
    if 74 <= prefix <= 76:
        if int(cep[:3]) >= 768: return 'RO'
        return 'GO'
    if prefix == 77: return 'TO'
    if prefix == 78:
        if cep.startswith('789'): return 'RO'
        return 'MT'
    if prefix == 79: return 'MS'
    if 80 <= prefix <= 87: return 'PR'
    if 88 <= prefix <= 89: return 'SC'
    if 90 <= prefix <= 99: return 'RS'
    return 'NÃO ENCONTRADO'

def consolidate_data(faturas_file=None, coletas_file=None):
    print("Iniciando a leitura dos arquivos...")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Se os arquivos não foram passados pela interface do Streamlit, busca na pasta local
    is_local_mode = False
    if faturas_file is None or coletas_file is None:
        is_local_mode = True
        faturas_files = [f for f in glob.glob(os.path.join(script_dir, 'Resumo todas faturas e-recoli*.xlsx')) if not os.path.basename(f).startswith('~$')]
        coletas_files = [f for f in glob.glob(os.path.join(script_dir, 'Tracking Coletas*.xlsx')) if not os.path.basename(f).startswith('~$')]
        
        if not faturas_files: raise FileNotFoundError("Não encontrei a planilha 'Resumo todas faturas e-recoli.xlsx'.")
        if not coletas_files: raise FileNotFoundError("Não encontrei a planilha de 'Tracking Coletas'.")
        
        faturas_file = max(faturas_files, key=os.path.getmtime)
        coletas_file = max(coletas_files, key=os.path.getmtime)
        
    # 1. Carregar a planilha de Coletas
    print("Lendo Coletas...")
    df_coletas = pd.read_excel(coletas_file, sheet_name='Coletas')
    
    # Padronizar nomes de colunas no Coletas para facilitar a busca (removendo espaos em branco)
    # Procurar colunas de NFD, NFS e outras necessárias pelo nome aproximado
    col_nfd = [c for c in df_coletas.columns if 'NFD' in str(c).upper()][0]
    col_nfs = [c for c in df_coletas.columns if 'SA' in str(c).upper() and 'DA' in str(c).upper() and 'NF' in str(c).upper()][0]
    col_transp = [c for c in df_coletas.columns if 'TRANSPORTADORA' in str(c).upper()][0]
    col_status = [c for c in df_coletas.columns if 'STATUS' in str(c).upper()][0]
    col_data_coleta = [c for c in df_coletas.columns if 'DATA COLETA' in str(c).upper()][0]
    try:
        col_ticket = [c for c in df_coletas.columns if 'TICKET' in str(c).upper()][0]
    except IndexError:
        col_ticket = None
    try:
        col_motivo = [c for c in df_coletas.columns if 'MOTIVO' in str(c).upper()][0]
    except IndexError:
        col_motivo = None
    
    # Garantir que as colunas de chave sejam strings para facilitar o match
    df_coletas['KEY_NFD'] = df_coletas[col_nfd].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df_coletas['KEY_NFS'] = df_coletas[col_nfs].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    
    # 2. Carregar o arquivo de Faturas
    print("Lendo as abas de faturas...")
    xl_faturas = pd.ExcelFile(faturas_file)
    abas_fatura = [aba for aba in xl_faturas.sheet_names if 'fatura' in aba.lower()]
    
    # 2.1 Tentar extrair as datas de vencimento da aba Resumo
    dict_vencimentos = {}
    if 'Resumo' in xl_faturas.sheet_names:
        try:
            df_resumo = xl_faturas.parse('Resumo', header=3)
            col_fatura_resumo = [c for c in df_resumo.columns if 'FATURA' in str(c).upper()][0]
            col_vencimento_resumo = [c for c in df_resumo.columns if 'VENCIMENTO' in str(c).upper()][0]
            
            for _, row in df_resumo.iterrows():
                f_val = str(row[col_fatura_resumo]).strip()
                if f_val.endswith('.0'):
                    f_val = f_val[:-2]
                v_val = row[col_vencimento_resumo]
                if pd.notna(v_val):
                    try:
                        v_str = pd.to_datetime(v_val).strftime('%d/%m/%Y')
                    except:
                        v_str = str(v_val)[:10]
                    dict_vencimentos[f_val] = v_str
        except Exception as e:
            print(f"Não foi possível ler os vencimentos da aba Resumo: {e}")
            
    resultados = []
    
    for aba in abas_fatura:
        print(f"Processando aba: {aba}")
        
        # Ler a aba sem header
        df_fatura = xl_faturas.parse(aba, header=None)
        
        # Encontrar a linha que contem o cabecalho (deve ter 'Nota Fiscal')
        header_row_idx = -1
        col_nf_idx = -1
        
        for idx, row in df_fatura.iterrows():
            for c_idx, cell in enumerate(row):
                if 'NOTA FISCAL' in str(cell).upper():
                    header_row_idx = idx
                    col_nf_idx = c_idx
                    break
            if header_row_idx != -1:
                break
                
        if header_row_idx == -1:
            print(f"  Aviso: Coluna 'Nota Fiscal' não encontrada na aba {aba}. Ignorando.")
            continue
            
        # Definir as colunas usando a linha encontrada
        df_fatura.columns = df_fatura.iloc[header_row_idx]
        # Pegar os dados a partir da linha seguinte
        df_fatura = df_fatura.iloc[header_row_idx+1:].reset_index(drop=True)
        
        col_nf_fatura = df_fatura.columns[col_nf_idx]
        
        # Identificar a coluna de Valor do Frete inteligentemente
        col_valor_name = None
        for col in df_fatura.columns:
            col_upper = str(col).upper()
            if 'VALOR' in col_upper and 'FRETE' in col_upper:
                col_valor_name = col
                break
        
        if not col_valor_name:
            for col in df_fatura.columns:
                if 'VALOR' in str(col).upper():
                    col_valor_name = col
                    break
                    
        col_cep_name = None
        for col in df_fatura.columns:
            if 'CEP' in str(col).upper():
                col_cep_name = col
                break
                
        col_uf_name = None
        for col in df_fatura.columns:
            if str(col).strip().upper() == 'UF':
                col_uf_name = col
                break
        
        # Extrair o nmero da fatura do nome da aba (ex: 'Fatura_197571' -> '197571')
        num_fatura = aba.lower().replace('fatura_', '').replace('fatura', '').strip()
        
        # Iterar sobre as linhas da fatura
        for idx, row in df_fatura.iterrows():
            nf_val = str(row[col_nf_fatura]).strip()
            if pd.isna(row[col_nf_fatura]) or nf_val == 'nan' or nf_val == '':
                continue
            
            # Remover '.0' se foi lido como float
            if nf_val.endswith('.0'):
                nf_val = nf_val[:-2]
                
            # Tratar Valor do Frete
            valor_frete = 0.0
            if col_valor_name is not None:
                val = row[col_valor_name]
                if pd.notna(val):
                    if isinstance(val, (int, float)):
                        valor_frete = float(val)
                    else:
                        val_str = str(val).upper().replace('R$', '').replace(' ', '').strip()
                        if val_str:
                            if ',' in val_str and '.' in val_str:
                                val_str = val_str.replace('.', '')
                            val_str = val_str.replace(',', '.')
                            try:
                                valor_frete = float(val_str)
                            except ValueError:
                                pass
                                
            # Tratar CEP e UF
            cep_val = 'NÃO ENCONTRADO'
            uf_val = 'NÃO ENCONTRADO'
            
            if col_cep_name is not None:
                c_val = row[col_cep_name]
                if pd.notna(c_val) and str(c_val).strip() != '':
                    cep_val = str(c_val).strip()
                    if cep_val.endswith('.0'): cep_val = cep_val[:-2]
                    uf_val = get_uf_from_cep(cep_val)
                    
            if col_uf_name is not None:
                u_val = row[col_uf_name]
                if pd.notna(u_val) and str(u_val).strip() != '':
                    uf_val = str(u_val).strip().upper()
                
            # Buscar no dataframe de coletas: tenta achar em NFD ou NFS
            match_nfd = df_coletas[df_coletas['KEY_NFD'] == nf_val]
            match_nfs = df_coletas[df_coletas['KEY_NFS'] == nf_val]
            
            # Priorizar match_nfd, senao usar match_nfs
            match_df = None
            if not match_nfd.empty:
                match_df = match_nfd.iloc[0]
            elif not match_nfs.empty:
                match_df = match_nfs.iloc[0]
            
            vencimento_fatura = dict_vencimentos.get(str(num_fatura).strip(), 'NÃO ENCONTRADO')

            if match_df is not None:
                resultados.append({
                    'NUMERO DA FATURA': num_fatura,
                    'DATA DE VENCIMENTO': vencimento_fatura,
                    'TICKET': match_df[col_ticket] if col_ticket else 'NÃO ENCONTRADO',
                    'NOTA FISCAL': nf_val,
                    'CEP': cep_val,
                    'UF': uf_val,
                    'VALOR FRETE': valor_frete,
                    'NOME DA TRANSPORTADORA': match_df[col_transp],
                    'STATUS': match_df[col_status],
                    'MOTIVO': match_df[col_motivo] if col_motivo else 'NÃO ENCONTRADO',
                    'DATA DA COLETA': match_df[col_data_coleta]
                })
            else:
                resultados.append({
                    'NUMERO DA FATURA': num_fatura,
                    'DATA DE VENCIMENTO': vencimento_fatura,
                    'TICKET': 'NÃO ENCONTRADO',
                    'NOTA FISCAL': nf_val,
                    'CEP': cep_val,
                    'UF': uf_val,
                    'VALOR FRETE': valor_frete,
                    'NOME DA TRANSPORTADORA': 'NÃO ENCONTRADO',
                    'STATUS': 'NÃO ENCONTRADO',
                    'MOTIVO': 'NÃO ENCONTRADO',
                    'DATA DA COLETA': 'NÃO ENCONTRADO'
                })

    # 3. Consolidar e salvar
    df_resultados = pd.DataFrame(resultados)
    
    if is_local_mode:
        print("\n" + "="*80)
        print("RESUMO DA CONSOLIDAÇÃO DOS DADOS (Tabela Oculta para não travar o console)")
        print("="*80)
        # with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.width', 1000):
        #     print(df_resultados)
        
        output_file = os.path.join(script_dir, 'Consolidado_Faturas_Coletas.xlsx')
        df_resultados.to_excel(output_file, index=False)
        print(f"\nOs dados acima também foram salvos no arquivo: {output_file}")
        
    return df_resultados

if __name__ == '__main__':
    consolidate_data()
