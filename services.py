import pandas as pd
import requests
import tempfile
import os
import io
from datetime import datetime


def download_excel_from_url(url: str) -> str:
    """
    Baixa arquivo Excel de uma URL com token embutido (padrão Microsoft).
    Salva o arquivo em um arquivo temporário e retorna o caminho.
    
    Args:
        url: URL de download direto com token embutido
        
    Returns:
        str: Caminho do arquivo temporário salvo
        
    Raises:
        Exception: Se o download falhar ou se o arquivo não for .xlsx válido
    """
    try:
        # Fazer download do arquivo
        response = requests.get(url, timeout=30, stream=True)
        response.raise_for_status()
        
        # Criar arquivo temporário com extensão .xlsx
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    tmp.write(chunk)
            tmp_path = tmp.name
        
        return tmp_path
        
    except requests.exceptions.RequestException as e:
        raise Exception(f"Erro ao baixar arquivo da URL: {str(e)}")
    except Exception as e:
        raise Exception(f"Erro ao salvar arquivo temporário: {str(e)}")


def delete_file(file_path: str) -> bool:
    """
    Deleta um arquivo do sistema de arquivos de forma segura.
    
    Args:
        file_path: Caminho completo do arquivo a deletar
        
    Returns:
        bool: True se deletado com sucesso, False caso não exista
    """
    try:
        if file_path and os.path.exists(file_path):
            os.remove(file_path)
            return True
        return False
    except Exception as e:
        print(f"Aviso: Erro ao deletar arquivo {file_path}: {str(e)}")
        return False


def process_excel_data(file_path):
    """Processa o arquivo Excel e retorna os dados estruturados"""
    
    # Tentar ler a folha esperada; se não existir, tentar alternativas
    try:
        df_descarga = pd.read_excel(file_path, sheet_name='Base Descarga Automática')
    except Exception:
        try:
            df_descarga = pd.read_excel(file_path, sheet_name='Planilha1')
        except Exception:
            # fallback: usar a primeira planilha presente no arquivo
            xls = pd.ExcelFile(file_path)
            first_sheet = xls.sheet_names[0]
            df_descarga = pd.read_excel(file_path, sheet_name=first_sheet)
    
    df_descarga.columns = df_descarga.columns.str.strip()
    df_descarga.columns = df_descarga.columns.str.replace('\xa0', '', regex=False)
    # Helper para localizar colunas por substring e fornecer erro legível
    def _find_col(containing: str, required: bool = True):
        matches = [col for col in df_descarga.columns if containing in col]
        if not matches:
            if required:
                raise ValueError(f"Coluna contendo '{containing}' não encontrada no arquivo. Colunas disponíveis: {list(df_descarga.columns)}")
            return None
        return matches[0]

    # Detectar se o template usa uma coluna 'Endereço' como referencia
    detect_address_header = any('Endereço' in str(c) for c in df_descarga.columns)
    if detect_address_header:
        def _col_by_letter(letter: str):
            idx = ord(letter.upper()) - 65
            if idx < 0 or idx >= len(df_descarga.columns):
                raise ValueError(f"Coluna por letra '{letter}' não existe no arquivo (total cols={len(df_descarga.columns)}).")
            return df_descarga.columns[idx]

        # Mapeamento fornecido pelo usuário
        col_nota = _col_by_letter('A')
        col_cte_origem = _col_by_letter('L')
        col_cte_descarga = _col_by_letter('J')
        col_vlr_cte = _col_by_letter('K')
        col_cnpj = _col_by_letter('E')
        col_cod_emissor = _col_by_letter('N')
        # colunas opcionais
        col_num_calculo = _find_col('Nº Cálculo', required=False)
        col_error = _find_col('Erro', required=False) or _find_col('detalhe_erro', required=False)
    else:
        col_nota = _find_col('Nota Fiscal')
        col_cte_origem = _find_col('CTe Origem')
        col_cte_descarga = _find_col('Nº CTe da Descarga')
        col_vlr_cte = _find_col('Valor CTe de Descarga')
        col_cnpj = _find_col('CNPJ Emissor do CTe de Origem')
        col_cod_emissor = _find_col('EMISSOR')
        col_num_calculo = _find_col('Nº Cálculo', required=False)
        col_error = _find_col('Erro', required=False) or _find_col('detalhe_erro', required=False)

    result = {
        'notas': [],
        'itens': []
    }
    

    notas_unicas = df_descarga[[col_nota, col_cte_origem]].drop_duplicates()
    for _, row in notas_unicas.iterrows():
        try:
            num_nfe = int(row[col_nota]) if pd.notna(row[col_nota]) else None
        except (ValueError, TypeError):
            num_nfe = str(row[col_nota]) if pd.notna(row[col_nota]) else None
        
        try:
            cte_origem = int(row[col_cte_origem]) if pd.notna(row[col_cte_origem]) else None
        except (ValueError, TypeError):
            cte_origem = str(row[col_cte_origem]) if pd.notna(row[col_cte_origem]) else None
        
        result['notas'].append({
            'num_nfe': str(num_nfe),
            'cte_origem': str(cte_origem)
        })

    # Agrupar usando colunas encontradas; col_error e col_num_calculo são opcionais
    group_cols = [col_cte_descarga, col_cte_origem, col_cod_emissor]
    if col_num_calculo:
        group_cols.append(col_num_calculo)
    if col_error:
        group_cols.append(col_error)

    descargas_grouped = df_descarga.groupby(group_cols, dropna=False).agg({
        col_vlr_cte: 'sum'
    }).reset_index()
    
    for _, row in descargas_grouped.iterrows():
        try:
            num_cte_da_des = int(row[col_cte_descarga]) if pd.notna(row[col_cte_descarga]) else None
        except (ValueError, TypeError):
            num_cte_da_des = str(row[col_cte_descarga]) if pd.notna(row[col_cte_descarga]) else None
        
        try:
            vlr_cte = round(float(row[col_vlr_cte]), 2) if pd.notna(row[col_vlr_cte]) else 0
        except (ValueError, TypeError):
            vlr_cte = 0
        
        try:
            cte_origem = int(row[col_cte_origem]) if pd.notna(row[col_cte_origem]) else None
        except (ValueError, TypeError):
            cte_origem = str(row[col_cte_origem]) if pd.notna(row[col_cte_origem]) else None
        
        try:
            cod_emissor = int(row[col_cod_emissor]) if pd.notna(row[col_cod_emissor]) else None
        except (ValueError, TypeError):
            cod_emissor = str(row[col_cod_emissor]) if pd.notna(row[col_cod_emissor]) else None

        result['itens'].append({
            'cte_des': str(num_cte_da_des),
            'valor': str(vlr_cte),
            'cte_origem': str(cte_origem),
            #'cnpj_origem': str(row[col_cnpj]) if pd.notna(row[col_cnpj]) else None,
            'cod_emissor': str(cod_emissor),
        })
    
    result['total_itens'] = len(result['itens'])     
    
    return result


def merge_notas_descargas(data):
    """
    Mescla notas com descargas usando cte_origem como chave de união.
    Para cada nota, inclui todas as descargas relacionadas.
    """
    notas = data['notas']
    descargas = data['itens']

    descargas_por_cte = {}
    for descarga in descargas:
        cte_origem = descarga['cte_origem']
        if cte_origem not in descargas_por_cte:
            descargas_por_cte[cte_origem] = []
        descargas_por_cte[cte_origem].append(descarga)

    resultado = []
    for nota in notas:
        cte_origem = nota['cte_origem']
        descargas_relacionadas = descargas_por_cte.get(cte_origem, [])
        
        if descargas_relacionadas:
            for descarga in descargas_relacionadas:
                registro = {
                    **nota, 
                    **descarga 
                }
                resultado.append(registro)
        else:
            resultado.append(nota)
    
    return resultado


def format_mesclado_records(mesclado):
    """Formata os registros mesclados para as colunas desejadas:

    Colunas retornadas (ordem):
      - Nota
      - Cod. Emissor
      - CNPJ
      - CTe Origem
      - CTe Descarga
      - Valor
      - Status
      - Erro (pega detalhe_erro.error quando disponível)
      - Num. Calculo

    Recebe a lista retornada por `merge_notas_descargas` e normaliza os campos.
    """
    formatted = []
    for rec in mesclado:
        detalhe = rec.get('detalhe_erro') if isinstance(rec, dict) else None
        erro = None
        if isinstance(detalhe, dict):
            # prefer key 'error' inside detalhe_erro
            erro = detalhe.get('error') or detalhe.get('Erro')
        elif isinstance(detalhe, str):
            erro = detalhe

        # fallbacks for common field names
        nota = rec.get('num_nfe') or rec.get('num_nfe')
        cod_emissor = rec.get('cod_emissor') or rec.get('codEmissor')
        cnpj = rec.get('cnpj_origem') or rec.get('cnpj') or rec.get('cnpjOrigem')
        cte_origem = rec.get('cte_origem') or rec.get('cteOrigem')
        cte_des = rec.get('cte_des') or rec.get('cteDes')
        valor = rec.get('valor')
        try:
            if valor is not None and str(valor) != '':
                valor = float(valor)
        except Exception:
            # keep as-is (string) if it cannot be converted
            pass

        formatted.append({
            'Nota': nota,
            'Cod. Emissor': cod_emissor,
            'CNPJ': cnpj,
            'CTe Origem': cte_origem,
            'CTe Descarga': cte_des,
            'Valor': valor,
            'Status': rec.get('status'),
            'Erro': erro,
            'Num. Calculo': rec.get('num_calculo')
        })

    return formatted


def build_filename(nome_arquivo: str) -> str:
    """Gera nome de arquivo com timestamp: <nome>_DDMMAA_HHMM.xlsx"""
    timestamp = datetime.now().strftime('%d%m%y_%H%M')
    return f"{nome_arquivo}_{timestamp}.xlsx"


def excel_bytes_from_records(records) -> bytes:
    """Converte uma lista de dicionários em um arquivo XLSX em bytes."""
    df = pd.DataFrame(records)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Mesclado')
    output.seek(0)
    return output.getvalue()


def generate_mesclado_and_errors(data):
    """Retorna os registros mesclados e a lista de itens com erro.

    Identifica erros quando:
      - status == 'erro' (independente de detalhe_erro)
      - status != 'sucesso' mas existe detalhe_erro com mensagem de erro
    
    Retorna error_items com campos: num_nfe, cte_origem, cod_emissor, erro
    """
    mesclado = merge_notas_descargas(data)
    error_items = []
    for rec in mesclado:
        status = rec.get('status')
        detalhe_erro = rec.get('detalhe_erro')
        
        # Verificar se há erro
        has_error = False
        erro_msg = ''
        
        # Caso 1: status = 'erro'
        if status and isinstance(status, str) and status.lower() == 'erro':
            has_error = True
        
        # Caso 2: existe detalhe_erro com mensagem de erro
        if isinstance(detalhe_erro, dict) and detalhe_erro:
            error_text = detalhe_erro.get('error') or detalhe_erro.get('Erro')
            if error_text:
                has_error = True
                erro_msg = error_text
        
        # Se encontrou erro, adicionar à lista
        if has_error:
            # Extrair mensagem de erro do detalhe_erro se não tiver sido extraída
            if not erro_msg and isinstance(detalhe_erro, dict):
                erro_msg = detalhe_erro.get('error') or detalhe_erro.get('Erro') or ''
            
            error_record = {
                'num_nfe': rec.get('num_nfe'),
                'cte_origem': rec.get('cte_origem'),
                'cod_emissor': rec.get('cod_emissor'),
                'erro': erro_msg
            }
            error_items.append(error_record)
    
    return mesclado, error_items
