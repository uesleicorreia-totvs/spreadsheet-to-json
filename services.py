import pandas as pd
import requests
import tempfile
import os


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
    
    df_descarga = pd.read_excel(file_path, sheet_name='Base Descarga Automática')
    
    df_descarga.columns = df_descarga.columns.str.strip()
    df_descarga.columns = df_descarga.columns.str.replace('\xa0', '', regex=False)

    col_nota = [col for col in df_descarga.columns if 'Nota Fiscal' in col][0]
    col_cte_origem = [col for col in df_descarga.columns if 'CTe Origem' in col][0]
    col_cte_descarga = [col for col in df_descarga.columns if 'Nº CTe da Descarga' in col][0]
    col_vlr_cte = [col for col in df_descarga.columns if 'Valor CTe de Descarga' in col][0]
    col_cnpj = [col for col in df_descarga.columns if 'CNPJ Emissor do CTe de Origem' in col][0]
    col_cod_emissor = [col for col in df_descarga.columns if 'Cód. Emissor' in col][0]
    
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

    descargas_grouped = df_descarga.groupby(
        [col_cte_descarga, col_cte_origem, col_cnpj, col_cod_emissor],
        dropna=False
    ).agg({
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
            'cnpj_origem': str(row[col_cnpj]) if pd.notna(row[col_cnpj]) else None,
            'cod_emissor': str(cod_emissor)
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
