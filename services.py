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
    result = {
        'notas': [],
        'itens': []
    }
    
    # Agrupar por CTe Descarga, CTe Origem, CNPJ e Cod Emissor
    group_cols = [df_descarga.iloc[:, 9], df_descarga.iloc[:, 11], df_descarga.iloc[:, 4], df_descarga.iloc[:, 13]]
    
    notas_unicas = df_descarga[[df_descarga.columns[0], df_descarga.columns[11]]].drop_duplicates()
    for _, row in notas_unicas.iterrows():
        try:
            num_nfe = int(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        except (ValueError, TypeError):
            num_nfe = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        
        try:
            cte_origem = int(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        except (ValueError, TypeError):
            cte_origem = str(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        
        result['notas'].append({
            'num_nfe': str(num_nfe),
            'cte_origem': str(cte_origem)
        })

    descargas_grouped = df_descarga.groupby(by=[df_descarga.iloc[:, 9], df_descarga.iloc[:, 11], df_descarga.iloc[:, 4], df_descarga.iloc[:, 13]], dropna=False).agg({
        df_descarga.columns[10]: 'sum'
    }).reset_index()
    
    for _, row in descargas_grouped.iterrows():
        try:
            num_cte_da_des = int(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        except (ValueError, TypeError):
            num_cte_da_des = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        
        try:
            vlr_cte = round(float(row.iloc[4]), 2) if pd.notna(row.iloc[4]) else 0
        except (ValueError, TypeError):
            vlr_cte = 0
        
        try:
            cte_origem = int(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        except (ValueError, TypeError):
            cte_origem = str(row.iloc[1]) if pd.notna(row.iloc[1]) else None
        
        try:
            cod_emissor = int(row.iloc[3]) if pd.notna(row.iloc[3]) else None
        except (ValueError, TypeError):
            cod_emissor = str(row.iloc[3]) if pd.notna(row.iloc[3]) else None

        result['itens'].append({
            'cte_des': str(num_cte_da_des),
            'valor': f"{vlr_cte:.2f}",
            'cte_origem': str(cte_origem),
            'cnpj_origem': str(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
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
    # Keep track of which nota CTEs exist so we can detect unmatched descargas
    nota_ctes = set()
    for nota in notas:
        cte_origem = nota.get('cte_origem')
        nota_ctes.add(str(cte_origem))
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

    # Append descargas that have no matching nota (so errors on these aren't lost)
    for descarga in descargas:
        if str(descarga.get('cte_origem')) not in nota_ctes:
            resultado.append(descarga)

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
      - Data Lançamento

    Recebe a lista retornada por `merge_notas_descargas` e normaliza os campos.
    """
    formatted = []
    for rec in mesclado:
        detalhe = rec.get('detalhe_erro') if isinstance(rec, dict) else None
        erro = None
        # aceitar vários formatos de detalhe_erro: dict, list[e.g. [{'error': 'msg'}]], ou string
        if isinstance(detalhe, dict):
            erro = detalhe.get('error') or detalhe.get('Erro')
        elif isinstance(detalhe, list) and detalhe:
            # juntar mensagens caso exista mais de uma entrada
            msgs = []
            for item in detalhe:
                if isinstance(item, dict):
                    m = item.get('error') or item.get('Erro')
                    if m:
                        msgs.append(str(m))
                elif isinstance(item, str):
                    msgs.append(item)
            if msgs:
                erro = '; '.join(msgs)
        elif isinstance(detalhe, str):
            erro = detalhe

        # fallbacks for common field names
        nota = rec.get('num_nfe') or rec.get('num_nfe')
        cod_emissor = rec.get('cod_emissor') or rec.get('codEmissor')
        cnpj = rec.get('cnpj_origem') or rec.get('cnpj') or rec.get('cnpjOrigem')
        cte_origem = rec.get('cte_origem') or rec.get('cteOrigem')
        cte_des = rec.get('cte_des') or rec.get('cteDes')
        valor = rec.get('valor')
        if valor is None or str(valor) == '':
            valor = None
        else:
            try:
                valor = f"{float(valor):.2f}"
            except Exception:
                valor = str(valor)

        # formatar campo Data Lançamento se vier como string ISO (ex: 2026-06-16T10:16:00.196084Z)
        data_lancamento_formatted = None
        dl = rec.get('data_lancamento')
        if dl is not None:
            try:
                dt = None
                if isinstance(dl, datetime):
                    dt = dl
                elif isinstance(dl, str):
                    s = dl.strip()
                    if s.endswith('Z'):
                        # converter Z para offset compatível com fromisoformat
                        s = s[:-1] + '+00:00'
                    dt = datetime.fromisoformat(s)

                if dt:
                    data_lancamento_formatted = dt.strftime('%m-%d-%Y %H:%M')
                else:
                    data_lancamento_formatted = dl
            except Exception:
                # se não conseguir parsear, manter o valor original
                data_lancamento_formatted = dl

        formatted.append({
            'Nota': nota,
            'Cod. Emissor': cod_emissor,
            'CNPJ': cnpj,
            'CTe Origem': cte_origem,
            'CTe Descarga': cte_des,
            'Valor': valor,
            'Status': rec.get('status'),
            'Erro': erro,
            'Num. Calculo': rec.get('num_calculo'),
            'Data Lançamento': data_lancamento_formatted
        })

    return formatted


def build_mesclado_for_xlsx_by_cte(data):
    """Gera registros por `nota` (não por CTe) para uso no XLSX.

    Para cada nota em `data['notas']` o método agrega os valores
    do CTe de origem correspondente e replica esses valores na linha
    da nota. Assim o número de registros retornados será igual ao
    número de notas.
    """
    notas = data.get('notas', [])
    descargas = data.get('itens', [])

    # Agregar informações por CTe (mesma lógica anterior)
    desc_by_cte = {}
    for d in descargas:
        cte = str(d.get('cte_origem')) if d.get('cte_origem') is not None else ''
        desc_by_cte.setdefault(cte, []).append(d)

    agg_by_cte = {}
    for cte, descs in desc_by_cte.items():
        valor_total = 0.0
        cte_descs = []
        erros = []
        num_calculos = []
        data_lancamentos = []
        cod_emissores = set()
        status = None
        cnjps = set()

        for d in descs:
            v = d.get('valor')
            try:
                if v is not None and str(v) != '':
                    valor_total += float(v)
            except Exception:
                pass

            if d.get('cte_des'):
                cte_descs.append(str(d.get('cte_des')))

            if status is None and d.get('status'):
                status = d.get('status')

            detalhe = d.get('detalhe_erro')
            if detalhe:
                if isinstance(detalhe, list):
                    for it in detalhe:
                        if isinstance(it, dict):
                            m = it.get('error') or it.get('Erro')
                            if m:
                                erros.append(str(m))
                        elif isinstance(it, str):
                            erros.append(it)
                elif isinstance(detalhe, dict):
                    m = detalhe.get('error') or detalhe.get('Erro')
                    if m:
                        erros.append(str(m))
                elif isinstance(detalhe, str):
                    erros.append(detalhe)

            if d.get('num_calculo'):
                num_calculos.append(str(d.get('num_calculo')))
            if d.get('data_lancamento'):
                data_lancamentos.append(str(d.get('data_lancamento')))
            if d.get('cod_emissor'):
                cod_emissores.add(str(d.get('cod_emissor')))
            if d.get('cnpj_origem'):
                cnjps.add(str(d.get('cnpj_origem')))

        agg_by_cte[cte] = {
            'CTe Descarga': ','.join(cte_descs) if cte_descs else '',
            'Valor Total': f"{valor_total:.2f}" if valor_total else '',
            'Status': status,
            'Erro': '; '.join(erros) if erros else '',
            'Cod. Emissor': ','.join(sorted(cod_emissores)) if cod_emissores else '',
            'Num. Calculo': ','.join(num_calculos) if num_calculos else '',
            'Data Lançamento': ','.join(data_lancamentos) if data_lancamentos else '',
            'CNPJ': ','.join(sorted(cnjps)) if cnjps else ''
        }

    # Agora, para cada nota, gerar um registro replicando os agregados do CTe
    records = []
    for n in notas:
        nota_num = n.get('num_nfe')
        cte = str(n.get('cte_origem')) if n.get('cte_origem') is not None else ''

        agg = agg_by_cte.get(cte, None)
        record = {
            'Nota': nota_num,
            'Cod. Emissor': agg.get('Cod. Emissor') if agg else '',
            'CNPJ': agg.get('CNPJ') if agg else '',
            'CTe Origem': cte,
            'CTe Descarga': agg.get('CTe Descarga') if agg else '',
            'Valor Total': agg.get('Valor Total') if agg else '',
            'Status': agg.get('Status') if agg else None,
            'Erro': agg.get('Erro') if agg else '',
            'Num. Calculo': agg.get('Num. Calculo') if agg else '',
            'Data Lançamento': agg.get('Data Lançamento') if agg else ''
        }
        records.append(record)

    return records


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


def upar_arquivo_sharedpoint(file_bytes: bytes, url_upload: str, filename: str = None, timeout: int = 30) -> dict:
    """Faz upload de um arquivo (bytes) para um endpoint (ex: SharePoint) usando headers esperados.

    Args:
        file_bytes: Conteúdo do arquivo em bytes (XLSX)
        url_upload: URL de upload (PUT)
        filename: Nome do arquivo para usar no header Content-Disposition
        timeout: timeout em segundos para a requisição

    Returns:
        dict: resultado com chaves: url, status_code, ok, text (parcial) ou error
    """
    print(f"Iniciando upload do arquivo para {url_upload} com timeout de {timeout} segundos")
    try:
        size = len(file_bytes) if file_bytes is not None else 0
        name = filename or 'upload.xlsx'
        headers = {
            'Content-Disposition': f'attachment; filename="{name}"',
            'Content-Length': str(size),
            'Content-Range': f'bytes 0-{max(size-1,0)}/{size}',
            'Accept-Ranges': 'bytes'
        }

        resp = requests.put(url_upload, data=file_bytes, headers=headers, timeout=timeout)

        return {
            'url': url_upload,
            'status_code': getattr(resp, 'status_code', None),
            'ok': getattr(resp, 'ok', False),
            'text': getattr(resp, 'text', '')[:1000]
        }

    except Exception as e:
        return {'url': url_upload, 'ok': False, 'error': str(e)}


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

        # Caso 2: existe detalhe_erro com mensagem de erro (suporta dict ou lista)
        if isinstance(detalhe_erro, dict) and detalhe_erro:
            error_text = detalhe_erro.get('error') or detalhe_erro.get('Erro')
            if error_text:
                has_error = True
                erro_msg = error_text
        elif isinstance(detalhe_erro, list) and detalhe_erro:
            msgs = []
            for item in detalhe_erro:
                if isinstance(item, dict):
                    m = item.get('error') or item.get('Erro')
                    if m:
                        msgs.append(str(m))
                elif isinstance(item, str):
                    msgs.append(item)
            if msgs:
                has_error = True
                erro_msg = '; '.join(msgs)

        # Se encontrou erro, adicionar à lista
        if has_error:
            # Extrair mensagem de erro do detalhe_erro se não tiver sido extraída
            if not erro_msg:
                if isinstance(detalhe_erro, dict):
                    erro_msg = detalhe_erro.get('error') or detalhe_erro.get('Erro') or ''
                if isinstance(detalhe_erro, list):
                    erro_msg = detalhe_erro[0].get('error') or detalhe_erro[0].get('Erro') or ''
            error_record = {
                'num_nfe': rec.get('num_nfe'),
                'cte_origem': rec.get('cte_origem'),
                'cod_emissor': rec.get('cod_emissor'),
                'erro': erro_msg
            }
            error_items.append(error_record)
        
    return mesclado, error_items
