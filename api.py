
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
import io
from datetime import datetime
from pydantic import BaseModel
import os
from dotenv import load_dotenv
import tempfile
import logging
import smtplib
import ssl
import asyncio
from services import (
    process_excel_data,
    merge_notas_descargas,
    download_excel_from_url,
    delete_file,
        generate_mesclado_and_errors,
        format_mesclado_records,
        excel_bytes_from_records,
        build_filename,
    upar_arquivo_sharedpoint,
)
from emailer import Emailer

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = FastAPI(title="Descarga API", version="1.0.0")


class URLDownloadRequest(BaseModel):
    """Modelo para requisição de download via URL"""
    url: str


# Carregar variáveis de ambiente de .env quando executado localmente
load_dotenv()


@app.post("/api/descargas-mescladas")
async def get_descargas_mescladas(data: dict):
    """
    POST: Recebe JSON com listas de 'notas' e 'descargas'
    
    Payload esperado:
    {
        "notas": [...],
        "itens": [...]
    }
    
    Retorna os dados mesclados usando cte_origem como chave.
    """
    try:
        if not data:
            logger.warning(f"Status 400 error: Payload JSON não fornecido ou inválido")
            raise HTTPException(status_code=400, detail="Payload JSON não fornecido ou inválido")

        if 'notas' not in data or 'itens' not in data:
            logger.warning(f'Status 400 error: O JSON deve conter as chaves "notas" e "itens"')
            raise HTTPException(
                status_code=400,
                detail='O JSON deve conter as chaves "notas" e "itens"'
            )

        mesclado = merge_notas_descargas(data)

        return JSONResponse(content=mesclado, status_code=200)

    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: Erro ao processar dados: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Erro ao processar dados: {str(e)}")


@app.post("/api/descargas-mescladas-xlsx")
async def get_descargas_mescladas_xlsx(data: dict):
    """
    POST: Recebe JSON com listas de 'notas' e 'descargas' e uma chave 'nome_arquivo'

    Payload esperado:
    {
        "notas": [...],
        "itens": [...],
        "nome_arquivo": "arquivo_xpto"
    }

    Retorna um arquivo .xlsx contendo os registros mesclados. O nome do arquivo
    será: <nome_arquivo>_DDMMAA_HHMM.xlsx (ex: arquivo_xpto_090626_2311.xlsx)
    """
    try:
        if not data:
            logger.warning(f"Status 400 error: Payload JSON não fornecido ou inválido")
            raise HTTPException(status_code=400, detail="Payload JSON não fornecido ou inválido")

        if 'notas' not in data or 'itens' not in data:
            logger.warning(f'Status 400 error: O JSON deve conter as chaves "notas" e "itens"')
            raise HTTPException(
                status_code=400,
                detail='O JSON deve conter as chaves "notas" e "itens"'
            )

        nome_arquivo = data.get('nome_arquivo')
        if not nome_arquivo or str(nome_arquivo).strip() == '':
            logger.warning('Status 400 error: nome_arquivo não fornecido ou inválido')
            raise HTTPException(status_code=400, detail='nome_arquivo não fornecido ou inválido')

        # Gerar os dados mesclados usando a função existente
        mesclado = merge_notas_descargas(data)

        # Converter para DataFrame e salvar em Excel na memória
        import pandas as pd

        df = pd.DataFrame(mesclado)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Mesclado')
        output.seek(0)

        # Calcular tamanho e Content-Range para envio via REST
        size = output.getbuffer().nbytes

        timestamp = datetime.now().strftime('%d%m%y_%H%M')
        filename = f"{nome_arquivo}_{timestamp}.xlsx"

        headers = {
            'Content-Disposition': f'attachment; filename="{filename}"',
            'Content-Length': str(size),
            'Content-Range': f'bytes 0-{max(size-1,0)}/{size}',
            'Accept-Ranges': 'bytes'
        }

        return StreamingResponse(
            output,
            media_type='application/octet-stream',
            headers=headers
        )

    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: Erro ao gerar XLSX: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Erro ao gerar XLSX: {str(e)}")


@app.post("/api/upload-headers")
async def upload_and_return_headers(file: UploadFile = File(...)):
    """
    Recebe um arquivo via multipart/form-data e retorna em JSON
    os headers calculados necessários para enviar o binário a outro endpoint:
    Content-Disposition, Content-Length, Content-Range, Accept-Ranges
    """
    try:
        if not file:
            logger.warning("Status 400 error: Nenhum arquivo enviado")
            raise HTTPException(status_code=400, detail="Nenhum arquivo enviado")

        # Ler conteúdo em memória
        content = await file.read()
        size = len(content)

        # Usar o nome original do arquivo se fornecido, caso contrário gerar um nome
        filename = file.filename if file.filename and file.filename.strip() != '' else f"upload_{datetime.now().strftime('%d%m%y_%H%M%S')}"

        headers = {
            'Content-Disposition': f'attachment; filename="{filename}"',
            'Content-Length': str(size),
            'Content-Range': f'bytes 0-{max(size-1,0)}/{size}',
            'Accept-Ranges': 'bytes'
        }

        return JSONResponse(content=headers, status_code=200)

    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: Erro ao processar upload: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Erro ao processar upload: {str(e)}")


@app.post("/api/descargas-mescladas-email")
async def get_descargas_mescladas_and_email(data: dict):
    """
    Recebe JSON com 'notas', 'itens', 'nome_arquivo' e dados SMTP:

    {
      "notas": [...],
      "itens": [...],
      "nome_arquivo": "arquivo_xpto",
      "smtp": {
         "host": "smtp.example.com",
         "port": 587,
         "user": "user",
         "password": "pass",
         "use_ssl": false,
         "starttls": true
      },
      "from_email": "from@example.com",
      "to_emails": ["to1@example.com"],
      "subject": "Assunto opcional"
    }

    Gera o XLSX, envia por SMTP com um template HTML e lista os itens com erro.
    """
    try:
        if not data:
            logger.warning("Status 400 error: Payload JSON não fornecido ou inválido")
            raise HTTPException(status_code=400, detail="Payload JSON não fornecido ou inválido")

        if 'notas' not in data or 'itens' not in data:
            logger.warning('Status 400 error: O JSON deve conter as chaves "notas" e "itens"')
            raise HTTPException(status_code=400, detail='O JSON deve conter as chaves "notas" e "itens"')

        nome_arquivo = data.get('nome_arquivo')
        if not nome_arquivo or str(nome_arquivo).strip() == '':
            logger.warning('Status 400 error: nome_arquivo não fornecido ou inválido')
            raise HTTPException(status_code=400, detail='nome_arquivo não fornecido ou inválido')

        # Construir Emailer a partir do payload (ele fará fallback para
        # variáveis de ambiente quando necessário). Não carregar env aqui.
        smtp_conf = data.get('smtp') or None
        from_email = data.get('from_email')
        to_emails = data.get('to_emails')
        subject = data.get('subject')

        # Mesclar dados, identificar erros e gerar XLSX em bytes via services
        mesclado, error_items = generate_mesclado_and_errors(data)

        # Formatar registros para as colunas requeridas e gerar o XLSX
        records_for_excel = format_mesclado_records(mesclado)
        excel_bytes = excel_bytes_from_records(records_for_excel)
        filename = build_filename(nome_arquivo)

        # Preparar Emailer (ele monta a mensagem e realiza o envio)
        emailer = Emailer(from_email=from_email, to_emails=to_emails, subject=subject, smtp_conf=smtp_conf)
        # total_records deve ser o número de CTe Origem únicos
        unique_ctes = {str(r.get('cte_origem')) for r in mesclado if r.get('cte_origem') not in (None, '')}
        total_records = len(unique_ctes)
        await emailer.send(filename=filename, attachment_bytes=excel_bytes, error_items=error_items, total_records=total_records)

        upload_result = None
        upload_url = data.get('upload_url') or data.get('uploadUrl')
        if upload_url:
            try:
                upload_result = upar_arquivo_sharedpoint(excel_bytes, upload_url, filename)
            except Exception as upload_exc:
                logger.error(f"Erro ao upar arquivo para SharePoint: {str(upload_exc)}")
                upload_result = f"Erro ao upar arquivo para SharePoint: {str(upload_exc)}"  

        return JSONResponse(content={
            'status': 'sent',
            'filename': filename,
            'recipients': to_emails,
            'error_items': len(error_items),
            'upload': upload_result
        }, status_code=200)

    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: Erro ao enviar e-mail: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Erro ao enviar e-mail: {str(e)}")


@app.get("/api/descargas")
async def get_descargas_get():
    """
    GET: Processa arquivo padrão se existir no diretório
    
    Retorna:
    {
        "notas": [{"num_nfe": ..., "cte_origem": ...}],
        "descargas": [{"num_cte_da_des": ..., "vlr_cte": ..., ...}]
    }
    """
    try:
        default_file = 'Base - Descarga Automática 2204.xlsx'
        if os.path.exists(default_file):
            data = process_excel_data(default_file)
            logger.info(f"Arquivo padrão processado com sucesso: {default_file}")
            return JSONResponse(content=data, status_code=200)
        else:
            logger.warning(f"Status 404 error: Arquivo padrão não encontrado. {default_file}")
            raise HTTPException(
                status_code=404,
                detail={
                    'error': 'Arquivo padrão não encontrado. Use POST /api/descargas para enviar um arquivo.',
                    'expectedFile': default_file
                }
            )
    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/descargas")
async def get_descargas_post(file: UploadFile = File(...)):
    """
    POST: Recebe arquivo via multipart/form-data
    
    Retorna:
    {
        "notas": [{"num_nfe": ..., "cte_origem": ...}],
        "descargas": [{"num_cte_da_des": ..., "vlr_cte": ..., ...}]
    }
    """
    try:
        if not file:
            logger.warning("Status 400 error: Nenhum arquivo enviado")
            raise HTTPException(status_code=400, detail="Nenhum arquivo enviado")
        
        if file.filename == '':
            logger.warning("Status 400 error: Arquivo sem nome")
            raise HTTPException(status_code=400, detail="Arquivo sem nome")

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name
        
        try:
            data = process_excel_data(tmp_path)
            logger.info(f"Arquivo {file.filename} processado com sucesso")
            return JSONResponse(content=data, status_code=200)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    
    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: Erro ao processar arquivo: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/descargas-url")
async def get_descargas_from_url(request: URLDownloadRequest):
    """
    POST: Recebe URL de download de arquivo Excel com token embutido (padrão Microsoft)
    
    Payload esperado:
    {
        "url": "https://example.sharepoint.com/...?token=..."
    }
    
    Retorna:
    {
        "notas": [{"num_nfe": ..., "cte_origem": ...}],
        "descargas": [{"num_cte_da_des": ..., "vlr_cte": ..., ...}]
    }
    """
    try:
        if not request.url or request.url.strip() == '':
            logger.warning("Status 400 error: URL não fornecida ou inválida")
            raise HTTPException(status_code=400, detail="URL não fornecida ou inválida")
        
        # Baixar arquivo da URL
        tmp_path = download_excel_from_url(request.url)
        logger.info(f"Arquivo baixado da URL com sucesso")
        
        try:
            # Processar dados do arquivo
            data = process_excel_data(tmp_path)
            logger.info(f"Arquivo da URL processado com sucesso")
            return JSONResponse(content=data, status_code=200)
        finally:
            # Garantir que o arquivo temporário seja deletado
            delete_file(tmp_path)
    
    except HTTPException as http_exc:
        logger.error(f"Status {http_exc.status_code} error: {http_exc.detail}")
        raise
    except Exception as e:
        logger.error(f"Status 500 error: Erro ao processar URL: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/health")
async def health():
    """Health check endpoint"""
    return JSONResponse(content={"status": "ok"}, status_code=200)


if __name__ == '__main__':
    import uvicorn
    print("Iniciando API na porta 8000...")
    print("\nEndpoints disponíveis:")
    print("  GET  /api/descargas - Processa arquivo padrão (se existir)")
    print("  POST /api/descargas - Recebe arquivo Excel via multipart/form-data")
    print("  POST /api/descargas-mescladas - Mescla notas com descargas")
    print("  POST /api/descargas-mescladas-xlsx - Mescla e retorna XLSX para download")
    print("  GET  /api/health - Health check")
    print("  GET  /docs - Documentação interativa (Swagger UI)")
    uvicorn.run(app, host="0.0.0.0", port=8082)


