
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import os
import tempfile
from services import process_excel_data, merge_notas_descargas, download_excel_from_url, delete_file

app = FastAPI(title="Descarga API", version="1.0.0")


class URLDownloadRequest(BaseModel):
    """Modelo para requisição de download via URL"""
    url: str


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
            raise HTTPException(status_code=400, detail="Payload JSON não fornecido ou inválido")

        if 'notas' not in data or 'itens' not in data:
            raise HTTPException(
                status_code=400,
                detail='O JSON deve conter as chaves "notas" e "itens"'
            )

        mesclado = merge_notas_descargas(data)

        return JSONResponse(content=mesclado, status_code=200)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao processar dados: {str(e)}")


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
            return JSONResponse(content=data, status_code=200)
        else:
            raise HTTPException(
                status_code=404,
                detail={
                    'error': 'Arquivo padrão não encontrado. Use POST /api/descargas para enviar um arquivo.',
                    'expectedFile': default_file
                }
            )
    except HTTPException:
        raise
    except Exception as e:
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
            raise HTTPException(status_code=400, detail="Nenhum arquivo enviado")
        
        if file.filename == '':
            raise HTTPException(status_code=400, detail="Arquivo sem nome")

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name
        
        try:
            data = process_excel_data(tmp_path)
            return JSONResponse(content=data, status_code=200)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    
    except HTTPException:
        raise
    except Exception as e:
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
            raise HTTPException(status_code=400, detail="URL não fornecida ou inválida")
        
        # Baixar arquivo da URL
        tmp_path = download_excel_from_url(request.url)
        
        try:
            # Processar dados do arquivo
            data = process_excel_data(tmp_path)
            return JSONResponse(content=data, status_code=200)
        finally:
            # Garantir que o arquivo temporário seja deletado
            delete_file(tmp_path)
    
    except HTTPException:
        raise
    except Exception as e:
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
    print("  GET  /api/health - Health check")
    print("  GET  /docs - Documentação interativa (Swagger UI)")
    uvicorn.run(app, host="0.0.0.0", port=8082)


