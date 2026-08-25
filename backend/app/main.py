import os
import logging
from datetime import datetime
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import openpyxl
import pandas as pd

# Настройка логирования
LOG_DIR = os.environ.get('LOG_DIR', '/logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, 'app.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = FastAPI(title="RN Reconciliation API", version="1.0.0")

@app.get("/")
async def root():
    logger.info("Root endpoint called")
    return {"message": "RN Reconciliation API", "status": "running"}

@app.get("/health")
async def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Загрузка Excel файла для обработки"""
    logger.info(f"Received file: {file.filename}")
    
    try:
        # Сохраняем загруженный файл
        file_path = os.path.join('/data', file.filename)
        os.makedirs('/data', exist_ok=True)
        
        content = await file.read()
        with open(file_path, 'wb') as f:
            f.write(content)
        
        logger.info(f"File saved: {file_path}")
        
        # Пробуем прочитать Excel
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheets = wb.sheetnames
            logger.info(f"File contains sheets: {sheets}")
            return JSONResponse({
                "status": "success",
                "filename": file.filename,
                "sheets": sheets,
                "message": "File uploaded and validated successfully"
            })
        except Exception as e:
            logger.error(f"Failed to parse Excel: {str(e)}")
            return JSONResponse({
                "status": "error",
                "filename": file.filename,
                "error": str(e),
                "message": "File uploaded but could not parse as Excel"
            }, status_code=400)
            
    except Exception as e:
        logger.error(f"Upload failed: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/reconcile")
async def reconcile():
    """Запуск сверки (порт VBA логики)"""
    logger.info("Reconciliation started")
    
    try:
        # TODO: Реализовать логику сверки на основе VBA кода
        # Сейчас возвращаем заглушку
        result = {
            "status": "in_progress",
            "message": "Reconciliation logic is being ported from VBA",
            "timestamp": datetime.now().isoformat()
        }
        
        logger.info("Reconciliation completed (placeholder)")
        return JSONResponse(result)
        
    except Exception as e:
        logger.error(f"Reconciliation failed: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/logs")
async def get_logs(limit: int = 100):
    """Получение последних логов"""
    log_file = os.path.join(LOG_DIR, 'app.log')
    
    if not os.path.exists(log_file):
        return {"logs": [], "message": "No logs found"}
    
    try:
        with open(log_file, 'r') as f:
            lines = f.readlines()
            last_lines = lines[-limit:] if len(lines) > limit else lines
            return {"logs": last_lines, "total_lines": len(lines)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.on_event("startup")
async def startup_event():
    logger.info("Application startup")
    logger.info(f"LOG_DIR: {LOG_DIR}")

@app.on_event("shutdown")
async def shutdown_event():
    logger.info("Application shutdown")
