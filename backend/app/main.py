import os
import logging
from datetime import datetime
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
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

# CORS для фронтенда
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    logger.info("Root endpoint called")
    return {"message": "RN Reconciliation API", "status": "running"}

@app.get("/health")
async def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

def process_excel_file(filepath: str):
    """Общая функция для обработки Excel-файлов"""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheets = wb.sheetnames
        
        # Читаем данные из первого листа
        df = pd.read_excel(filepath, sheet_name=sheets[0])
        rows = len(df)
        
        return {
            "sheets": sheets,
            "rows": rows,
            "columns": list(df.columns)[:10]  # Первые 10 колонок для информации
        }
    except Exception as e:
        logger.error(f"Ошибка обработки Excel: {str(e)}")
        raise

@app.post("/upload-bank")
async def upload_bank(file: UploadFile = File(...)):
    """Загрузка банковского файла"""
    logger.info(f"Банковский файл: {file.filename}")
    
    try:
        file_path = os.path.join('/data', f"bank_{file.filename}")
        os.makedirs('/data', exist_ok=True)
        
        content = await file.read()
        with open(file_path, 'wb') as f:
            f.write(content)
        
        result = process_excel_file(file_path)
        
        return JSONResponse({
            "status": "success",
            "filename": file.filename,
            "rows": result["rows"],
            "sheets": result["sheets"],
            "columns": result["columns"],
            "message": "Банковский файл загружен"
        })
    except Exception as e:
        logger.error(f"Ошибка загрузки банковского файла: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/upload-rn")
async def upload_rn(file: UploadFile = File(...)):
    """Загрузка файла РН-Карт"""
    logger.info(f"РН-Карт файл: {file.filename}")
    
    try:
        file_path = os.path.join('/data', f"rn_{file.filename}")
        os.makedirs('/data', exist_ok=True)
        
        content = await file.read()
        with open(file_path, 'wb') as f:
            f.write(content)
        
        result = process_excel_file(file_path)
        
        return JSONResponse({
            "status": "success",
            "filename": file.filename,
            "rows": result["rows"],
            "sheets": result["sheets"],
            "columns": result["columns"],
            "message": "РН-Карт файл загружен"
        })
    except Exception as e:
        logger.error(f"Ошибка загрузки РН-Карт файла: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/reconcile")
async def reconcile():
    """Запуск сверки"""
    logger.info("="*50)
    logger.info("RECONCILE ENDPOINT CALLED")
    logger.info("="*50)
    
    try:
        # Проверяем наличие файлов
        bank_files = [f for f in os.listdir('/data') if f.startswith('bank_')]
        rn_files = [f for f in os.listdir('/data') if f.startswith('rn_')]
        
        logger.info(f"Bank files: {bank_files}")
        logger.info(f"RN files: {rn_files}")
        
        if not bank_files or not rn_files:
            logger.error("Files not found")
            return JSONResponse({
                "status": "error",
                "message": "Не загружены оба файла"
            }, status_code=400)
        
        # Загружаем последние файлы
        bank_path = os.path.join('/data', bank_files[-1])
        rn_path = os.path.join('/data', rn_files[-1])
        
        logger.info(f"Loading bank: {bank_path}")
        logger.info(f"Loading rn: {rn_path}")
        
        # Читаем Excel
        bank_df = pd.read_excel(bank_path)
        rn_df = pd.read_excel(rn_path)
        
        logger.info(f"Bank columns: {bank_df.columns.tolist()}")
        logger.info(f"RN columns: {rn_df.columns.tolist()}")
        logger.info(f"Bank shape: {bank_df.shape}")
        logger.info(f"RN shape: {rn_df.shape}")
        
        # Показываем первые 5 строк для примера
        logger.info(f"Bank head:\n{bank_df.head()}")
        logger.info(f"RN head:\n{rn_df.head()}")
        
        # Инициализируем движок и запускаем
        engine = ReconciliationEngine()
        engine.set_data(bank_df, rn_df)
        result = engine.run()
        
        logger.info(f"Reconciliation result: {result['summary']}")
        
        return JSONResponse({
            "status": "success",
            "results": result
        })
        
    except Exception as e:
        logger.error(f"Reconciliation failed: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return JSONResponse({
            "status": "error",
            "message": str(e)
        }, status_code=500)
                "bank_count": 0,
                "rn_count": 0,
                "summary": {
                    "total_matches": 0,
                    "unmatched_bank": 0,
                    "unmatched_rn": 0,
                    "match_rate_bank": 0,
                    "match_rate_rn": 0
                },
                "matches": [],
                "unmatched_bank": [],
                "unmatched_rn": [],
                "link_analysis": {"by_contragent": {}},
                "base_comparison": {
                    "total_checked": 0,
                    "in_base": [],
                    "not_in_base": []
                }
            },
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
