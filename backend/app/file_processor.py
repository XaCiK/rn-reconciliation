"""Module for Excel file processing for reconciliation."""

import pandas as pd
import os
import logging
from typing import Dict, List, Optional
from openpyxl import load_workbook

logger = logging.getLogger(__name__)

class FileProcessor:
    """Excel file processor for reconciliation."""
    
    def __init__(self, data_dir: str = '/data'):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)
    
    def read_bank_file(self, filepath: str) -> pd.DataFrame:
        """Read bank file and normalize it."""
        logger.info(f"Reading bank file: {filepath}")
        
        try:
            xl = pd.ExcelFile(filepath)
            sheet_names = xl.sheet_names
            logger.info(f"Sheets found: {sheet_names}")
            
            target_sheet = None
            for sheet in sheet_names:
                if 'Name' in sheet or 'Лист' in sheet or sheet == sheet_names[0]:
                    target_sheet = sheet
                    break
            
            if not target_sheet:
                target_sheet = sheet_names[0]
            
            df = pd.read_excel(filepath, sheet_name=target_sheet)
            logger.info(f"Loaded {len(df)} rows from sheet '{target_sheet}'")
            
            df = self._normalize_bank_data(df)
            return df
            
        except Exception as e:
            logger.error(f"Error reading bank file: {e}")
            raise
    
    def read_rn_card_file(self, filepath: str) -> pd.DataFrame:
        """Read RN-Card file and normalize it."""
        logger.info(f"Reading RN-Card file: {filepath}")
        
        try:
            xl = pd.ExcelFile(filepath)
            sheet_names = xl.sheet_names
            logger.info(f"Sheets found: {sheet_names}")
            
            target_sheet = None
            for sheet in sheet_names:
                if 'Name' in sheet or 'Лист' in sheet or sheet == sheet_names[0]:
                    target_sheet = sheet
                    break
            
            if not target_sheet:
                target_sheet = sheet_names[0]
            
            df = pd.read_excel(filepath, sheet_name=target_sheet)
            logger.info(f"Loaded {len(df)} rows from sheet '{target_sheet}'")
            
            df = self._normalize_rn_data(df)
            return df
            
        except Exception as e:
            logger.error(f"Error reading RN-Card file: {e}")
            raise
    
    def _normalize_bank_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize bank data columns."""
        columns_mapping = {}
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if 'дата' in col_lower or 'date' in col_lower:
                columns_mapping['Date'] = col
            elif 'сумма' in col_lower or 'amount' in col_lower or 'сум' in col_lower:
                columns_mapping['Amount'] = col
            elif 'контрагент' in col_lower or 'плательщик' in col_lower or 'получатель' in col_lower or 'название' in col_lower:
                columns_mapping['Contragent'] = col
            elif 'назначение' in col_lower or 'payment' in col_lower:
                columns_mapping['Purpose'] = col
            elif 'номер' in col_lower or 'документ' in col_lower or '№' in col_lower:
                columns_mapping['DocNumber'] = col
        
        if columns_mapping:
            df = df.rename(columns={v: k for k, v in columns_mapping.items()})
        
        for col in ['Date', 'Amount', 'Contragent', 'Purpose', 'DocNumber']:
            if col not in df.columns:
                df[col] = None
        
        return df
    
    def _normalize_rn_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize RN-Card data columns."""
        columns_mapping = {}
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if 'дата' in col_lower or 'date' in col_lower:
                columns_mapping['Date'] = col
            elif 'сумма' in col_lower or 'amount' in col_lower or 'сум' in col_lower:
                columns_mapping['Amount'] = col
            elif 'контрагент' in col_lower or 'покупатель' in col_lower or 'клиент' in col_lower:
                columns_mapping['Contragent'] = col
            elif 'товар' in col_lower or 'услуг' in col_lower:
                columns_mapping['Product'] = col
            elif 'накладная' in col_lower or 'номер' in col_lower:
                columns_mapping['DocNumber'] = col
        
        if columns_mapping:
            df = df.rename(columns={v: k for k, v in columns_mapping.items()})
        
        for col in ['Date', 'Amount', 'Contragent', 'Product', 'DocNumber']:
            if col not in df.columns:
                df[col] = None
        
        return df

    def get_sample_files(self, directory: str) -> List[str]:
        """Get list of Excel files in a directory."""
        if not os.path.exists(directory):
            logger.warning(f"Directory not found: {directory}")
            return []
        
        files = []
        for f in os.listdir(directory):
            if f.endswith(('.xlsx', '.xls', '.xlsm')):
                files.append(os.path.join(directory, f))
        
        return sorted(files)
