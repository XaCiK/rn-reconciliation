"""Reconciliation engine for bank payments and RN-Card data."""

import pandas as pd
import logging
from typing import Dict, List, Optional, Tuple
from difflib import SequenceMatcher
from datetime import datetime
from .file_processor import FileProcessor

logger = logging.getLogger(__name__)

class ReconciliationEngine:
    """Main reconciliation engine."""
    
    def __init__(self, dictionary_path: Optional[str] = None):
        self.dictionary = {}
        self.contragents_base = {}
        self.bank_data = None
        self.rn_data = None
        self.results = {}
        self._load_default_dictionary()
        
    def _load_default_dictionary(self):
        """Load default contragent dictionary."""
        self.dictionary = {
            'ООО РН-Абхазия': 'РН-Абхазия',
            'РН-Абхазия': 'РН-Абхазия',
            'RN-Abkhazia': 'РН-Абхазия',
        }
        self.contragents_base = {
            'РН-Абхазия': {'inn': '11000831', 'kpp': '111000236'},
        }
        logger.info(f"Loaded {len(self.dictionary)} dictionary entries")
    
    def set_data(self, bank_df: pd.DataFrame, rn_df: pd.DataFrame):
        """Set data for reconciliation."""
        self.bank_data = bank_df
        self.rn_data = rn_df
        logger.info(f"Bank: {len(bank_df)} rows, RN: {len(rn_df)} rows")
    
    def normalize_contragent(self, name: str) -> str:
        """Normalize contragent name using dictionary with fuzzy matching."""
        if not name or pd.isna(name):
            return None
        
        name = str(name).strip()
        
        # Exact match
        if name in self.dictionary:
            return self.dictionary[name]
        
        # Fuzzy match
        best_match = None
        best_ratio = 0.0
        for key, value in self.dictionary.items():
            ratio = SequenceMatcher(None, name.lower(), key.lower()).ratio()
            if ratio > best_ratio and ratio > 0.7:
                best_ratio = ratio
                best_match = value
        
        if best_match:
            logger.debug(f"Fuzzy match: '{name}' -> '{best_match}' ({best_ratio:.2f})")
            return best_match
        
        return name
    
    def run(self) -> Dict:
        """Run full reconciliation."""
        if self.bank_data is None or self.rn_data is None:
            logger.error("Data not loaded")
            return {"error": "Data not loaded"}
        
        logger.info("Starting reconciliation...")
        results = {
            'timestamp': datetime.now().isoformat(),
            'bank_count': len(self.bank_data),
            'rn_count': len(self.rn_data),
            'matches': [],
            'unmatched_bank': [],
            'unmatched_rn': [],
            'summary': {}
        }
        
        # Step 1: Normalize contragents
        self._normalize_contragents()
        
        # Step 2: Find matches
        matches, unmatched_bank, unmatched_rn = self._find_matches()
        
        results['matches'] = matches
        results['unmatched_bank'] = unmatched_bank
        results['unmatched_rn'] = unmatched_rn
        
        # Step 3: Link analysis
        results['link_analysis'] = self._analyze_links(matches)
        
        # Step 4: Compare with base
        results['base_comparison'] = self._compare_with_base()
        
        # Step 5: Summary
        results['summary'] = {
            'total_matches': len(matches),
            'unmatched_bank': len(unmatched_bank),
            'unmatched_rn': len(unmatched_rn),
            'match_rate_bank': len(matches) / len(self.bank_data) * 100 if len(self.bank_data) > 0 else 0,
            'match_rate_rn': len(matches) / len(self.rn_data) * 100 if len(self.rn_data) > 0 else 0,
        }
        
        logger.info(f"Reconciliation complete. Matches: {len(matches)}")
        return results
    
    def _normalize_contragents(self):
        """Normalize contragent names in both datasets."""
        if 'Contragent' in self.bank_data.columns:
            self.bank_data['Normalized'] = self.bank_data['Contragent'].apply(
                lambda x: self.normalize_contragent(x)
            )
        
        if 'Contragent' in self.rn_data.columns:
            self.rn_data['Normalized'] = self.rn_data['Contragent'].apply(
                lambda x: self.normalize_contragent(x)
            )
    
    def _find_matches(self) -> Tuple[List[Dict], List[Dict], List[Dict]]:
        """Find matches between bank and RN data."""
        matches = []
        unmatched_bank = []
        unmatched_rn = []
        
        bank_cols = self.bank_data.columns.tolist()
        rn_cols = self.rn_data.columns.tolist()
        
        bank_contragent_col = 'Normalized' if 'Normalized' in bank_cols else 'Contragent'
        rn_contragent_col = 'Normalized' if 'Normalized' in rn_cols else 'Contragent'
        
        self.bank_data['Amount_num'] = pd.to_numeric(self.bank_data['Amount'], errors='coerce')
        self.rn_data['Amount_num'] = pd.to_numeric(self.rn_data['Amount'], errors='coerce')
        
        bank_sorted = self.bank_data.sort_values(by=bank_contragent_col, na_position='last')
        rn_sorted = self.rn_data.sort_values(by=rn_contragent_col, na_position='last')
        
        used_rn_indices = set()
        
        for bank_idx, bank_row in bank_sorted.iterrows():
            bank_contragent = bank_row.get(bank_contragent_col)
            bank_amount = bank_row.get('Amount_num')
            
            if pd.isna(bank_contragent) or pd.isna(bank_amount):
                unmatched_bank.append(bank_row.to_dict())
                continue
            
            best_match = None
            best_diff = float('inf')
            
            for rn_idx, rn_row in rn_sorted.iterrows():
                if rn_idx in used_rn_indices:
                    continue
                
                rn_contragent = rn_row.get(rn_contragent_col)
                rn_amount = rn_row.get('Amount_num')
                
                if pd.isna(rn_contragent) or pd.isna(rn_amount):
                    continue
                
                if str(bank_contragent).strip().lower() == str(rn_contragent).strip().lower():
                    diff = abs(bank_amount - rn_amount)
                    if diff < best_diff and diff < 1.0:
                        best_diff = diff
                        best_match = (rn_idx, rn_row)
            
            if best_match:
                rn_idx, rn_row = best_match
                used_rn_indices.add(rn_idx)
                
                match = {
                    'bank': bank_row.to_dict(),
                    'rn': rn_row.to_dict(),
                    'difference': best_diff,
                    'contragent': bank_contragent
                }
                matches.append(match)
            else:
                unmatched_bank.append(bank_row.to_dict())
        
        for rn_idx, rn_row in rn_sorted.iterrows():
            if rn_idx not in used_rn_indices:
                unmatched_rn.append(rn_row.to_dict())
        
        return matches, unmatched_bank, unmatched_rn
    
    def _analyze_links(self, matches: List[Dict]) -> Dict:
        """Analyze links between payments and invoices."""
        link_analysis = {
            'total_linked': len(matches),
            'by_contragent': {},
            'by_date': {}
        }
        
        for match in matches:
            bank = match.get('bank', {})
            rn = match.get('rn', {})
            contragent = match.get('contragent', 'Unknown')
            
            if contragent not in link_analysis['by_contragent']:
                link_analysis['by_contragent'][contragent] = []
            link_analysis['by_contragent'][contragent].append({
                'bank_amount': bank.get('Amount', 0),
                'rn_amount': rn.get('Amount', 0),
                'difference': match.get('difference', 0)
            })
            
            date_key = str(bank.get('Date', 'unknown'))[:10] if 'Date' in bank else 'unknown'
            if date_key not in link_analysis['by_date']:
                link_analysis['by_date'][date_key] = 0
            link_analysis['by_date'][date_key] += 1
        
        return link_analysis
    
    def _compare_with_base(self) -> Dict:
        """Compare contragents with base."""
        comparison = {
            'in_base': [],
            'not_in_base': [],
            'total_checked': 0
        }
        
        all_contragents = set()
        if 'Normalized' in self.bank_data.columns:
            all_contragents.update(self.bank_data['Normalized'].dropna().unique())
        if 'Normalized' in self.rn_data.columns:
            all_contragents.update(self.rn_data['Normalized'].dropna().unique())
        
        for contragent in all_contragents:
            comparison['total_checked'] += 1
            if contragent in self.contragents_base:
                comparison['in_base'].append({
                    'name': contragent,
                    'data': self.contragents_base[contragent]
                })
            else:
                comparison['not_in_base'].append(contragent)
        
        logger.info(f"Checked {comparison['total_checked']} contragents, {len(comparison['in_base'])} in base")
        return comparison
