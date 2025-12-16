"""Google Sheets service for orderbook management.

This module provides a clean interface to Google Sheets API with proper
error handling, logging, and data validation.
"""

import os
import datetime
import logging
from typing import List, Dict, Any, Optional

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound

from config import get_config

logger = logging.getLogger(__name__)


class SheetsServiceError(Exception):
    """Base exception for SheetsService errors."""
    pass


class SheetsService:
    """Service class for Google Sheets operations."""
    
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive'
    ]
    
    def __init__(self, config=None):
        """Initialize the Sheets service.
        
        Args:
            config: Configuration object (uses get_config() if None)
            
        Raises:
            SheetsServiceError: If initialization fails
        """
        self.config = config or get_config()
        self.client = None
        self.spreadsheet = None
        self.sheet = None
        
        try:
            self._initialize_client()
            self._initialize_spreadsheet()
            self._ensure_worksheets()
        except Exception as e:
            logger.error(f"Failed to initialize SheetsService: {e}")
            raise SheetsServiceError(f"Initialization failed: {e}") from e
    
    def _initialize_client(self):
        """Initialize the gspread client with credentials."""
        # Check if we have SERVICE_ACCOUNT_JSON environment variable
        if self.config.SERVICE_ACCOUNT_JSON:
            logger.info("Using SERVICE_ACCOUNT_JSON from environment")
            import json
            import tempfile
            
            # Write JSON to temporary file
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as f:
                f.write(self.config.SERVICE_ACCOUNT_JSON)
                temp_path = f.name
            
            try:
                creds = Credentials.from_service_account_file(temp_path, scopes=self.SCOPES)
            finally:
                os.unlink(temp_path)
        else:
            # Use file path
            if not os.path.exists(self.config.SERVICE_ACCOUNT_FILE):
                raise FileNotFoundError(
                    f"Service account file not found: {self.config.SERVICE_ACCOUNT_FILE}"
                )
            logger.info(f"Using service account file: {self.config.SERVICE_ACCOUNT_FILE}")
            creds = Credentials.from_service_account_file(
                self.config.SERVICE_ACCOUNT_FILE, 
                scopes=self.SCOPES
            )
        
        self.client = gspread.authorize(creds)
        logger.info("Google Sheets client initialized successfully")
    
    def _initialize_spreadsheet(self):
        """Open the target spreadsheet."""
        try:
            self.spreadsheet = self.client.open_by_key(self.config.SHEET_ID)
            logger.info(f"Opened spreadsheet: {self.spreadsheet.title}")
        except SpreadsheetNotFound:
            raise SheetsServiceError(
                f"Spreadsheet not found with ID: {self.config.SHEET_ID}. "
                "Check SHEET_ID and service account permissions."
            )
    
    def _ensure_worksheets(self):
        """Ensure required worksheets exist."""
        # Ensure main orders worksheet
        try:
            self.sheet = self.spreadsheet.worksheet(self.config.MAIN_WORKSHEET_NAME)
            logger.info(f"Using existing worksheet: {self.config.MAIN_WORKSHEET_NAME}")
        except WorksheetNotFound:
            logger.info(f"Creating worksheet: {self.config.MAIN_WORKSHEET_NAME}")
            self.sheet = self.spreadsheet.add_worksheet(
                title=self.config.MAIN_WORKSHEET_NAME, 
                rows=1000, 
                cols=20
            )
            headers = ['Order Number', 'Date', 'Company', 'Product', 'Brand', 'Quantity', 'Price']
            self.sheet.append_row(headers)
        
        # Ensure Order Number column exists
        self._ensure_order_number_column()
    
    def _ensure_order_number_column(self):
        """Ensure the Order Number column exists as first column."""
        try:
            headers = self.sheet.row_values(1)
            if headers and 'order number' not in [h.strip().lower() for h in headers]:
                logger.info("Adding Order Number column")
                all_vals = self.sheet.get_all_values()
                new_headers = ['Order Number'] + (all_vals[0] if all_vals else headers)
                new_rows = [[''] + r for r in (all_vals[1:] if len(all_vals) > 1 else [])]
                
                self.sheet.clear()
                if new_rows:
                    self.sheet.update('A1', [new_headers] + new_rows)
                else:
                    self.sheet.update('A1', [new_headers])
        except Exception as e:
            logger.warning(f"Could not ensure Order Number column: {e}")
    
    def load_lists(self) -> Dict[str, List[str]]:
        """Load product, company and brand lists from worksheets.
        
        Returns:
            Dict with keys 'products', 'companies', 'brands'
        """
        result = {'products': [], 'companies': [], 'brands': []}
        
        worksheets = {
            'products': self.config.PRODUCT_WORKSHEET_NAME,
            'companies': self.config.COMPANY_WORKSHEET_NAME,
            'brands': self.config.BRAND_WORKSHEET_NAME
        }
        
        for key, ws_name in worksheets.items():
            try:
                ws = self.spreadsheet.worksheet(ws_name)
                vals = ws.col_values(1)
                result[key] = [v for v in vals[1:] if v]  # Skip header
                logger.debug(f"Loaded {len(result[key])} {key}")
            except WorksheetNotFound:
                logger.warning(f"Worksheet '{ws_name}' not found, using empty list for {key}")
            except Exception as e:
                logger.error(f"Error loading {key} from '{ws_name}': {e}")
        
        return result
    
    def add_order(
        self, 
        company: str, 
        product: str, 
        quantity: int, 
        price: float, 
        brand: str = '', 
        gst_adjusted: bool = False, 
        order_number: Optional[str] = None
    ) -> None:
        """Add a new order to the spreadsheet.
        
        Args:
            company: Company name
            product: Product name
            quantity: Order quantity
            price: Unit price
            brand: Brand name (optional)
            gst_adjusted: Whether price has GST removed
            order_number: Specific order number (auto-generated if None)
            
        Raises:
            SheetsServiceError: If adding order fails
        """
        try:
            date_str = datetime.date.today().isoformat()
            
            # Generate order number if not provided
            if order_number is None:
                order_number = self._generate_next_order_number()
            
            data_row = [order_number, date_str, company, product, brand or '', int(quantity), float(price)]
            
            # Try to insert with formula copying for efficiency
            target_row = self._find_empty_product_row()
            
            if target_row:
                self._insert_order_with_formulas(data_row, target_row)
            else:
                self.sheet.append_row(data_row)
                self._copy_formulas_to_last_row()
            
            logger.info(f"Added order #{order_number}: {product} for {company}")
            
        except Exception as e:
            logger.error(f"Failed to add order: {e}")
            raise SheetsServiceError(f"Failed to add order: {e}") from e
    
    def _generate_next_order_number(self) -> str:
        """Generate the next sequential order number."""
        try:
            all_vals = self.sheet.get_all_values()
            existing = []
            
            if len(all_vals) >= 2:
                for r in all_vals[1:]:
                    if len(r) >= 1:
                        try:
                            existing.append(int(str(r[0]).strip()))
                        except (ValueError, TypeError):
                            continue
            
            next_num = max(existing) + 1 if existing else 1
            return str(next_num)
        except Exception as e:
            logger.warning(f"Error generating order number: {e}, using '1'")
            return '1'
    
    def _find_empty_product_row(self) -> Optional[int]:
        """Find first row with empty product cell."""
        try:
            all_vals = self.sheet.get_all_values()
            headers = all_vals[0] if all_vals else []
            
            # Find product column index
            prod_col_idx = None
            for idx, h in enumerate(headers, start=1):
                if h and ('blade' in h.strip().lower() or 'product' in h.strip().lower()):
                    prod_col_idx = idx
                    break
            
            if prod_col_idx and len(all_vals) >= 2:
                for ridx, row in enumerate(all_vals[1:], start=2):
                    cell = row[prod_col_idx - 1] if len(row) >= prod_col_idx else ''
                    if not str(cell).strip():
                        return ridx
        except Exception as e:
            logger.debug(f"Could not find empty product row: {e}")
        
        return None
    
    def _insert_order_with_formulas(self, data_row: list, target_row: int):
        """Insert order and copy formulas from previous row."""
        try:
            range_target = f"'{self.sheet.title}'!B{target_row}:G{target_row}"
            body = {'values': [data_row[1:]]}  # Skip order number, it's in column A
            self.spreadsheet.values_update(
                range_target, 
                params={'valueInputOption': 'USER_ENTERED'}, 
                body=body
            )
            
            # Copy formulas from previous row
            self._copy_formulas(target_row - 1, target_row)
        except Exception as e:
            logger.warning(f"Could not insert with formulas: {e}, falling back to append")
            self.sheet.append_row(data_row)
    
    def _copy_formulas(self, source_row: int, dest_row: int):
        """Copy formulas from source row to destination row."""
        try:
            sheet_id = int(self.sheet._properties.get('sheetId'))
            headers = self.sheet.row_values(1)
            header_len = len(headers)
            data_cols = 7  # Order Number through Price
            
            if data_cols < header_len:
                requests = [{
                    'copyPaste': {
                        'source': {
                            'sheetId': sheet_id,
                            'startRowIndex': source_row - 1,
                            'endRowIndex': source_row,
                            'startColumnIndex': data_cols,
                            'endColumnIndex': header_len
                        },
                        'destination': {
                            'sheetId': sheet_id,
                            'startRowIndex': dest_row - 1,
                            'endRowIndex': dest_row,
                            'startColumnIndex': data_cols,
                            'endColumnIndex': header_len
                        },
                        'pasteType': 'PASTE_FORMULA',
                        'pasteOrientation': 'NORMAL'
                    }
                }]
                self.spreadsheet.batch_update({'requests': requests})
        except Exception as e:
            logger.debug(f"Could not copy formulas: {e}")
    
    def _copy_formulas_to_last_row(self):
        """Copy formulas to the last row after append."""
        try:
            all_vals = self.sheet.get_all_values()
            if len(all_vals) >= 3:  # Need at least 2 data rows
                self._copy_formulas(len(all_vals) - 1, len(all_vals))
        except Exception as e:
            logger.debug(f"Could not copy formulas to last row: {e}")
    
    def get_recent_orders(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Get most recent orders.
        
        Args:
            limit: Maximum number of orders to return
            
        Returns:
            List of order dictionaries (newest first)
        """
        try:
            all_values = self.sheet.get_all_values()
            if not all_values or len(all_values) <= 1:
                return []
            
            headers = all_values[0]
            rows = all_values[1:]
            
            # Convert to dicts
            dicts = []
            for r in rows:
                padded = r + [''] * (len(headers) - len(r))
                entry = {headers[i]: padded[i] for i in range(len(headers))}
                dicts.append(entry)
            
            # Return newest first
            return list(reversed(dicts))[:limit]
        except Exception as e:
            logger.error(f"Error fetching recent orders: {e}")
            return []
    
    def add_dispatch(
        self, 
        company: str, 
        product: str, 
        quantity: int, 
        order_number: str
    ) -> None:
        """Record a dispatch.
        
        Args:
            company: Company name
            product: Product name
            quantity: Dispatched quantity
            order_number: Order number being dispatched
            
        Raises:
            SheetsServiceError: If adding dispatch fails
        """
        try:
            date_str = datetime.date.today().isoformat()
            
            # Ensure dispatch worksheet exists
            try:
                dispatch_ws = self.spreadsheet.worksheet(self.config.DISPATCH_WORKSHEET_NAME)
            except WorksheetNotFound:
                dispatch_ws = self.spreadsheet.add_worksheet(
                    title=self.config.DISPATCH_WORKSHEET_NAME, 
                    rows=1000, 
                    cols=10
                )
                dispatch_ws.append_row(['Date', 'Company', 'Product', 'Quantity', 'Order Number'])
            
            dispatch_ws.append_row([date_str, company, product, int(quantity), order_number])
            logger.info(f"Added dispatch: {quantity} x {product} for order #{order_number}")
            
        except Exception as e:
            logger.error(f"Failed to add dispatch: {e}")
            raise SheetsServiceError(f"Failed to add dispatch: {e}") from e
    
    def get_orders_by_product(self, product: str) -> List[Dict[str, Any]]:
        """Get aggregated orders by product with dispatch tracking.
        
        Args:
            product: Product name to filter by
            
        Returns:
            List of order aggregations with remaining quantities
        """
        try:
            all_orders = self.get_recent_orders(1000)
            filtered = [
                o for o in all_orders 
                if o.get('blade type', '').strip().lower() == product.strip().lower()
            ]
            
            # Aggregate by company and serial
            company_orders = {}
            for o in filtered:
                company = o.get('Party Name', '').strip()
                qty = self._extract_quantity(o)
                serial = o.get('Order Number', '')
                
                if company not in company_orders:
                    company_orders[company] = []
                company_orders[company].append({
                    'quantity': qty,
                    'serial': serial,
                    'dispatched': 0
                })
            
            # Apply dispatches
            dispatch_rows = self._get_dispatch_rows()
            for row in dispatch_rows:
                if len(row) < 5:
                    continue
                
                d_product = row[2].strip().lower()
                if d_product != product.strip().lower():
                    continue
                
                d_company = row[1].strip()
                d_qty = self._parse_int(row[3])
                d_serial = row[4].strip()
                
                if d_company in company_orders:
                    for order in company_orders[d_company]:
                        if str(order.get('serial', '')).strip() == d_serial:
                            order['dispatched'] += d_qty
                            break
            
            # Build result
            result = []
            for company, orders in company_orders.items():
                for order in orders:
                    remaining = max(order['quantity'] - order['dispatched'], 0)
                    if remaining > 0:
                        result.append({
                            'company': company,
                            'ordered': order['quantity'],
                            'dispatched': order['dispatched'],
                            'remaining': remaining,
                            'serial': order['serial']
                        })
            
            return result
        except Exception as e:
            logger.error(f"Error getting orders by product: {e}")
            return []
    
    def get_orders_by_party(self, company: str) -> List[Dict[str, Any]]:
        """Get aggregated orders by company/party.
        
        Args:
            company: Company name to filter by
            
        Returns:
            List of product aggregations with remaining quantities
        """
        try:
            all_orders = self.get_recent_orders(1000)
            
            # Find keys
            product_key, company_key = self._find_product_company_keys(all_orders)
            
            filtered = [
                o for o in all_orders 
                if company_key and o.get(company_key, '').strip().lower() == company.strip().lower()
            ]
            
            orders = []
            for o in filtered:
                product = o.get(product_key, '').strip() if product_key else ''
                qty = self._extract_quantity(o)
                serial = o.get('Order Number', '')
                
                orders.append({
                    'product': product,
                    'ordered': qty,
                    'dispatched': 0,
                    'serial': serial
                })
            
            # Apply dispatches
            dispatch_rows = self._get_dispatch_rows()
            for row in dispatch_rows:
                if len(row) < 5:
                    continue
                
                d_company = row[1].strip()
                if d_company.strip().lower() != company.strip().lower():
                    continue
                
                d_qty = self._parse_int(row[3])
                d_serial = row[4].strip()
                
                for order in orders:
                    if str(order.get('serial', '')).strip() == d_serial:
                        order['dispatched'] += d_qty
                        break
            
            # Filter to remaining > 0
            result = []
            for o in orders:
                remaining = max(o['ordered'] - o['dispatched'], 0)
                if remaining > 0:
                    result.append({
                        'product': o['product'],
                        'ordered': o['ordered'],
                        'dispatched': o['dispatched'],
                        'remaining': remaining,
                        'serial': o['serial']
                    })
            
            return result
        except Exception as e:
            logger.error(f"Error getting orders by party: {e}")
            return []
    
    def get_pivot_data(
        self, 
        product_filter: str = '', 
        party_filter: str = ''
    ) -> Dict[str, Any]:
        """Get pivot table data for products vs parties.
        
        Args:
            product_filter: Comma-separated product filters
            party_filter: Comma-separated party filters
            
        Returns:
            Dict with 'pivot', 'products', 'parties'
        """
        try:
            all_orders = self.get_recent_orders(1000)
            
            # Normalize filters
            prod_filters = self._normalize_filter(product_filter)
            party_filters = self._normalize_filter(party_filter)
            
            # Check for balance column
            balance_key = self._find_balance_key(all_orders)
            
            # Aggregate
            product_orders = {}
            party_orders = {}
            
            for order in all_orders:
                product = order.get('blade type', '').strip()
                company = order.get('Party Name', '').strip()
                
                if not product or not company:
                    continue
                
                # Apply filters
                if prod_filters and not self._matches_filter(product, prod_filters):
                    continue
                if party_filters and not self._matches_filter(company, party_filters):
                    continue
                
                # Get quantity
                qty = self._extract_balance_quantity(order, balance_key)
                
                if qty <= 0:
                    continue
                
                # Track
                if product not in product_orders:
                    product_orders[product] = {}
                if company not in product_orders[product]:
                    product_orders[product][company] = 0
                product_orders[product][company] += qty
                
                if company not in party_orders:
                    party_orders[company] = {}
                if product not in party_orders[company]:
                    party_orders[company][product] = 0
                party_orders[company][product] += qty
            
            # Build pivot
            products = sorted(product_orders.keys())
            parties = sorted(party_orders.keys())
            
            pivot = []
            for party in parties:
                row = []
                for product in products:
                    pending = party_orders.get(party, {}).get(product, 0)
                    row.append(pending)
                pivot.append(row)
            
            return {
                'pivot': pivot,
                'products': products,
                'parties': parties
            }
        except Exception as e:
            logger.error(f"Error generating pivot data: {e}")
            return {'pivot': [], 'products': [], 'parties': []}
    
    # Helper methods
    
    def _extract_quantity(self, order: dict) -> int:
        """Extract quantity from order dict."""
        bal_raw = None
        for k in order.keys():
            if 'balance' in k.strip().lower():
                bal_raw = order.get(k)
                break
        
        if bal_raw is None:
            bal_raw = order.get('quantity', order.get('Quantity', '0'))
        
        return self._parse_int(bal_raw)
    
    def _extract_balance_quantity(self, order: dict, balance_key: Optional[str]) -> int:
        """Extract balance quantity considering balance column."""
        if balance_key:
            return self._parse_int(order.get(balance_key, '0'))
        else:
            return self._extract_quantity(order)
    
    def _parse_int(self, value: Any) -> int:
        """Safely parse int from various types."""
        try:
            return int(float(str(value).replace(',', '') or 0))
        except (ValueError, TypeError):
            return 0
    
    def _get_dispatch_rows(self) -> list:
        """Get all dispatch rows."""
        try:
            dispatch_ws = self.spreadsheet.worksheet(self.config.DISPATCH_WORKSHEET_NAME)
            return dispatch_ws.get_all_values()[1:]  # Skip header
        except Exception:
            return []
    
    def _find_product_company_keys(self, orders: list) -> tuple:
        """Find product and company column keys."""
        if not orders:
            return None, None
        
        product_key = None
        company_key = None
        
        for k in list(orders[0].keys()):
            lk = k.strip().lower()
            if not product_key and ('blade' in lk or 'product' in lk):
                product_key = k
            if not company_key and ('party' in lk or 'company' in lk):
                company_key = k
        
        return product_key, company_key
    
    def _find_balance_key(self, orders: list) -> Optional[str]:
        """Find balance column key."""
        if not orders:
            return None
        
        for h in list(orders[0].keys()):
            if h and h.strip().lower() == 'balance order':
                return h
        return None
    
    def _normalize_filter(self, filter_str: str) -> list:
        """Normalize comma-separated filter string."""
        if not filter_str:
            return []
        if isinstance(filter_str, list):
            return [x.strip().lower() for x in filter_str if x.strip()]
        return [x.strip().lower() for x in filter_str.split(',') if x.strip()]
    
    def _matches_filter(self, value: str, filters: list) -> bool:
        """Check if value matches any filter."""
        value_lower = value.lower()
        return any(f in value_lower for f in filters)
