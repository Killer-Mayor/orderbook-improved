"""Configuration module for the orderbook application."""
import os
from typing import Optional
import logging

logger = logging.getLogger(__name__)


class Config:
    """Base configuration class."""
    
    # Flask settings
    SECRET_KEY = os.environ.get('FLASK_SECRET')
    
    # Google Sheets settings
    SHEET_ID = os.environ.get('SHEET_ID')
    SERVICE_ACCOUNT_FILE = os.environ.get('SERVICE_ACCOUNT_FILE', 'service_account.json')
    SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')
    
    # Worksheet names
    MAIN_WORKSHEET_NAME = os.environ.get('MAIN_WORKSHEET_NAME', 'orders')
    PRODUCT_WORKSHEET_NAME = os.environ.get('PRODUCT_WORKSHEET_NAME', 'products')
    COMPANY_WORKSHEET_NAME = os.environ.get('COMPANY_WORKSHEET_NAME', 'companies')
    BRAND_WORKSHEET_NAME = os.environ.get('BRAND_WORKSHEET_NAME', 'brands')
    DISPATCH_WORKSHEET_NAME = os.environ.get('DISPATCH_WORKSHEET_NAME', 'dispatch')
    
    # Application settings
    MAX_RECENT_ORDERS = 50
    DEFAULT_GST_RATE = 0.05  # 5%
    
    # Logging
    LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO')
    
    @classmethod
    def validate(cls) -> tuple[bool, list[str]]:
        """Validate required configuration variables.
        
        Returns:
            Tuple of (is_valid, list of missing variables)
        """
        missing = []
        
        if not cls.SECRET_KEY:
            missing.append('FLASK_SECRET')
        
        if not cls.SHEET_ID:
            missing.append('SHEET_ID')
        
        # Check if we have service account credentials
        has_file = os.path.exists(cls.SERVICE_ACCOUNT_FILE)
        has_json = bool(cls.SERVICE_ACCOUNT_JSON)
        
        if not has_file and not has_json:
            missing.append('SERVICE_ACCOUNT_JSON or service_account.json file')
        
        is_valid = len(missing) == 0
        
        if not is_valid:
            logger.error(f"Missing required configuration: {', '.join(missing)}")
        
        return is_valid, missing


class DevelopmentConfig(Config):
    """Development configuration."""
    DEBUG = True
    TESTING = False


class ProductionConfig(Config):
    """Production configuration."""
    DEBUG = False
    TESTING = False


class TestingConfig(Config):
    """Testing configuration."""
    DEBUG = True
    TESTING = True


def get_config() -> Config:
    """Get configuration based on environment."""
    env = os.environ.get('FLASK_ENV', 'production')
    
    config_map = {
        'development': DevelopmentConfig,
        'production': ProductionConfig,
        'testing': TestingConfig
    }
    
    return config_map.get(env, ProductionConfig)
