# tests/conftest.py

import pytest
import logging
from pathlib import Path

@pytest.fixture
def get_test_data_dir():
    """
    Fixture to provide the base directory for test data.
    Adjust the path according to your project structure.
    """
    return Path(__file__).parent / 'data' / 'test_scenarios' / 'test_1'

@pytest.fixture
def setup_logging():
    """
    Fixture to set up logging for tests.
    """
    logger = logging.getLogger('test_logger')
    logger.setLevel(logging.DEBUG)
    
    # Create console handler with a higher log level
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    
    # Create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    
    # Add the handlers to the logger
    if not logger.handlers:
        logger.addHandler(ch)
    
    return logger
