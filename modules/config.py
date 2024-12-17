import logging
import yaml
from yaml.loader import SafeLoader
import os
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Define base projects directory
BASE_PROJECTS_DIR = Path.cwd() / "projects"
BASE_PROJECTS_DIR.mkdir(exist_ok=True)
logger.info(f"Base projects directory set to: {BASE_PROJECTS_DIR.resolve()}")

# Load configuration
config = {}
try:
    with open('config.yaml', 'r', encoding='utf-8') as file:
        config = yaml.load(file, Loader=SafeLoader)
    logger.info("Configuration loaded successfully.")
except FileNotFoundError:
    logger.error("Configuration file 'config.yaml' not found.")
