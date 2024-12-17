# modules/authentication.py
import bcrypt
from .config import config, logger

def authenticate_user(username, password):
    if 'credentials' in config and 'usernames' in config['credentials']:
        if username in config['credentials']['usernames']:
            stored_hashed_password = config['credentials']['usernames'][username]['password']
            if bcrypt.checkpw(password.encode('utf-8'), stored_hashed_password.encode('utf-8')):
                logger.info(f"User {username} authenticated successfully.")
                return True
            else:
                logger.warning(f"Incorrect password for user {username}.")
                return False
        else:
            logger.warning(f"Unknown username: {username}")
            return False
    else:
        logger.error("Authentication configuration is missing.")
        return False
