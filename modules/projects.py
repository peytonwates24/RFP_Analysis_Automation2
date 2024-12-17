# modules/projects.py
import streamlit as st
import shutil
from .config import BASE_PROJECTS_DIR, logger

def get_user_projects(username):
    user_dir = BASE_PROJECTS_DIR / username
    user_dir.mkdir(parents=True, exist_ok=True)
    projects = [p.name for p in user_dir.iterdir() if p.is_dir()]
    logger.info(f"Retrieved projects for user '{username}': {projects}")
    return projects

def create_project(username, project_name):
    user_dir = BASE_PROJECTS_DIR / username
    project_dir = user_dir / project_name
    if project_dir.exists():
        st.error(f"Project '{project_name}' already exists.")
        logger.warning(f"Attempted to create duplicate project '{project_name}'.")
        return False
    try:
        project_dir.mkdir(parents=True)
        subfolders = ["Baseline", "Round 1 Analysis", "Round 2 Analysis", "Supplier Feedback", "Negotiations"]
        for subfolder in subfolders:
            (project_dir / subfolder).mkdir()
            logger.info(f"Created subfolder '{subfolder}' in project '{project_name}'.")
        st.success(f"Project '{project_name}' created successfully.")
        logger.info(f"User '{username}' created project '{project_name}'.")
        return True
    except Exception as e:
        st.error(f"Error creating project '{project_name}': {e}")
        logger.error(f"Error creating project '{project_name}': {e}")
        return False

def delete_project(username, project_name):
    user_dir = BASE_PROJECTS_DIR / username
    project_dir = user_dir / project_name
    if not project_dir.exists():
        st.error(f"Project '{project_name}' does not exist.")
        logger.warning(f"Attempted to delete non-existent project '{project_name}'.")
        return False
    try:
        shutil.rmtree(project_dir)
        st.success(f"Project '{project_name}' deleted successfully.")
        logger.info(f"User '{username}' deleted project '{project_name}'.")
        return True
    except Exception as e:
        st.error(f"Error deleting project '{project_name}': {e}")
        logger.error(f"Error deleting project '{project_name}': {e}")
        return False
