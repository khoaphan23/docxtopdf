import os
import configparser
from pathlib import Path


# Tạo đối tượng configparser
config = configparser.ConfigParser()
# Đọc file config.ini với UTF-8 từ đường dẫn tuyệt đối (../config.ini)
CONFIG_PATH = (Path(__file__).resolve().parent.parent / "config.ini")
config.read(CONFIG_PATH, encoding="utf-8")

# ==================================================
# KHAI BÁO CÁC BIẾN TỪ CONFIG.INI
# ==================================================

# [DEFAULT] Section
APP_NAME = config.get('DEFAULT', 'app_name', fallback='Word to PDF Converter')
VERSION = config.get('DEFAULT', 'version', fallback='1.0.0')
AUTHOR = config.get('DEFAULT', 'author', fallback='thangvk')

# [PATHS] Section
OUTPUT_FOLDER = config.get('PATHS', 'output_folder', fallback='PDF_Output')
DOWNLOADS_FOLDER = config.get('PATHS', 'downloads_folder', fallback='Downloads')
LOG_FOLDER = config.get('PATHS', 'log_folder', fallback='logs')

# [CONVERSION] Section
DEFAULT_METHOD = config.get('CONVERSION', 'default_method', fallback='auto')
BACKUP_METHOD = config.get('CONVERSION', 'backup_method', fallback='docx2pdf')

# [LOGGING] Section
LOG_LEVEL = config.get('LOGGING', 'level', fallback='INFO')
FILE_LOGGING = config.getboolean('LOGGING', 'file_logging', fallback=True)
MAX_FILE_SIZE = config.getint('LOGGING', 'max_file_size', fallback=5242880)  # 5MB
BACKUP_COUNT = config.getint('LOGGING', 'backup_count', fallback=5)

# [UI] Section
WINDOW_WIDTH = config.getint('UI', 'window_width', fallback=600)
WINDOW_HEIGHT = config.getint('UI', 'window_height', fallback=500)
THEME = config.get('UI', 'theme', fallback='clam')

# ==================================================
# COMPUTED PATHS (Đường dẫn được tính toán)
# ==================================================

# Base directory của project (thư mục chứa config.ini)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Các đường dẫn tuyệt đối
OUTPUT_PATH = PROJECT_ROOT / OUTPUT_FOLDER
DOWNLOADS_PATH = PROJECT_ROOT / DOWNLOADS_FOLDER  
LOG_PATH = PROJECT_ROOT / LOG_FOLDER

# ==================================================
# CONSTANTS (Hằng số ứng dụng)
# ==================================================

# Supported file extensions
SUPPORTED_EXTENSIONS = ['.doc', '.docx']

# Conversion methods
CONVERSION_METHODS = {
    'win32com': 'Microsoft Word COM',
    'docx2pdf': 'docx2pdf Library',
    'libreoffice': 'LibreOffice'
}

# Default timeout for conversion (seconds)
CONVERSION_TIMEOUT = 60

# Logging format constant added to fix missing attribute error
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"


# ===== Added by fix =====
WINDOW_TITLE = 'Word to PDF Converter'
SELECT_BUTTON_TEXT = 'Chọn file Word'
CONVERT_BUTTON_TEXT = 'Chuyển sang PDF'
DOWNLOAD_BUTTON_TEXT = 'Mở thư mục PDF'
QUIT_BUTTON_TEXT = 'Thoát'
OUTPUT_FORMAT = 'pdf'
