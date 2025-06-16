from pydantic import BaseModel
import os


class Config(BaseModel):
    """Excel插件配置类"""
    
    # ===== 基础路径配置 =====
    # Excel文件目录路径 - 支持环境变量自定义
    # 默认指向项目根目录的xlsx文件夹
    excel_folder: str = os.getenv("EXCEL_FOLDER", os.path.join(os.path.dirname(__file__), "..", "..", "xlsx"))
    # 其他可选的路径示例：
    # excel_folder: str = "D:/Desktop/test1/Xlsx/data"  # 自定义绝对路径
    # excel_folder: str = os.path.join(os.path.dirname(__file__), "data")  # 插件目录下的data文件夹
    # excel_folder: str = os.path.dirname(__file__)  # 当前插件目录
    
    # ===== 调试配置 =====
    # 是否启用调试模式
    debug_mode: bool = os.getenv("DEBUG_MODE", "false").lower() == "true"
    
    # ===== Excel格式配置 =====
    # 行高设置（单位：磅）
    row_height: float = float(os.getenv("ROW_HEIGHT", "50.0"))
    
    # A列列宽设置（单位：字符数）
    name_column_width: float = float(os.getenv("NAME_COLUMN_WIDTH", "20.0"))
    
    # ===== 游戏逻辑配置 =====
    # 完成一个周期所需的次数
    completion_count: int = int(os.getenv("COMPLETION_COUNT", "30"))
    
    # ===== 文件导入配置 =====
    # 文件选择超时时间（秒）
    file_selection_timeout: int = int(os.getenv("FILE_SELECTION_TIMEOUT", "30"))
    
    # 私聊文件等待超时时间（秒）
    private_file_timeout: int = int(os.getenv("PRIVATE_FILE_TIMEOUT", "30"))
    
    # 群文件导入配置
    # 最多显示的xlsx文件记录数
    max_xlsx_records: int = int(os.getenv("MAX_XLSX_RECORDS", "5"))
    
    # 群文件导入超时时间（秒）
    group_file_timeout: int = int(os.getenv("GROUP_FILE_TIMEOUT", "30"))
    
    # ===== 查询配置 =====
    # 默认查询显示的最新记录数
    default_lookup_count: int = int(os.getenv("DEFAULT_LOOKUP_COUNT", "3"))