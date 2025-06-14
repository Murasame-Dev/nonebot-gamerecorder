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