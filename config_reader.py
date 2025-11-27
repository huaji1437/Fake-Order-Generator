import json
from pathlib import Path

from typing import Dict, Any



class ConfigReader:
    """读取和管理配置文件的类"""
    config_file_path: Path

    @classmethod
    def get_default_config_view(cls) -> Dict[str, Any]:
        default_config = {
            "college_name": "烤番薯学院魔法学院请假单",
            "cause" : "？？部门",
            "title_format": "%s",
            "font_settings": {
                "normal_font": "等线",
                "title_font": "等线",
                "time_font": "楷体",
                "content_font": "仿宋",
                "font_size": {
                    "title": 14,
                    "normal": 11,
                    "small": 10,
                    "content": 12
                }
            },
            "table_settings": {
                "header_shading": "D9D9D9",
                "max_columns": 6
            },
            "output_settings": {
                "save_path": "desktop",
                "file_name_format": "{year}年{month}月{day}日_{cause}假单.docx"
            },
            "class_mappings": {
                "网络技术": "网络",
                "计算机应用": "计应",
                "计算机应用技术": "计应",
                "软件技术": "软件",
                "云计算": "云计算",
                "电子竞技": "电竞",
                "人工智能技术": "人工智能",
                "人工智能技术应用": "人工智能",
                "大数据技术": "大数据",
                "数字媒体": "数媒",
                "数字媒体技术": "数媒",
                "数字媒体技术班": "数媒"
            },
            "leave_types": {
                "morning": "早自习",
                "evening": "晚自习"
            }
        }
        return default_config

    def __init__(self, config_file_path: str = "config.json"):
        self.config_file_path = Path(config_file_path)
        self.default_config = self.get_default_config_view()
        self.config = self.load_config()

    def load_config(self) -> Dict[str, Any]:
        """加载配置文件，如果不存在则创建默认配置"""
        if not self.config_file_path.exists():
            self.create_default_config()
            return self.default_config

        try:
            with self.config_file_path.open('r', encoding='utf-8') as f:
                user_config = json.load(f)
                # 合并默认配置和用户配置
                return self._deep_merge(self.default_config, user_config)
        except Exception as e:
            print(f"读取配置文件失败，使用默认配置: {e}")
            self.create_default_config()
            return self.default_config


    def create_default_config(self):
        """创建默认配置文件"""
        try:
            with self.config_file_path.open('w', encoding='utf-8') as f:
                json.dump(self.get_default_config_view(), f, ensure_ascii=False, indent=2)
            print(f"已创建默认配置文件: {self.config_file_path.absolute()}")
        except Exception as e:
            print(f"创建配置文件失败: {e}")

    def _deep_merge(self, base: Dict, update: Dict) -> Dict:
        """深度合并两个字典"""
        result = base.copy()
        for key, value in update.items():
            if isinstance(value, dict) and key in result and isinstance(result[key], dict):
                result[key] = self._deep_merge(result[key], value)
            else:
                result[key] = value
        return result

    def get(self, key: str, default=None):
        """获取配置值"""
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        return value

    def update_config(self, new_config: Dict[str, Any]):
        """更新配置"""
        self.config = self._deep_merge(self.config, new_config)
        try:
            with self.config_file_path.open('w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"更新配置文件失败: {e}")