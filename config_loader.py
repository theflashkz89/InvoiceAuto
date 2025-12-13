"""
配置加载模块
从 config.ini 文件读取配置信息
"""

import os
import configparser
from pathlib import Path


def load_config():
    """
    加载配置文件
    
    返回:
        configparser.ConfigParser: 配置对象
        
    异常:
        如果配置文件不存在，会抛出 FileNotFoundError
    """
    config = configparser.ConfigParser()
    config_file = os.path.join(os.path.dirname(__file__), 'config.ini')
    
    if not os.path.exists(config_file):
        raise FileNotFoundError(
            f"配置文件不存在: {config_file}\n"
            f"请复制 config.example.ini 为 config.ini 并填写你的配置信息。"
        )
    
    config.read(config_file, encoding='utf-8')
    return config


def get_email_config():
    """
    获取邮箱配置
    
    返回:
        tuple: (mail_user, mail_pass)
    """
    config = load_config()
    mail_user = config.get('EMAIL', 'mail_user', fallback='')
    mail_pass = config.get('EMAIL', 'mail_pass', fallback='')
    
    if not mail_user or not mail_pass:
        raise ValueError("邮箱配置不完整，请检查 config.ini 中的 [EMAIL] 部分")
    
    return mail_user, mail_pass


def get_api_key():
    """
    获取 API Key
    
    返回:
        str: API Key
    """
    config = load_config()
    api_key = config.get('API', 'api_key', fallback='')
    
    if not api_key:
        raise ValueError("API Key 未配置，请检查 config.ini 中的 [API] 部分")
    
    return api_key

