# -*- coding: utf-8 -*-
"""
日志系统配置模块。
负责初始化和配置全局的logging实例。
"""
import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .config import LOGGING_CONFIG

# --- 新增：定义并注册一个自定义的日志级别 ---
# NOTICE级别用于普通的、非高亮的信息性日志
NOTICE_LEVEL_NUM = 25
logging.addLevelName(NOTICE_LEVEL_NUM, "NOTICE")

def notice(self, message, *args, **kws):
    if self.isEnabledFor(NOTICE_LEVEL_NUM):
        self._log(NOTICE_LEVEL_NUM, message, args, **kws)

# 将新方法绑定到Logger类上
logging.Logger.notice = notice
# -----------------------------------------

def setup_logging(base_path):
    """
    配置全局日志系统。
    将日志同时输出到控制台、文件和UI界面（通过自定义处理器）。
    """
    log_level = getattr(logging, LOGGING_CONFIG["log_level"].upper(), logging.INFO)

    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    # 避免重复添加处理器
    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    # --- 1. 配置控制台处理器 (用于开发调试) ---
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # --- 2. 配置文件处理器 (根据配置决定是否启用) ---
    if LOGGING_CONFIG.get("enable_file_logging", True):
        log_dir = base_path / "logs"
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / LOGGING_CONFIG["log_filename"]

        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=LOGGING_CONFIG["log_max_bytes"],
            backupCount=LOGGING_CONFIG["log_backup_count"],
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        logging.info("文件日志已启用。")
    else:
        logging.info("文件日志已禁用。")

    logging.info("日志系统初始化完成。")

