# -*- coding: utf-8 -*-
"""
模块功能：管理程序的临时文件和文件夹。
"""
import tempfile
import shutil
import stat
import atexit
from pathlib import Path

TEMP_DIR = None
# 用于存储已创建的临时文件路径，以便在清理时进行检查
_temp_files = set()

def _remove_readonly(func, path, exc_info):
    """
    shutil.rmtree 的错误处理器。
    它会在删除失败时被调用，尝试移除文件的只读属性并重试。
    """
    if exc_info and exc_info[1] and getattr(exc_info[1], 'winerror', 0) == 5:
        Path(path).chmod(stat.S_IWRITE)
        func(path)
    else:
        raise

def _cleanup_temp_dir():
    """在程序退出时，自动清理创建的临时文件夹。"""
    global TEMP_DIR
    if TEMP_DIR and TEMP_DIR.exists():
        try:
            shutil.rmtree(TEMP_DIR, onerror=_remove_readonly)
            print(f"成功清理临时文件夹: {TEMP_DIR}")
        except Exception as e:
            print(f"清理临时文件夹 {TEMP_DIR} 时出错: {e}")

def get_temp_dir():
    """
    为当前程序运行会话创建一个唯一的临时文件夹。
    如果文件夹已存在，则直接返回路径。
    返回一个Path对象。
    """
    global TEMP_DIR
    if TEMP_DIR is None:
        TEMP_DIR = Path(tempfile.mkdtemp(prefix="FollowUpApp_"))
        atexit.register(_cleanup_temp_dir)
    return TEMP_DIR

def create_temp_read_only_copy(original_path_str):
    """
    将一个文件复制到临时位置，将其设置为只读，并返回新路径。
    此版本更健壮，能处理上次运行遗留的只读文件。
    返回一个Path对象。
    :param original_path_str: 原始文件的字符串路径。
    """
    if not original_path_str:
        return None
    
    original_path = Path(original_path_str)
    if not original_path.is_file():
        print(f"错误: 原始路径不是一个文件: {original_path}")
        return None
    
    temp_dir = get_temp_dir()
    temp_path = temp_dir / original_path.name
    
    # --- 新增的健壮性逻辑 ---
    # 如果目标临时文件已存在 (可能是上次运行遗留的)
    if temp_path.exists():
        try:
            # 尝试移除只读属性并删除它
            temp_path.chmod(stat.S_IWRITE)  # 移除只读
            temp_path.unlink()              # 删除文件
            print(f"成功清理了遗留的临时文件: {temp_path}")
        except Exception as e:
            # 如果删除失败 (例如，文件仍被另一个程序锁定)
            print(f"清理遗留的临时文件 {temp_path} 失败: {e}")
            # 此时应通知用户并阻止继续操作
            return None

    # --- 原始逻辑 ---
    try:
        # 现在路径应该是干净的，可以安全地复制
        shutil.copy2(original_path, temp_path)
        # 将新文件设置为只读
        temp_path.chmod(stat.S_IREAD)
        return temp_path
    except Exception as e:
        # 用户的错误日志在这里被触发
        print(f"创建临时文件副本失败 {original_path.name}: {e}")
        return None
