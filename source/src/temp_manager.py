# -*- coding: utf-8 -*-
"""
模块功能：管理程序的临时文件和文件夹。
"""
import os
import tempfile
import shutil
import stat
import atexit

TEMP_DIR = None

def _cleanup_temp_dir():
    """在程序退出时，自动清理创建的临时文件夹。"""
    global TEMP_DIR
    if TEMP_DIR and os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            print(f"成功清理临时文件夹: {TEMP_DIR}")
        except Exception as e:
            print(f"清理临时文件夹 {TEMP_DIR} 时出错: {e}")

def get_temp_dir():
    """
    为当前程序运行会话创建一个唯一的临时文件夹。
    如果文件夹已存在，则直接返回路径。
    """
    global TEMP_DIR
    if TEMP_DIR is None:
        TEMP_DIR = tempfile.mkdtemp(prefix="FollowUpApp_")
        # 注册清理函数，确保程序退出时调用
        atexit.register(_cleanup_temp_dir)
    return TEMP_DIR

def create_temp_read_only_copy(original_path):
    """
    将一个文件复制到临时位置，将其设置为只读，并返回新路径。
    """
    if not original_path or not os.path.exists(original_path):
        return None
    
    temp_dir = get_temp_dir()
    filename = os.path.basename(original_path)
    temp_path = os.path.join(temp_dir, filename)
    
    try:
        shutil.copy2(original_path, temp_path)
        # 将文件权限设置为只读 (stat.S_IREAD)
        os.chmod(temp_path, stat.S_IREAD)
        return temp_path
    except Exception as e:
        print(f"创建临时文件副本失败 {filename}: {e}")
        return None
