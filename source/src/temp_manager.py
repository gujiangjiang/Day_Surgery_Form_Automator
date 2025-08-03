# -*- coding: utf-8 -*-
"""
模块功能：管理程序的临时文件和文件夹。
"""
import tempfile
import shutil
import stat
import atexit
from pathlib import Path # 导入Path类

TEMP_DIR = None

def _remove_readonly(func, path, exc_info):
    """
    shutil.rmtree 的错误处理器。
    它会在删除失败时被调用，尝试移除文件的只读属性并重试。
    """
    # exc_info[1] 包含异常实例，我们可以检查它的 winerror 属性
    # 在 Windows 上, "拒绝访问" 的错误代码是 5
    # func in (os.rmdir, os.remove, os.unlink) and ... -> 原始代码依赖os，这里简化为直接检查
    if exc_info and exc_info[1] and getattr(exc_info[1], 'winerror', 0) == 5:
        # 将路径转换为Path对象以使用chmod
        Path(path).chmod(stat.S_IWRITE) # 移除只读属性
        func(path) # 重试删除
    else:
        raise # 如果是其他错误，则重新引发异常

def _cleanup_temp_dir():
    """在程序退出时，自动清理创建的临时文件夹。"""
    global TEMP_DIR
    if TEMP_DIR and TEMP_DIR.exists():
        try:
            # 使用 onerror 处理器来处理因只读属性导致的删除失败
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
        # tempfile.mkdtemp 返回字符串，我们将其转换为Path对象
        TEMP_DIR = Path(tempfile.mkdtemp(prefix="FollowUpApp_"))
        # 注册清理函数，确保程序退出时调用
        atexit.register(_cleanup_temp_dir)
    return TEMP_DIR

def create_temp_read_only_copy(original_path_str):
    """
    将一个文件复制到临时位置，将其设置为只读，并返回新路径。
    返回一个Path对象。
    :param original_path_str: 原始文件的字符串路径。
    """
    if not original_path_str:
        return None
    
    original_path = Path(original_path_str) # 将输入字符串转换为Path对象
    if not original_path.exists():
        return None
    
    temp_dir = get_temp_dir()
    # 使用 / 运算符拼接路径，使用 .name 获取文件名
    temp_path = temp_dir / original_path.name
    
    try:
        shutil.copy2(original_path, temp_path)
        # 将文件权限设置为只读 (stat.S_IREAD)
        temp_path.chmod(stat.S_IREAD)
        return temp_path
    except Exception as e:
        print(f"创建临时文件副本失败 {original_path.name}: {e}")
        return None
