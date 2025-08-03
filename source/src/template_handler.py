# -*- coding: utf-8 -*-
"""
模块功能：定位并提供内置模板文件的正确路径。
"""
import sys
from pathlib import Path # 导入Path类

def _get_base_path():
    """
    确定应用的根目录（即main.py所在的目录），
    兼容源码运行和PyInstaller打包后的情况。
    返回一个Path对象。
    """
    if getattr(sys, 'frozen', False):
        # 如果程序被打包成一个文件, sys.executable是exe的路径
        return Path(sys.executable).parent
    else:
        # 如果从源码运行, 此文件在 src/template_handler.py
        # 我们需要找到项目根目录，即向上两级
        # Path(__file__) -> .../src/template_handler.py
        # .parent -> .../src/
        # .parent -> .../ (项目根目录)
        return Path(__file__).resolve().parent.parent

def get_template_path(template_filename):
    """
    根据模板文件名，构建并返回其完整的绝对路径。
    返回一个Path对象。
    """
    if not template_filename:
        return None
    
    base_path = _get_base_path()
    return base_path / 'templates' / template_filename

# 我们还需要一个函数来定位assets文件夹，main.py会用到
def get_asset_path(asset_filename):
    """
    根据资源文件名，构建并返回其完整的绝对路径。
    返回一个Path对象。
    """
    if not asset_filename:
        return None
        
    base_path = _get_base_path()
    return base_path / 'assets' / asset_filename
