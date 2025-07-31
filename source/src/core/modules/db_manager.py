# -*- coding: utf-8 -*-
"""
模块功能：管理所有与SQLite数据库的交互。
"""
import sqlite3
from ...config import CONFIG

class DatabaseManager:
    """封装所有数据库操作的类"""
    def __init__(self, log_func):
        self.log = log_func
        self.conn = sqlite3.connect(':memory:')
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _create_tables(self):
        """在数据库中创建所需的表"""
        cur = self.conn.cursor()
        all_cols = CONFIG['column_mapping'].keys()
        cols_definitions = ", ".join([f'"{col}" TEXT' for col in all_cols])
        
        cur.execute(f'CREATE TABLE patients ({cols_definitions})')
        cur.execute('''
            CREATE TABLE surgery_lookup (
                hospital_id_cleaned TEXT,
                name_cleaned TEXT,
                bed_number TEXT,
                PRIMARY KEY (hospital_id_cleaned, name_cleaned) ON CONFLICT REPLACE
            )
        ''')
        self.conn.commit()
        self.log("数据库表结构创建成功。")

    def load_surgery_data(self, records):
        """批量加载手术数据到数据库"""
        if not records:
            return 0
        cur = self.conn.cursor()
        cur.executemany(
            "INSERT INTO surgery_lookup (hospital_id_cleaned, name_cleaned, bed_number) VALUES (?, ?, ?)",
            records
        )
        self.conn.commit()
        return len(records)

    def load_patient_data(self, records):
        """批量加载患者数据到数据库"""
        if not records:
            return 0
        internal_keys = list(CONFIG['column_mapping'].keys())
        placeholders = ", ".join(["?"] * len(internal_keys))
        cols_names = ", ".join([f'"{key}"' for key in internal_keys])
        
        cur = self.conn.cursor()
        cur.executemany(f"INSERT INTO patients ({cols_names}) VALUES ({placeholders})", records)
        self.conn.commit()
        return len(records)

    def query_final_data(self):
        """执行SQL查询以合并数据并筛选出日间手术患者"""
        self.log("正在通过SQL查询合并床号并筛选日间手术患者...")
        cur = self.conn.cursor()
        
        query = f"""
            SELECT
                p.*,
                COALESCE(
                    NULLIF(p.bed_number, ''), 
                    s.bed_number
                ) AS final_bed_number
            FROM
                patients p
            LEFT JOIN
                surgery_lookup s ON 
                    (ltrim(p.hospital_id, '0') = s.hospital_id_cleaned OR p.hospital_id = s.hospital_id_cleaned)
                    AND p.name = s.name_cleaned
            WHERE
                CAST(p.hospital_days AS REAL) <= ?
        """
        
        cur.execute(query, (CONFIG['day_surgery_max_days'],))
        return cur.fetchall()

    def close(self):
        """关闭数据库连接"""
        if self.conn:
            self.conn.close()
            self.log("数据库连接已关闭。")
