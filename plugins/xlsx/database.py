#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import os
import datetime
from typing import List, Tuple, Optional
from .config import Config

class DatabaseManager:
    """数据库管理类"""
    
    def __init__(self):
        self.config = Config()
        self.db_path = os.path.join(self.config.excel_folder, "records.db")
        self.init_database()
    
    def init_database(self):
        """初始化数据库"""
        # 确保目录存在
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建游戏表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS games (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建用户表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                game_id INTEGER NOT NULL,
                cycle INTEGER DEFAULT 1,
                is_completed BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (game_id) REFERENCES games (id),
                UNIQUE(name, game_id, cycle)
            )
        ''')
        
        # 创建记录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                record_date TEXT NOT NULL,
                count INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        conn.commit()
        conn.close()
        
        if self.config.debug_mode:
            print(f"数据库初始化完成: {self.db_path}")
    
    def add_game(self, game_name: str) -> int:
        """添加游戏，返回游戏ID"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('INSERT INTO games (name) VALUES (?)', (game_name,))
            game_id = cursor.lastrowid
            conn.commit()
            return game_id
        except sqlite3.IntegrityError:
            # 游戏已存在，获取ID
            cursor.execute('SELECT id FROM games WHERE name = ?', (game_name,))
            return cursor.fetchone()[0]
        finally:
            conn.close()
    
    def get_game_id(self, game_name: str) -> Optional[int]:
        """获取游戏ID"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT id FROM games WHERE name = ?', (game_name,))
        result = cursor.fetchone()
        conn.close()
        
        return result[0] if result else None
    
    def add_user(self, username: str, game_id: int, cycle: int = 1) -> int:
        """添加用户，返回用户ID"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute(
                'INSERT INTO users (name, game_id, cycle) VALUES (?, ?, ?)',
                (username, game_id, cycle)
            )
            user_id = cursor.lastrowid
            conn.commit()
            return user_id
        except sqlite3.IntegrityError:
            # 用户已存在，获取ID
            cursor.execute(
                'SELECT id FROM users WHERE name = ? AND game_id = ? AND cycle = ?',
                (username, game_id, cycle)
            )
            return cursor.fetchone()[0]
        finally:
            conn.close()
    
    def get_user_id(self, username: str, game_id: int, cycle: int = 1) -> Optional[int]:
        """获取用户ID"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            'SELECT id FROM users WHERE name = ? AND game_id = ? AND cycle = ?',
            (username, game_id, cycle)
        )
        result = cursor.fetchone()
        conn.close()
        
        return result[0] if result else None
    
    def add_record(self, user_id: int, record_date: str, count: int):
        """添加记录"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            'INSERT INTO records (user_id, record_date, count) VALUES (?, ?, ?)',
            (user_id, record_date, count)
        )
        conn.commit()
        conn.close()
    
    def get_user_records(self, username: str, game_id: int, cycle: int = 1) -> List[Tuple[str, int]]:
        """获取用户的所有记录"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT r.record_date, r.count
            FROM records r
            JOIN users u ON r.user_id = u.id
            WHERE u.name = ? AND u.game_id = ? AND u.cycle = ?
            ORDER BY r.id
        ''', (username, game_id, cycle))
        
        result = cursor.fetchall()
        conn.close()
        return result
    
    def get_user_latest_count(self, username: str, game_id: int, cycle: int = 1) -> int:
        """获取用户最新的计数"""
        records = self.get_user_records(username, game_id, cycle)
        if not records:
            return 0
        return records[-1][1]  # 返回最后一条记录的count
    
    def complete_user_cycle(self, username: str, game_id: int, cycle: int = 1):
        """标记用户周期完成"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            'UPDATE users SET is_completed = TRUE WHERE name = ? AND game_id = ? AND cycle = ?',
            (username, game_id, cycle)
        )
        conn.commit()
        conn.close()
    
    def import_from_excel_data(self, game_name: str, excel_data: List[List[str]]):
        """从Excel数据导入到数据库"""
        game_id = self.add_game(game_name)
        
        if self.config.debug_mode:
            print(f"开始导入游戏: {game_name} (ID: {game_id})")
        
        imported_count = 0
        
        for row_data in excel_data:
            if not row_data or not row_data[0]:  # 跳过空行
                continue
            
            username = str(row_data[0]).strip()
            if not username:
                continue
            
            # 解析用户名和周期
            cycle = 1
            original_username = username
            
            # 检查是否有周期标记，如 "用户名(2)"
            if '(' in username and username.endswith(')'):
                try:
                    base_name, cycle_part = username.rsplit('(', 1)
                    cycle = int(cycle_part.rstrip(')'))
                    username = base_name
                except (ValueError, IndexError):
                    pass  # 如果解析失败，使用原始用户名
            
            # 获取或创建用户
            user_id = self.get_user_id(username, game_id, cycle)
            if not user_id:
                user_id = self.add_user(username, game_id, cycle)
            
            # 处理记录数据（从第二列开始）
            for record_data in row_data[1:]:
                if not record_data or record_data in ['', '无', 'NaN', None]:
                    continue
                
                record_str = str(record_data).strip()
                if not record_str or record_str == '无':
                    continue
                  # 解析记录格式：MM-DD_次数
                try:
                    if '_' in record_str:
                        date_part, count_part = record_str.split('_', 1)
                        # 处理特殊情况，如 "5-13_30(续)"
                        count_part = count_part.split('(')[0].split('完')[0]
                        count = int(count_part)
                          # 添加记录
                        self.add_record(user_id, date_part, count)
                        imported_count += 1
                        
                        # 检查是否达到完成次数
                        if count >= self.config.completion_count:
                            self.complete_user_cycle(username, game_id, cycle)
                            
                except (ValueError, IndexError) as e:
                    if self.config.debug_mode:
                        print(f"解析记录失败: {record_str}, 错误: {e}")
                    continue
        
        if self.config.debug_mode:
            print(f"导入完成，共导入 {imported_count} 条记录")
        
        return imported_count
    
    def get_games_list(self) -> List[Tuple[str]]:
        """获取所有游戏列表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT name FROM games ORDER BY created_at')
        result = cursor.fetchall()
        conn.close()
        return result
    
    def add_user_record(self, username: str, game_name: str, count: int = 1) -> str:
        """为用户添加指定次数的记录"""
        game_id = self.get_game_id(game_name)
        if not game_id:
            return f"❌ 游戏 {game_name} 不存在"
        
        # 查找用户的最新周期
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT cycle, is_completed FROM users 
            WHERE name = ? AND game_id = ? 
            ORDER BY cycle DESC LIMIT 1
        ''', (username, game_id))
        
        result = cursor.fetchone()
        conn.close()
        
        if result:
            cycle, is_completed = result
            if is_completed:
                # 如果当前周期已完成，创建新周期
                cycle += 1
        else:
            # 新用户
            cycle = 1
        
        # 获取或创建用户
        user_id = self.get_user_id(username, game_id, cycle)
        if not user_id:
            user_id = self.add_user(username, game_id, cycle)
        
        # 获取当前计数
        current_count = self.get_user_latest_count(username, game_id, cycle)
        
        # 添加记录（支持批量添加）
        records_added = []
        total_new_count = current_count
        
        for i in range(count):
            total_new_count += 1
            # 获取当前日期
            today = datetime.datetime.now().strftime("%m-%d")
              # 添加记录
            self.add_record(user_id, today, total_new_count)
            records_added.append(f"{today}_{total_new_count}")
            
            # 检查是否达到完成次数
            if total_new_count >= self.config.completion_count:
                self.complete_user_cycle(username, game_id, cycle)
                break
        
        # 生成结果消息
        if count == 1:
            # 单次记录的简洁消息
            if total_new_count >= self.config.completion_count:
                result_msg = f"✅ 已更新 {username} 的记录: {records_added[-1]} 🎉 恭喜完成{self.config.completion_count}次！"
            else:
                result_msg = f"✅ 已更新 {username} 的记录: {records_added[-1]}"
        else:
            # 批量记录的详细消息
            if total_new_count >= self.config.completion_count:
                result_msg = f"✅ 已为 {username} 添加 {len(records_added)} 条记录\n"
                result_msg += f"记录: {', '.join(records_added)}\n"
                result_msg += f"🎉 恭喜完成{self.config.completion_count}次！"
            else:
                result_msg = f"✅ 已为 {username} 添加 {len(records_added)} 条记录\n"
                result_msg += f"记录: {', '.join(records_added)}\n"
                result_msg += f"当前进度: {total_new_count}/{self.config.completion_count}"
        
        return result_msg
