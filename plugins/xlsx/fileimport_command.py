#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import httpx
import json
import os
import time
from typing import Dict, Any, List, Optional
from nonebot import on_command, on_message
from nonebot.adapters.onebot.v11 import Bot, GroupMessageEvent, MessageEvent
from nonebot.permission import SUPERUSER
from nonebot.matcher import Matcher
from nonebot.exception import FinishedException
from nonebot.rule import Rule

from .config import Config
from .database import DatabaseManager
from .excel_importer import ExcelImporter

# 获取配置
plugin_config = Config()
# 初始化Excel导入器
excel_importer = ExcelImporter()

# 注册fileimport命令
fileimport_handler = on_command("fileimport", priority=5, permission=SUPERUSER)

@fileimport_handler.handle()
async def handle_fileimport(bot: Bot, event: MessageEvent):
    """处理群文件导入命令 - 仅限onebot-v11和群聊"""
    # 检查是否为群聊消息
    if not isinstance(event, GroupMessageEvent):
        await fileimport_handler.finish("❌ 此命令仅支持群聊使用！")
    
    try:
        # 获取群文件列表
        group_id = event.group_id
        await fileimport_handler.send("🔍 正在获取群文件列表...")
        
        # 获取群文件列表
        file_list = await get_group_files(bot, group_id)
        
        if not file_list:
            await fileimport_handler.finish("❌ 未找到群文件或获取失败")
        
        # 筛选xlsx文件
        xlsx_files = []
        for file_info in file_list:
            file_name = file_info.get("name", "")
            if file_name.lower().endswith(('.xlsx', '.xls')):
                xlsx_files.append(file_info)
        
        if not xlsx_files:
            await fileimport_handler.finish("❌ 群文件中没有找到Excel文件(.xlsx/.xls)")
        
        # 限制显示数量
        max_show = min(len(xlsx_files), plugin_config.max_xlsx_records)
        xlsx_files = xlsx_files[:max_show]
        
        # 构建文件选择消息
        file_msg = "📁 找到以下Excel文件:\n"
        for i, file_info in enumerate(xlsx_files, 1):
            file_name = file_info.get("name", "未知文件")
            file_size = file_info.get("size", 0)
            file_size_mb = file_size / (1024 * 1024) if file_size > 0 else 0
            file_msg += f"{i}. {file_name} ({file_size_mb:.2f}MB)\n"
        
        file_msg += f"\n请输入序号(1-{len(xlsx_files)})选择要导入的文件\n"
        file_msg += f"或输入 '我不传了' 取消操作\n"
        file_msg += f"⏰ 超时时间: {plugin_config.group_file_timeout}秒"
        
        await fileimport_handler.send(file_msg)
        
        # 等待用户选择
        selected_file = await wait_for_file_selection(bot, event, xlsx_files)
        
        if selected_file is None:
            await fileimport_handler.finish("⏰ 操作超时或已取消")
          # 下载并导入文件
        await download_and_import_file(bot, event, selected_file)
        
    except FinishedException:
        # 重新抛出FinishedException，这是NoneBot的正常流程控制
        raise
    except Exception as e:
        await fileimport_handler.finish(f"❌ 文件导入失败: {str(e)}")

async def get_group_files(bot: Bot, group_id: int) -> List[Dict[str, Any]]:
    """获取群文件列表"""
    try:
        # 调用OneBot V11 API获取群文件
        result = await bot.call_api("get_group_files", group_id=group_id)
        
        if result and "files" in result:
            return result["files"]
        else:
            return []
    except Exception as e:
        if plugin_config.debug_mode:
            print(f"获取群文件失败: {e}")
        return []

async def wait_for_file_selection(bot: Bot, event: GroupMessageEvent, xlsx_files: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """等待用户选择文件"""
    user_id = event.user_id
    group_id = event.group_id
    
    # 创建一个消息监听器
    from nonebot import on_message
    from nonebot.rule import Rule
    
    def check_user_and_group():
        def _check(temp_event: GroupMessageEvent) -> bool:
            return (isinstance(temp_event, GroupMessageEvent) and 
                   temp_event.user_id == user_id and 
                   temp_event.group_id == group_id)
        return Rule(_check)
    
    # 创建临时消息处理器
    temp_matcher = on_message(rule=check_user_and_group(), temp=True, priority=1)
    
    selected_file = None
    is_cancelled = False
    
    @temp_matcher.handle()
    async def handle_selection(temp_event: GroupMessageEvent):
        nonlocal selected_file, is_cancelled
        
        message = temp_event.get_plaintext().strip()
        
        # 检查取消操作
        if message == "我不传了":
            is_cancelled = True
            await temp_matcher.send("❌ 已取消文件导入操作")
            temp_matcher.destroy()
            return
        
        # 检查是否是数字
        if message.isdigit():
            selection = int(message)
            if 1 <= selection <= len(xlsx_files):
                selected_file = xlsx_files[selection - 1]
                await temp_matcher.send(f"✅ 已选择文件: {xlsx_files[selection - 1].get('name', '未知文件')}")
                temp_matcher.destroy()
                return
            else:
                await temp_matcher.send(f"❌ 序号无效，请输入1-{len(xlsx_files)}之间的数字")
        else:
            await temp_matcher.send("❌ 请输入有效的序号或 '我不传了' 取消操作")
    
    # 等待用户响应或超时
    start_time = time.time()
    timeout = plugin_config.group_file_timeout
    
    while time.time() - start_time < timeout:
        if selected_file is not None or is_cancelled:
            break
        await asyncio.sleep(0.5)
    
    # 清理临时处理器
    temp_matcher.destroy()
    
    if is_cancelled:
        return None
    
    return selected_file

async def download_and_import_file(bot: Bot, event: GroupMessageEvent, file_info: Dict[str, Any]):
    """下载并导入文件"""
    try:
        file_name = file_info.get("name", "unknown.xlsx")
        file_id = file_info.get("id", "")
        
        await fileimport_handler.send(f"📥 正在下载文件: {file_name}")
        
        # 获取文件下载链接
        file_url = await get_file_download_url(bot, file_id)
        
        if not file_url:
            await fileimport_handler.finish("❌ 获取文件下载链接失败")
        
        # 下载文件
        temp_dir = os.path.join(plugin_config.excel_folder, "temp_downloads")
        os.makedirs(temp_dir, exist_ok=True)
        
        temp_file_path = os.path.join(temp_dir, file_name)
        
        async with httpx.AsyncClient() as client:
            response = await client.get(file_url)
            response.raise_for_status()
            
            with open(temp_file_path, 'wb') as f:
                f.write(response.content)
        
        await fileimport_handler.send(f"✅ 文件下载完成，正在导入...")
        
        # 导入文件 - 使用临时文件路径直接导入
        result = excel_importer.import_excel(temp_file_path)
          # 如果导入成功，重新注册游戏命令
        if result.startswith("✅"):
            # 导入成功后需要通知主模块重新注册命令
            from . import __main__
            __main__.register_game_commands()
        
        await fileimport_handler.finish(result)
        
    except FinishedException:
        # 重新抛出FinishedException，这是NoneBot的正常流程控制
        raise
    except Exception as e:
        await fileimport_handler.finish(f"❌ 下载或导入文件失败: {str(e)}")

async def get_file_download_url(bot: Bot, file_id: str) -> Optional[str]:
    """获取文件下载链接"""
    try:
        # 调用OneBot V11 API获取文件信息
        result = await bot.call_api("get_file", file_id=file_id)
        
        if result and "url" in result:
            return result["url"]
        else:
            return None
    except Exception as e:
        if plugin_config.debug_mode:
            print(f"获取文件下载链接失败: {e}")
        return None
