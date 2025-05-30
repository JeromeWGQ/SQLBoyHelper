# -*- coding: utf-8 -*-
"""
MCP文件服务示例（支持Python 3.9+）
功能：提供文件系统访问能力，包含获取文件列表和写入文件两个工具
引用规范：[1,2,5,7](@ref)
"""

# 导入必要模块（参考网页1、网页2）
import os
from pathlib import Path
from typing import Union
from mcp.server.fastmcp import FastMCP  # 使用官方SDK的FastMCP（网页2）

# 初始化MCP服务实例（服务名显示在客户端）
mcp_service = FastMCP("File-Service")  # （网页2示例）


@mcp_service.tool()
def get_file_list(path: str = "~") -> list:
    """
    获取指定路径下的文件列表（网页1功能扩展）

    参数说明：
        path (str): 目标路径，支持波浪符(~)表示家目录，默认当前用户家目录

    返回：
        list: 按字母排序的文件名列表，包含文件/文件夹

    异常处理：
        自动展开用户目录，路径不存在时返回空列表

    示例：
        >>> get_file_list("~/Documents")
        ['report.pdf', 'projects', 'notes.txt']
    """
    expanded_path = os.path.expanduser(path)  # 处理路径中的~符号（网页1实现）
    try:
        return sorted(os.listdir(expanded_path))
    except FileNotFoundError:
        return []


@mcp_service.tool()
def write_file(file_path: str, content: Union[str, bytes], mode: str = "w") -> bool:
    """
    创建/写入文件（新增功能）

    参数说明：
        file_path (str): 文件绝对路径（建议使用绝对路径）
        content: 文本内容(str)或二进制内容(bytes)
        mode (str): 写入模式，支持：
            'w' - 覆盖写入（默认）
            'a' - 追加写入
            'wb' - 二进制写入

    返回：
        bool: 操作是否成功

    安全机制：
        自动创建父目录，限制最大文件大小为10MB

    示例：
        >>> write_file("/tmp/test.txt", "Hello MCP")
        True
    """
    try:
        # 路径处理（网页5安全建议）
        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        # 写入内容校验（网页6最佳实践）
        if len(content) > 10 * 1024 * 1024:  # 10MB限制
            return False

        # 根据模式写入文件（网页7类型处理）
        with open(path, mode) as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"写入失败: {str(e)}")  # 实际部署建议使用日志模块
        return False


if __name__ == "__main__":
    """
    服务启动配置（参考网页1、网页3传输协议）
    使用stdio通信模式，适合本地开发调试
    生产环境建议改用SSE协议（参考网页3）
    """
    mcp_service.run(
        transport="stdio",  # 标准输入输出通信（网页1配置）
        show_local_echo=True  # 显示本地调试信息
    )