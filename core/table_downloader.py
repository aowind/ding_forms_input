"""钉钉表格下载模块 - 通过 UI 操作导出 Excel"""

from __future__ import annotations

import asyncio
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# 下载超时（秒）
DOWNLOAD_TIMEOUT = 60


async def download_table_excel(page, download_dir: str | Path) -> Path | None:
    """通过操作 Table 菜单下载当前 Sheet 的 Excel。

    Args:
        page: Playwright Page 对象
        download_dir: 下载保存目录

    Returns:
        下载完成的 xlsx 文件路径，失败返回 None
    """
    download_path = Path(download_dir)

    # 监听下载事件
    async with page.expect_download(timeout=DOWNLOAD_TIMEOUT * 1000) as download_info:
        # 1. 找到 iframe
        iframe_el = await page.query_selector('iframe[src*="spreadsheet"]')
        if not iframe_el:
            logger.error("未找到表格 iframe")
            return None
        box = await iframe_el.bounding_box()
        if not box:
            logger.error("无法获取 iframe 位置")
            return None

        # Playwright bounding_box 返回 dict，不是对象
        bx, by = box["x"], box["y"]

        # 2. 点击 Table 菜单（工具栏第一个按钮）
        table_btn_x = bx + 34
        table_btn_y = by + 88
        await page.mouse.click(table_btn_x, table_btn_y)
        await asyncio.sleep(1.0)

        # 3. 鼠标悬停显示子菜单
        await page.mouse.move(bx + 30, by + 140)
        await asyncio.sleep(1.5)

        # 4. 悬停 Download 选项
        await page.mouse.move(bx + 250, by + 175)
        await asyncio.sleep(2.0)

        # 5. 点击 "Excel (.xlsx, table as a whole)" 选项
        #    Download 子菜单项通常在悬停后出现
        await page.mouse.click(bx + 320, by + 250)
        await asyncio.sleep(1.0)

    # 保存下载的文件
    download = download_info.value
    save_path = download_path / download.suggested_filename
    await download.save_as(str(save_path))
    logger.info("表格已下载: %s", save_path)
    return save_path


async def get_sheet_tabs(page) -> list[str]:
    """获取钉钉表格中的 Sheet 标签页列表。

    通过点击左侧 Sheet 标签区域的 DOM 获取名称。

    Returns:
        Sheet 名称列表
    """
    try:
        iframe_el = await page.query_selector('iframe[src*="spreadsheet"]')
        if not iframe_el:
            return []
        frame = await iframe_el.content_frame()
        if not frame:
            return []

        # 尝试获取 sheet tab 元素
        tabs = await frame.eval_on_selector_all(
            ".sheet-tab, .workbook-sheet-tab, [class*='sheet-tab'], [class*='sheetTab']",
            "els => els.map(e => e.textContent.trim()).filter(t => t)",
        )
        if tabs:
            return tabs
    except Exception as e:
        logger.debug("获取 sheet tabs 失败: %s", e)

    return ["Sheet1"]  # fallback
