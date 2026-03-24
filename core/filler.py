"""核心填入逻辑 - 模拟键盘操作填入钉钉表格"""

from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

# 默认参数
CELL_WAIT = 350          # 每次键盘操作等待 (ms)
TYPE_DELAY = 30           # 输入字符间隔 (ms)
ROW_WAIT = 1500           # 行间等待 (ms)


@dataclass
class FillResult:
    success: list[dict] = field(default_factory=list)
    failed: list[dict] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)


class TableFiller:
    """钉钉表格自动填入器。"""

    def __init__(self, page, log_callback=None):
        """
        Args:
            page: Playwright Page 对象
            log_callback: 日志回调函数 log_callback(str)
        """
        self.page = page
        self._log = log_callback or (lambda msg: logger.info(msg))
        self._current_row = 1
        self._abort = False

    def abort(self):
        """中断执行。"""
        self._abort = True

    async def init_position(self):
        """初始化：点击表格区域聚焦，Ctrl+Home 回到 A1。"""
        await self.page.mouse.click(500, 300)
        await asyncio.sleep(1.0)
        await self.page.keyboard.down("Control")
        await self.page.keyboard.press("Home")
        await self.page.keyboard.up("Control")
        await asyncio.sleep(1.0)
        self._current_row = 1
        self._log("已定位到 A1")

    async def navigate_to_row(self, target_row: int, match_col_offset: int) -> bool:
        """从当前行导航到目标行目标列，并验证定位。

        Args:
            target_row: 目标行号 (1-based, Excel 行号)
            match_col_offset: 匹配列距离 A 列的偏移量 (A=0, D=3)

        Returns:
            True 定位成功
        """
        # 回到 A1 精确导航
        await self.page.keyboard.down("Control")
        await self.page.keyboard.press("Home")
        await self.page.keyboard.up("Control")
        await asyncio.sleep(0.8)
        self._current_row = 1

        # 逐行 ArrowDown
        for _ in range(target_row - 1):
            if self._abort:
                return False
            await self.page.keyboard.press("ArrowDown")
            await asyncio.sleep(CELL_WAIT / 1000)
        self._current_row = target_row

        # Home 到 A 列
        await self.page.keyboard.press("Home")
        await asyncio.sleep(0.3)

        # ArrowRight 到匹配列
        for _ in range(match_col_offset):
            await self.page.keyboard.press("ArrowRight")
            await asyncio.sleep(CELL_WAIT / 1000)

        # Ctrl+C 验证
        return await self._verify_cell()

    async def _verify_cell(self) -> bool:
        """读取当前单元格值用于验证。"""
        await self.page.keyboard.down("Control")
        await self.page.keyboard.press("c")
        await self.page.keyboard.up("Control")
        await asyncio.sleep(0.3)
        try:
            val = await self.page.evaluate("() => navigator.clipboard.readText()")
            return True  # 读到值说明定位正确
        except Exception:
            return False

    async def fill_row(self, values: list[str], fill_col_offset: int) -> int:
        """在当前行从填入列开始填入值。

        Args:
            values: 待填入的值列表
            fill_col_offset: 第一个填入列距离匹配列的偏移 (通常是 Tab 1 次到 E)

        Returns:
            实际填入的非空值数量
        """
        # Tab 移到第一个填入列
        await self.page.keyboard.press("Tab")
        await asyncio.sleep(CELL_WAIT / 1000)

        filled = 0
        for i, val in enumerate(values):
            if self._abort:
                break

            if not val or val.strip() == "" or val == "None":
                # 空值也要 Tab 跳过，保持列对齐
                await self.page.keyboard.press("Tab")
                await asyncio.sleep(CELL_WAIT / 1000)
                continue

            # F2 进入编辑
            await self.page.keyboard.press("F2")
            await asyncio.sleep(0.2)

            # Ctrl+A 全选
            await self.page.keyboard.down("Control")
            await self.page.keyboard.press("a")
            await self.page.keyboard.up("Control")
            await asyncio.sleep(0.1)

            # 输入值
            await self.page.keyboard.type(val, delay=TYPE_DELAY)
            await asyncio.sleep(CELL_WAIT / 1000)

            # Tab 确认并移到下一列（不能用 Enter！）
            await self.page.keyboard.press("Tab")
            await asyncio.sleep(CELL_WAIT / 1000)

            filled += 1
            if filled <= 5 or i == len(values) - 1:
                self._log(f"    填入: {val}")
            elif filled == 6:
                self._log("    ... (省略中间)")

        # Home 回到 A 列
        await self.page.keyboard.press("Home")
        await asyncio.sleep(0.3)

        return filled

    async def run(
        self,
        items: list[dict],
        id_mapping: dict[str, int],
        match_col_letter: str = "D",
        progress_callback=None,
    ) -> FillResult:
        """执行完整填入流程。

        Args:
            items: 待填入数据列表，每项包含 match_value 和 fill_values
            id_mapping: {匹配值: 钉钉表格行号}
            match_col_letter: 匹配列字母 (A=0, B=1, ..., D=3)
            progress_callback: 进度回调 progress_callback(current, total)

        Returns:
            FillResult 填入结果
        """
        result = FillResult()
        col_offset = ord(match_col_letter.upper()) - ord("A")
        total = len(items)

        # 匹配
        matched: list[dict] = []
        for item in items:
            target = id_mapping.get(item["match_value"])
            if target:
                matched.append({**item, "target_row": target})
            else:
                result.skipped.append(item["match_value"])

        # 按目标行号排序
        matched.sort(key=lambda x: x["target_row"])

        self._log(f"匹配成功: {len(matched)} 行, 未找到: {len(result.skipped)} 行")
        if not matched:
            return result

        # 初始化位置
        await self.init_position()

        for i, item in enumerate(matched):
            if self._abort:
                self._log("用户中断执行")
                break

            mv = item["match_value"]
            tr = item["target_row"]
            vals = item["fill_values"]
            idx = i + 1

            self._log(f"[{idx}/{len(matched)}] {mv} → 第 {tr} 行")

            try:
                ok = await self.navigate_to_row(tr, col_offset)
                if not ok:
                    self._log(f"  ❌ 定位失败，跳过")
                    result.failed.append({"match_value": mv, "error": "定位失败"})
                    # 重新初始化
                    await self.init_position()
                    continue

                filled = await self.fill_row(vals, 1)
                self._log(f"  ✅ 填入 {filled} 个值")
                result.success.append({
                    "match_value": mv,
                    "target_row": tr,
                    "filled": filled,
                })
            except Exception as e:
                self._log(f"  ❌ {e}")
                result.failed.append({"match_value": mv, "error": str(e)})
                await self.init_position()

            # 行间等待
            if i < len(matched) - 1 and not self._abort:
                self._log(f"  等待 {ROW_WAIT / 1000:.1f}s...")
                await asyncio.sleep(ROW_WAIT / 1000)

            if progress_callback:
                progress_callback(idx, len(matched))

        self._log("")
        self._log(f"🎉 完成！成功: {len(result.success)}, 失败: {len(result.failed)}, 跳过: {len(result.skipped)}")
        return result
