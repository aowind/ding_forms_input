"""核心填入逻辑 - 模拟键盘操作填入钉钉表格"""

from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

# 默认参数
CELL_WAIT = 200            # 普通键盘操作等待 (ms)
NAV_WAIT = 450             # ArrowDown 导航等待 (ms)，防止被识别为长按
CELL_WAIT_SLOW = 500       # 回退重定位时用慢速
TYPE_DELAY = 30            # 输入字符间隔 (ms)
ROW_WAIT = 800             # 行间等待 (ms)


@dataclass
class FillResult:
    success: list[dict] = field(default_factory=list)
    failed: list[dict] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)


class TableFiller:
    """钉钉表格自动填入器。"""

    def __init__(self, page, log_callback=None):
        self.page = page
        self._log = log_callback or (lambda msg: logger.info(msg))
        self._current_row = 1
        self._current_col = 0  # 0=A
        self._abort = False

    def abort(self):
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
        self._current_col = 0
        self._log("已定位到 A1")

    async def navigate_to_row(self, target_row: int, match_col_offset: int) -> bool:
        """从当前位置相对移动到目标行目标列，并验证定位。

        只在定位失败时才回退到 A1 重来。

        Args:
            target_row: 目标行号 (1-based)
            match_col_offset: 匹配列距 A 列的偏移量 (A=0, D=3)

        Returns:
            True 定位成功
        """
        # 计算需要移动几行
        rows_to_move = target_row - self._current_row

        if rows_to_move > 0:
            # 向下移动 — 全部用 ArrowDown，绝对精确
            for _ in range(rows_to_move):
                if self._abort:
                    return False
                await self.page.keyboard.press("ArrowDown")
                await asyncio.sleep(NAV_WAIT / 1000)
        elif rows_to_move < 0:
            # 向上移动 — 回到 A1 重新向下导航
            self._log(f"  需要上移 {-rows_to_move} 行，回 A1 重定位...")
            await self._goto_a1()
            for _ in range(target_row - 1):
                if self._abort:
                    return False
                await self.page.keyboard.press("ArrowDown")
                await asyncio.sleep(NAV_WAIT / 1000)

        self._current_row = target_row

        # 先 Home 到 A 列
        await self.page.keyboard.press("Home")
        await asyncio.sleep(0.2)
        self._current_col = 0

        # ArrowRight 到匹配列
        for _ in range(match_col_offset):
            await self.page.keyboard.press("ArrowRight")
            await asyncio.sleep(CELL_WAIT / 1000)
        self._current_col = match_col_offset

        # Ctrl+C 验证（读到一个值即可，后续 run() 会精确比对）
        val = await self._read_cell()
        return len(val) > 0

    async def _goto_a1(self):
        """回到 A1（慢速，用于错误恢复）。"""
        await self.page.keyboard.down("Control")
        await self.page.keyboard.press("Home")
        await self.page.keyboard.up("Control")
        await asyncio.sleep(0.8)
        self._current_row = 1
        self._current_col = 0

    async def _verify_and_recover(self, target_row: int, match_col_offset: int, expected_value: str) -> bool:
        """验证当前单元格，如果不对就回 A1 重来一次。"""
        cell_val = await self._read_cell()
        if cell_val == expected_value:
            return True

        # 第一次失败，回 A1 精确重定位
        self._log(f"  ⚠️ 定位偏移（期望 {expected_value}，实际 {cell_val}），回 A1 重来...")
        await self._goto_a1()

        for _ in range(target_row - 1):
            if self._abort:
                return False
            await self.page.keyboard.press("ArrowDown")
            await asyncio.sleep(CELL_WAIT_SLOW / 1000)
        self._current_row = target_row

        await self.page.keyboard.press("Home")
        await asyncio.sleep(0.3)
        self._current_col = 0
        for _ in range(match_col_offset):
            await self.page.keyboard.press("ArrowRight")
            await asyncio.sleep(CELL_WAIT_SLOW / 1000)
        self._current_col = match_col_offset

        cell_val = await self._read_cell()
        if cell_val == expected_value:
            return True

        self._log(f"  ❌ 重定位仍失败（实际 {cell_val}），跳过")
        return False

    async def _read_cell(self) -> str:
        """Ctrl+C 读取当前单元格值。"""
        await self.page.keyboard.down("Control")
        await self.page.keyboard.press("c")
        await self.page.keyboard.up("Control")
        await asyncio.sleep(0.3)
        try:
            return (await self.page.evaluate("() => navigator.clipboard.readText()")).strip()
        except Exception:
            return ""

    async def fill_row(self, values: list[str]) -> int:
        """在当前行从填入列开始填入值。调用前光标在匹配列。"""
        # Tab 移到第一个填入列
        await self.page.keyboard.press("Tab")
        await asyncio.sleep(CELL_WAIT / 1000)

        filled = 0
        for i, val in enumerate(values):
            if self._abort:
                break

            if not val or val.strip() == "" or val == "None":
                await self.page.keyboard.press("Tab")
                await asyncio.sleep(CELL_WAIT / 1000)
                continue

            # F2 → Ctrl+A → type → Tab
            await self.page.keyboard.press("F2")
            await asyncio.sleep(0.15)

            await self.page.keyboard.down("Control")
            await self.page.keyboard.press("a")
            await self.page.keyboard.up("Control")
            await asyncio.sleep(0.08)

            await self.page.keyboard.type(val, delay=TYPE_DELAY)
            await asyncio.sleep(CELL_WAIT / 1000)

            await self.page.keyboard.press("Tab")
            await asyncio.sleep(CELL_WAIT / 1000)

            filled += 1
            if filled <= 5 or i == len(values) - 1:
                self._log(f"    填入: {val}")
            elif filled == 6:
                self._log("    ... (省略中间)")

        # Home 回到 A 列
        await self.page.keyboard.press("Home")
        await asyncio.sleep(0.2)
        self._current_col = 0

        return filled

    async def run(
        self,
        items: list[dict],
        id_mapping: dict[str, int],
        match_col_letter: str = "D",
        progress_callback=None,
        browser_manager=None,
    ) -> FillResult:
        """执行完整填入流程。"""
        result = FillResult()
        col_offset = ord(match_col_letter.upper()) - ord("A")

        # 匹配并排序
        matched: list[dict] = []
        for item in items:
            target = id_mapping.get(item["match_value"])
            if target:
                matched.append({**item, "target_row": target})
            else:
                result.skipped.append(item["match_value"])

        matched.sort(key=lambda x: x["target_row"])

        self._log(f"匹配成功: {len(matched)} 行, 未找到: {len(result.skipped)} 行")
        if result.skipped:
            self._log(f"未匹配 ID: {', '.join(result.skipped[:10])}" +
                      (f" 等 {len(result.skipped)} 个" if len(result.skipped) > 10 else ""))
        if not matched:
            return result

        # 将浏览器窗口移到屏幕外，防止用户干扰
        if browser_manager:
            await browser_manager.hide_window()

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
                    # navigate_to_row 失败，尝试恢复
                    ok = await self._verify_and_recover(tr, col_offset, mv)

                if not ok:
                    result.failed.append({"match_value": mv, "error": "定位失败"})
                    await self._goto_a1()
                    continue

                filled = await self.fill_row(vals)
                self._log(f"  ✅ 填入 {filled} 个值")
                result.success.append({
                    "match_value": mv,
                    "target_row": tr,
                    "filled": filled,
                })
            except Exception as e:
                self._log(f"  ❌ {e}")
                result.failed.append({"match_value": mv, "error": str(e)})
                await self._goto_a1()

            # 行间等待
            if i < len(matched) - 1 and not self._abort:
                # 预计下一行的移动行数
                if i + 1 < len(matched):
                    next_gap = matched[i + 1]["target_row"] - tr
                    wait = min(ROW_WAIT / 1000, max(0.3, next_gap * 0.01))
                else:
                    wait = ROW_WAIT / 1000
                await asyncio.sleep(wait)

            if progress_callback:
                progress_callback(idx, len(matched))

        self._log("")
        self._log(f"🎉 完成！成功: {len(result.success)}, 失败: {len(result.failed)}, 跳过: {len(result.skipped)}")

        # 恢复浏览器窗口位置
        if browser_manager:
            await browser_manager.show_window()

        return result
