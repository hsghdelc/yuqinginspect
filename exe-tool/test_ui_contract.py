import ast
import inspect
import unicodedata
import unittest
from pathlib import Path

import main


SOURCE_PATH = Path(main.__file__)
SOURCE = SOURCE_PATH.read_text(encoding="utf-8")


class UiContractTests(unittest.TestCase):
    def test_main_module_parses(self):
        ast.parse(SOURCE)

    def test_no_legacy_form_controls_remain(self):
        forbidden = (
            "ttk.Button",
            "ttk.Entry",
            "ttk.Combobox",
            "ttk.Checkbutton",
            "ttk.LabelFrame",
            "ttk.Notebook",
        )
        for token in forbidden:
            self.assertNotIn(token, SOURCE)

    def test_no_emoji_symbols_in_program_text(self):
        symbols = [char for char in SOURCE if unicodedata.category(char) == "So"]
        self.assertEqual([], symbols)

    def test_file_actions_keep_their_callbacks(self):
        build_source = inspect.getsource(main.ReviewTool._build_ui)
        self.assertIn('self.choose_file()', build_source)
        self.assertIn('self.clear_file', build_source)
        self.assertIn('self.run', build_source)

    def test_drag_drop_and_worker_thread_are_preserved(self):
        drop_source = inspect.getsource(main.ReviewTool._enable_drop)
        run_source = inspect.getsource(main.ReviewTool.run)
        self.assertIn('self._on_drop', drop_source)
        self.assertIn('threading.Thread', run_source)
        self.assertIn('daemon=True', run_source)

    def test_required_visual_sections_exist(self):
        required = (
            "任务准备",
            "本次质检结果",
            "月度专项",
            "舆情提醒",
            "无效舆情",
            "超时预警",
            "报送内容",
            "处理日志",
            "质检报送",
            "校验提醒",
            "当前方案：",
        )
        for text in required:
            self.assertIn(text, SOURCE)
        self.assertNotIn("当前用户：", SOURCE)

    def test_settings_pages_share_modern_controls(self):
        for window_class in (
            main.BackupRestoreWindow,
            main.SchemeConfigWindow,
            main.SpecialConfigWindow,
        ):
            source = inspect.getsource(window_class._build_ui)
            self.assertIn("ctk.CTkFrame", source)
        self.assertIn("ctk.CTkTabview", inspect.getsource(main.SchemeConfigWindow._build_ui))


if __name__ == "__main__":
    unittest.main()
