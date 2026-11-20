# -*- coding: utf-8 -*-
"""图纸识读解析的单元测试：python -m unittest discover -s tests"""

import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))

import ai_model
import ai_vision


class ParseRowsTest(unittest.TestCase):
    def test_plain_json(self):
        rows = ai_vision.parse_rows(
            '[{"station": "K0+000", "well": "Y1", "ground": "512.345", "dia": "DN400"}]')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["station"], "K0+000")
        self.assertEqual(rows[0]["dia"], "DN400")

    def test_fenced_json(self):
        text = '```json\n[{"well": "Y2", "len_c": "12.50"}]\n```'
        rows = ai_vision.parse_rows(text)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["well"], "Y2")

    def test_noise_around_json(self):
        rows = ai_vision.parse_rows('说明：结果如下 [{"well": "Y3"}] 以上')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["well"], "Y3")

    def test_dict_wrapper(self):
        rows = ai_vision.parse_rows('{"rows": [{"well": "Y4"}]}')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["well"], "Y4")

    def test_empty_and_bad(self):
        self.assertEqual(ai_vision.parse_rows(""), [])
        self.assertEqual(ai_vision.parse_rows("not json"), [])
        self.assertEqual(ai_vision.parse_rows("[{}]"), [])

    def test_normalize_units(self):
        row = ai_vision.normalize_row({"len_c": "12.50 m", "ground": "512.345m",
                                       "pipe_bottom": "510.200（设计）", "dia": "d600"})
        self.assertEqual(row["len_c"], "12.50")
        self.assertEqual(row["ground"], "512.345")
        self.assertEqual(row["pipe_bottom"], "510.200")
        self.assertEqual(row["dia"], "d600")


class VisionModelTest(unittest.TestCase):
    def test_flags(self):
        self.assertTrue(ai_model.is_vision_model("qwen-vl-max"))
        self.assertTrue(ai_model.is_vision_model("glm-4v-plus"))
        self.assertTrue(ai_model.is_vision_model("gpt-4o"))
        self.assertTrue(ai_model.is_vision_model("gemini-2.0-flash"))
        self.assertFalse(ai_model.is_vision_model("deepseek-chat"))
        self.assertFalse(ai_model.is_vision_model("qwen-plus"))

    def test_default_prefers_vision(self):
        self.assertEqual(ai_model.default_profile("通义千问")["model"], "qwen3-vl-plus")
        self.assertEqual(ai_model.default_profile("DeepSeek")["model"], "deepseek-chat")
        # 硅基流动的视觉候选必须是当前可用的模型
        self.assertTrue(ai_model.is_vision_model(
            ai_model.default_profile("硅基流动")["model"]))
        self.assertNotIn("deepseek-vl2",
                         " ".join(ai_model.model_choices("硅基流动")))


class ImageEncodeTest(unittest.TestCase):
    def test_large_image_auto_shrinks(self):
        try:
            from PIL import Image
        except Exception:
            self.skipTest("Pillow not installed")
        import shutil
        import tempfile

        workdir = tempfile.mkdtemp(prefix="ai_img_test_")
        try:
            path = os.path.join(workdir, "big.png")
            Image.new("RGB", (4200, 3000), "white").save(path)
            url = ai_model.image_data_url(path, max_side=2200)
            self.assertTrue(url.startswith("data:image/jpeg;base64,"))
            self.assertLess(len(url), int(3.6 * 1024 * 1024 * 1.4))
            self.assertGreater(len(url), 1000)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)


class DrawingPipelineTest(unittest.TestCase):
    def test_reject_non_vision_model(self):
        profile = {"model": "deepseek-chat", "endpoint": "https://x/v1", "key": "", "timeout": 5}
        with self.assertRaises(RuntimeError):
            ai_vision.ocr_drawing(profile, __file__)

    def test_pdf_pages_render(self):
        try:
            import fitz
        except Exception:
            self.skipTest("PyMuPDF not installed")
        import shutil
        import tempfile

        workdir = tempfile.mkdtemp(prefix="ai_vision_test_")
        try:
            pdf = os.path.join(workdir, "drawing.pdf")
            doc = fitz.open()
            page = doc.new_page(width=842, height=595)
            page.insert_text((40, 60), "STATION  WELL  GROUND  PIPE BOTTOM", fontsize=10)
            doc.save(pdf)
            doc.close()

            images = ai_vision.page_images(pdf, workdir)
            self.assertEqual(len(images), 1)
            self.assertTrue(os.path.exists(images[0]))
            self.assertGreater(os.path.getsize(images[0]), 1000)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
