# -*- coding: utf-8 -*-
"""Built-in unit tests for Word Formatter Pro."""

from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Pt

from wfp_config import DEFAULT_CONFIG
from wfp_core import (
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_PRESERVE,
    LegacyConversionUnavailable,
    SofficeConverter,
    WordProcessor,
)


class TextNormalizationTests(unittest.TestCase):
    def test_symbol_normalization_keeps_decimal_numbers(self):
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("你好,世界."),
            "你好，世界。",
        )
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("3.14 是 pi"),
            "3.14 是 pi",
        )

    def test_ellipsis_and_quotes(self):
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text('他说"你好"...'),
            "他说“你好”……",
        )
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("version 1.2"),
            "version 1.2",
        )

    def test_markdown_cleaning(self):
        raw = "# 标题\n**粗体** 和 [链接](https://example.com)\n![图片](a.png)\n> 引用"
        cleaned = WordProcessor._clean_markdown(raw)
        self.assertEqual(cleaned, "标题\n粗体 和 链接\n图片\n引用")

    def test_markdown_cleaning_preserves_source_numeric_numbering(self):
        raw = (
            "一、登录\n\n"
            "1. 打开软件\n"
            "2. 完成登录\n\n"
            "（一）首页搜索\n\n"
            "1. 输入关键词\n"
            "2. 点击院校\n"
            "1.2.3 小节号保持原样"
        )
        cleaned = WordProcessor._clean_markdown(raw)
        self.assertEqual(cleaned, raw)


class BlankLineTests(unittest.TestCase):
    def test_delete_single_blank_line_and_compress_multiple(self):
        text = "a\n\nb\n\n\nc"
        self.assertEqual(
            WordProcessor._remove_blank_lines_from_text(text),
            "a\nb\n\nc",
        )

    def test_keep_single_blank_line_and_compress_multiple(self):
        text = "a\n\nb\n\n\nc"
        self.assertEqual(
            WordProcessor._remove_blank_lines_from_text(text, keep_single_blank_lines=True),
            "a\n\nb\n\nc",
        )

    def test_blank_line_mode_aliases(self):
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("preserve"),
            BLANK_LINE_MODE_PRESERVE,
        )
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("delete_single"),
            BLANK_LINE_MODE_DELETE_SINGLE,
        )
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("keep_single"),
            BLANK_LINE_MODE_KEEP_SINGLE,
        )

    def test_processor_preserve_mode_leaves_text_unchanged(self):
        processor = WordProcessor({"blank_line_mode": BLANK_LINE_MODE_PRESERVE})
        text = "a\n\nb\n\n\nc"
        self.assertEqual(processor._normalize_text_blank_lines(text), text)


class TableHelperTests(unittest.TestCase):
    def test_numeric_table_text(self):
        self.assertTrue(WordProcessor._is_numeric_table_text("1,234.56"))
        self.assertTrue(WordProcessor._is_numeric_table_text("¥100元"))
        self.assertTrue(WordProcessor._is_numeric_table_text("12.5%"))
        self.assertFalse(WordProcessor._is_numeric_table_text("abc"))

    def test_short_table_text(self):
        self.assertTrue(WordProcessor._is_short_table_text("合计", max_len=4))
        self.assertFalse(WordProcessor._is_short_table_text("较长文本内容", max_len=4))

    def test_table_percentage_normalization(self):
        pcts = WordProcessor._normalize_table_pcts([1, 9], 20, 80)
        self.assertAlmostEqual(sum(pcts), 100.0)
        self.assertEqual(pcts, [20.0, 80.0])


class OoxmlProtectionTests(unittest.TestCase):
    def test_ooxml_element_detection(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run()
        run._r.append(OxmlElement("w:drawing"))
        self.assertTrue(WordProcessor._has_drawing_or_pict(para))
        self.assertFalse(WordProcessor._has_field_codes(para))

        para_field = doc.add_paragraph()
        field_run = para_field.add_run()
        field_run._r.append(OxmlElement("w:fldChar"))
        self.assertTrue(WordProcessor._has_field_codes(para_field))


class SpacingAndIndentUnitTests(unittest.TestCase):
    def test_default_pt_line_spacing_writes_exact_twips(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        processor = WordProcessor(DEFAULT_CONFIG.copy())

        processor._apply_line_spacing(
            para,
            'line_spacing',
            'line_spacing_unit',
            'line_spacing_multiple',
            28,
        )

        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(spacing.get(qn('w:line')), '560')
        self.assertEqual(spacing.get(qn('w:lineRule')), 'exact')

    def test_multiple_line_spacing_writes_auto_240_units(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        config = DEFAULT_CONFIG.copy()
        config.update({'line_spacing_unit': 'multiple', 'line_spacing_multiple': 1.5})
        processor = WordProcessor(config)

        processor._apply_line_spacing(
            para,
            'line_spacing',
            'line_spacing_unit',
            'line_spacing_multiple',
            28,
        )

        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(spacing.get(qn('w:line')), '360')
        self.assertEqual(spacing.get(qn('w:lineRule')), 'auto')

    def test_default_indent_keeps_centimeter_fields_and_two_char_first_line(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        processor = WordProcessor(DEFAULT_CONFIG.copy())

        processor._apply_text_indent_and_align(para)

        ind = para._p.get_or_add_pPr().get_or_add_ind()
        self.assertEqual(ind.get(qn('w:left')), '0')
        self.assertEqual(ind.get(qn('w:right')), '0')
        self.assertEqual(ind.get(qn('w:firstLineChars')), '200')
        self.assertIsNone(ind.get(qn('w:leftChars')))
        self.assertIsNone(ind.get(qn('w:rightChars')))

    def test_character_indent_writes_character_fields(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        config = DEFAULT_CONFIG.copy()
        config.update({
            'paragraph_indent_unit': 'chars',
            'left_indent_chars': 2,
            'right_indent_chars': 1.5,
            'first_line_indent_chars': 2,
        })
        processor = WordProcessor(config)

        processor._apply_text_indent_and_align(para)

        ind = para._p.get_or_add_pPr().get_or_add_ind()
        self.assertEqual(ind.get(qn('w:leftChars')), '200')
        self.assertEqual(ind.get(qn('w:rightChars')), '150')
        self.assertEqual(ind.get(qn('w:firstLineChars')), '200')
        self.assertIsNone(ind.get(qn('w:left')))
        self.assertIsNone(ind.get(qn('w:right')))

    def test_strip_leading_whitespace_removes_plain_blank_run(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("   ")
        para.add_run("正文")
        WordProcessor({})._strip_leading_whitespace(para)
        self.assertEqual(para.text, "正文")

    def test_strip_leading_whitespace_preserves_special_run(self):
        doc = Document()
        para = doc.add_paragraph()
        special_run = para.add_run()
        special_run._r.append(OxmlElement("w:fldChar"))
        para.add_run("正文")

        WordProcessor({})._strip_leading_whitespace(para)

        self.assertTrue(WordProcessor._has_field_codes(para))
        self.assertEqual(len(para.runs), 2)
        self.assertEqual(para.text, "正文")


class TempAndConversionTests(unittest.TestCase):
    def test_temp_docx_path_uses_system_temp_and_safe_name(self):
        processor = WordProcessor(DEFAULT_CONFIG.copy())
        temp_path = processor._make_temp_docx_path("copy", 'bad<name>:"?')
        try:
            self.assertEqual(
                os.path.normcase(os.path.abspath(os.path.dirname(temp_path))),
                os.path.normcase(os.path.abspath(tempfile.gettempdir())),
            )
            self.assertRegex(os.path.basename(temp_path), r"^~temp_copy_bad_name_[0-9]+_[0-9a-f]{8}\.docx$")
            self.assertIn(temp_path, processor.temp_files)
        finally:
            processor._cleanup_temp_files()

    def test_txt_conversion_uses_blank_line_mode_and_cleans_temp_file(self):
        with tempfile.TemporaryDirectory(prefix="wfp273_test_") as tmpdir:
            source = Path(tmpdir) / "sample.txt"
            source.write_text("标题\n\n正文一\n\n\n正文二", encoding="utf-8")

            processor = WordProcessor(DEFAULT_CONFIG.copy())
            temp_docx, is_from_txt = processor.convert_to_docx(str(source))
            try:
                self.assertTrue(is_from_txt)
                self.assertTrue(os.path.exists(temp_docx))
                self.assertEqual(
                    os.path.normcase(os.path.abspath(os.path.dirname(temp_docx))),
                    os.path.normcase(os.path.abspath(tempfile.gettempdir())),
                )
                self.assertFalse(list(Path(tmpdir).glob("~temp_*.docx")))

                converted = Document(temp_docx)
                self.assertEqual(
                    [p.text for p in converted.paragraphs],
                    ["标题", "正文一", "", "正文二"],
                )
            finally:
                processor._cleanup_temp_files()
            self.assertFalse(os.path.exists(temp_docx))

    def test_missing_soffice_marks_legacy_conversion_skipped(self):
        class MissingConverter:
            available = False

        processor = WordProcessor(DEFAULT_CONFIG.copy())
        processor.soffice_converter = MissingConverter()
        with self.assertRaises(LegacyConversionUnavailable):
            processor._convert_legacy_with_soffice("legacy.doc", "unused.docx")

    def test_format_document_skips_wps_when_soffice_missing(self):
        class MissingConverter:
            available = False

        with tempfile.TemporaryDirectory(prefix="wfp_legacy_skip_test_") as tmpdir:
            source = Path(tmpdir) / "legacy.wps"
            source.write_bytes(b"not a real wps file")

            processor = WordProcessor(DEFAULT_CONFIG.copy())
            processor.soffice_converter = MissingConverter()
            with patch.object(processor, '_com_available', return_value=False), self.assertRaises(LegacyConversionUnavailable):
                processor.format_document(str(source), str(Path(tmpdir) / "legacy_formatted.docx"))


class LegacyFormatTests(unittest.TestCase):
    def test_format_document_handles_doc_via_soffice_when_available(self):
        converter = SofficeConverter()
        if not converter.available:
            self.skipTest("LibreOffice soffice is not available")

        with tempfile.TemporaryDirectory(prefix="wfp_legacy_doc_test_") as tmpdir:
            root = Path(tmpdir)
            docx_source = root / "legacy_source.docx"
            source_doc = Document()
            source_doc.add_paragraph("旧格式标题")
            source_doc.add_paragraph("第一段正文")
            source_doc.save(docx_source)

            legacy_dir = root / "legacy"
            legacy_dir.mkdir()
            profile = root / "lo_profile"
            profile.mkdir()
            proc = subprocess.run(
                [
                    converter.soffice_path,
                    "--headless",
                    f"-env:UserInstallation={profile.as_uri()}",
                    "--convert-to",
                    "doc",
                    "--outdir",
                    str(legacy_dir),
                    str(docx_source),
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=120,
            )
            generated = sorted(legacy_dir.glob("*.doc"))
            if proc.returncode != 0 or not generated:
                detail = (proc.stderr or proc.stdout or "LibreOffice did not generate a .doc file").strip()
                self.skipTest(f"LibreOffice .doc export unavailable: {detail}")

            output = root / "legacy_formatted.docx"
            processor = WordProcessor(DEFAULT_CONFIG.copy(), soffice_path=converter.soffice_path)
            try:
                processor.format_document(str(generated[0]), str(output))
                self.assertTrue(output.exists())
                formatted = Document(output)
                self.assertTrue(any(paragraph.text.strip() for paragraph in formatted.paragraphs))
            finally:
                processor._cleanup_temp_files()


class EndToEndFormatTests(unittest.TestCase):
    def test_format_document_handles_direct_formats(self):
        with tempfile.TemporaryDirectory(prefix="wfp_format_test_") as tmpdir:
            root = Path(tmpdir)
            sources = []

            txt_source = root / "sample.txt"
            txt_source.write_text("测试标题\n\n第一段正文\n1. 小标题\n第二段正文", encoding="utf-8")
            sources.append(txt_source)

            md_source = root / "sample.md"
            md_source.write_text("# Markdown标题\n\n**第一段**\n\n1. 步骤一", encoding="utf-8")
            sources.append(md_source)

            docx_source = root / "sample.docx"
            source_doc = Document()
            source_doc.add_paragraph("Word标题")
            source_doc.add_paragraph("第一段正文")
            source_doc.save(docx_source)
            sources.append(docx_source)

            for source in sources:
                processor = WordProcessor(DEFAULT_CONFIG.copy())
                output = root / f"{source.stem}_formatted.docx"
                try:
                    processor.format_document(str(source), str(output))
                    self.assertTrue(output.exists(), f"missing output for {source.suffix}")
                    formatted = Document(output)
                    texts = [paragraph.text for paragraph in formatted.paragraphs]
                    self.assertTrue(any(text.strip() for text in texts), f"empty output for {source.suffix}")
                finally:
                    processor._cleanup_temp_files()


FLOATING_DRAWING_XML = (
    '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" behindDoc="0" '
    'locked="0" layoutInCell="1" allowOverlap="1">'
    '<wp:simplePos x="0" y="0"/>'
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    '<wp:extent cx="1905000" cy="1270000"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:wrapSquare wrapText="bothSides"/>'
    '<wp:docPr id="7" name="图片 7"/>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic>'
    '</wp:anchor></w:drawing>'
)

FLOATING_DRAWING_WITHOUT_ID_XML = (
    '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<wp:anchor distT="0"><wp:docPr name="图片 8"/>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic>'
    '</wp:anchor></w:drawing>'
)


def _add_drawing(paragraph, xml=FLOATING_DRAWING_XML):
    run = paragraph.add_run()
    run._r.append(parse_xml(xml))
    return run


class FloatingObjectTests(unittest.TestCase):
    def test_anchor_drawing_becomes_inline(self):
        doc = Document()
        para = doc.add_paragraph()
        drawing = _add_drawing(para)._r[-1]

        self.assertTrue(WordProcessor._anchor_to_inline(drawing, lambda: 99))

        self.assertEqual([WordProcessor._local_tag(child.tag) for child in drawing], ['inline'])
        inline = drawing[0]
        self.assertEqual(
            [WordProcessor._local_tag(child.tag) for child in inline],
            ['extent', 'effectExtent', 'docPr', 'graphic'],
        )
        self.assertEqual(inline.get('distL'), '0')
        self.assertEqual(inline.get('distR'), '0')
        self.assertEqual(inline.find(qn('wp:extent')).get('cx'), '1905000')
        self.assertEqual(inline.find(qn('wp:docPr')).get('id'), '7')
        self.assertIsNone(inline.find(qn('wp:positionH')))
        self.assertIsNone(inline.find(qn('wp:wrapSquare')))
        self.assertIsNone(inline.find(qn('wp:simplePos')))

    def test_missing_docpr_id_gets_unique_id(self):
        doc = Document()
        para = doc.add_paragraph()
        drawing = _add_drawing(para, FLOATING_DRAWING_WITHOUT_ID_XML)._r[-1]

        self.assertTrue(WordProcessor._anchor_to_inline(drawing, lambda: 99))

        inline = drawing[0]
        self.assertEqual(inline.find(qn('wp:docPr')).get('id'), '99')
        self.assertEqual(inline.find(qn('wp:extent')).get('cx'), '0')

    def test_inline_drawing_is_left_untouched(self):
        doc = Document()
        para = doc.add_paragraph()
        inline_drawing = OxmlElement('w:drawing')
        inline_drawing.append(OxmlElement('wp:inline'))
        para.add_run()._r.append(inline_drawing)

        self.assertFalse(WordProcessor._anchor_to_inline(inline_drawing, lambda: 1))
        self.assertEqual(WordProcessor._local_tag(inline_drawing[0].tag), 'inline')

    def test_vml_floating_shape_becomes_inline(self):
        shape = OxmlElement('w:pict')
        shape.set('style', 'width:200pt;height:150pt;position:absolute;left:0;z-index:2')

        self.assertTrue(WordProcessor._vml_shape_to_inline(shape))
        self.assertEqual(shape.get('style'), 'width:200pt;height:150pt')

        inline_shape = OxmlElement('w:pict')
        inline_shape.set('style', 'width:200pt;height:150pt')
        self.assertFalse(WordProcessor._vml_shape_to_inline(inline_shape))
        self.assertEqual(inline_shape.get('style'), 'width:200pt;height:150pt')

    def test_floating_table_becomes_inline(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        tbl_pr = table._tbl.find(qn('w:tblPr'))
        tbl_pr.append(OxmlElement('w:tblpPr'))
        tbl_pr.append(OxmlElement('w:tblOverlap'))
        processor = WordProcessor(DEFAULT_CONFIG.copy())

        self.assertGreaterEqual(processor._convert_floating_to_inline(doc), 1)
        tbl_pr = table._tbl.find(qn('w:tblPr'))
        self.assertIsNone(tbl_pr.find(qn('w:tblpPr')))
        self.assertIsNone(tbl_pr.find(qn('w:tblOverlap')))


class MediaLineSpacingTests(unittest.TestCase):
    def test_picture_paragraph_becomes_single_line_spacing(self):
        doc = Document()
        pic_para = doc.add_paragraph()
        _add_drawing(pic_para)
        body_para = doc.add_paragraph("正文内容")
        processor = WordProcessor(DEFAULT_CONFIG.copy())
        for para in (pic_para, body_para):
            processor._apply_line_spacing(
                para, 'line_spacing', 'line_spacing_unit', 'line_spacing_multiple', 28,
            )

        self.assertEqual(processor._apply_media_single_line_spacing(doc), 1)

        pic_spacing = pic_para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(pic_spacing.get(qn('w:line')), '240')
        self.assertEqual(pic_spacing.get(qn('w:lineRule')), 'auto')
        body_spacing = body_para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(body_spacing.get(qn('w:line')), '560')
        self.assertEqual(body_spacing.get(qn('w:lineRule')), 'exact')

    def test_picture_inside_table_cell_is_covered(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        para = table.cell(0, 0).paragraphs[0]
        _add_drawing(para)
        processor = WordProcessor(DEFAULT_CONFIG.copy())
        # 模拟表格排版把单元格写成了固定行距
        processor._apply_line_spacing(
            para, 'table_line_spacing', 'table_line_spacing_unit', 'table_line_spacing_multiple', 22,
        )

        self.assertEqual(processor._apply_media_single_line_spacing(doc), 1)

        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(spacing.get(qn('w:line')), '240')
        self.assertEqual(spacing.get(qn('w:lineRule')), 'auto')


class HeadingLevelPlanTests(unittest.TestCase):
    def test_numbered_heading_level_detection(self):
        self.assertEqual(WordProcessor._numbered_heading_level('一、总体要求'), 1)
        self.assertEqual(WordProcessor._numbered_heading_level('（一）工作要求'), 2)
        self.assertEqual(WordProcessor._numbered_heading_level('1. 方法'), 3)
        self.assertEqual(WordProcessor._numbered_heading_level('(1) 步骤'), 4)
        self.assertIsNone(WordProcessor._numbered_heading_level('普通正文'))

    def test_heading_text_heuristics(self):
        self.assertTrue(WordProcessor._looks_like_heading_text('工作要求'))
        self.assertFalse(WordProcessor._looks_like_heading_text('这是一段正文。'))
        self.assertFalse(WordProcessor._looks_like_heading_text('很长的一段正文内容' * 20))

    def test_format_signature_differs_by_font_and_bold(self):
        doc = Document()
        title = doc.add_paragraph("标题")
        title_run = title.runs[0]
        title_run.font.name = '黑体'
        title_run.font.size = Pt(16)
        title_run.font.bold = True
        body = doc.add_paragraph("正文")
        body_run = body.runs[0]
        body_run.font.name = '仿宋'
        body_run.font.size = Pt(16)

        signature = WordProcessor._paragraph_format_signature(title)
        self.assertIsNotNone(signature)
        self.assertNotEqual(signature, WordProcessor._paragraph_format_signature(body))
        self.assertEqual(signature, WordProcessor._paragraph_format_signature(title))
        self.assertIsNone(WordProcessor._paragraph_format_signature(doc.add_paragraph()))

    def test_same_format_paragraphs_join_anchor_level(self):
        records = [
            {'index': 0, 'signature': 'same', 'level': 2, 'evidence': 2, 'joinable': True},
            {'index': 1, 'signature': 'same', 'level': None, 'evidence': 0, 'joinable': True},
            {'index': 2, 'signature': 'other', 'level': None, 'evidence': 0, 'joinable': True},
            {'index': 3, 'signature': 'same', 'level': None, 'evidence': 0, 'joinable': False},
        ]
        self.assertEqual(WordProcessor._plan_heading_levels(records), {1: 2})

    def test_numbered_paragraph_keeps_its_own_level(self):
        records = [
            {'index': 0, 'signature': 'same', 'level': 1, 'evidence': 2,
             'numbered_level': 1, 'joinable': True},
            {'index': 1, 'signature': 'same', 'level': 3, 'evidence': 1,
             'numbered_level': 3, 'joinable': True},
            {'index': 2, 'signature': 'same', 'level': None, 'evidence': 0,
             'numbered_level': None, 'joinable': True},
        ]
        plan = WordProcessor._plan_heading_levels(records)
        self.assertEqual(plan, {2: 1})

    def test_heading_style_sample_wins_over_unstyled_sample(self):
        records = [
            {'index': 0, 'signature': 'same', 'level': 3, 'evidence': 2, 'joinable': True},
            {'index': 1, 'signature': 'same', 'level': 2, 'evidence': 3, 'joinable': True},
        ]
        self.assertEqual(WordProcessor._plan_heading_levels(records), {0: 2})

    def test_numbered_level_is_kept_when_group_level_differs(self):
        records = [
            {'index': 0, 'signature': 'same', 'level': 1, 'evidence': 2,
             'numbered_level': 1, 'joinable': True},
            {'index': 1, 'signature': 'same', 'level': 3, 'evidence': 1,
             'numbered_level': 3, 'joinable': True},
        ]
        plan = WordProcessor._plan_heading_levels(records)
        self.assertNotIn(1, plan, "自身带编号且与合并等级不同时应保留编号等级")

    def test_majority_level_wins_within_same_evidence(self):
        records = [
            {'index': index, 'signature': 'same', 'level': level, 'evidence': 2, 'joinable': True}
            for index, level in enumerate([3, 3, 2])
        ]
        self.assertEqual(WordProcessor._plan_heading_levels(records), {2: 3})

    def test_group_without_level_anchor_is_ignored(self):
        records = [
            {'index': 0, 'signature': 'same', 'level': None, 'evidence': 0, 'joinable': True},
            {'index': 1, 'signature': None, 'level': None, 'evidence': 0, 'joinable': True},
        ]
        self.assertEqual(WordProcessor._plan_heading_levels(records), {})

    def test_level_compaction_only_fills_gaps(self):
        self.assertEqual(WordProcessor._plan_level_compaction([1, 3]), {3: 2})
        self.assertEqual(WordProcessor._plan_level_compaction([1, 2, 3]), {})
        self.assertEqual(WordProcessor._plan_level_compaction([2, 3, 5]), {5: 4})
        self.assertEqual(WordProcessor._plan_level_compaction([4]), {})
        self.assertEqual(WordProcessor._plan_level_compaction([1, 3, 4, 6]), {3: 2, 4: 3, 6: 4})
        self.assertEqual(WordProcessor._plan_level_compaction([]), {})


class OptionalEnhancementIntegrationTests(unittest.TestCase):
    def test_optional_enhancements_apply_to_docx(self):
        with tempfile.TemporaryDirectory(prefix="wfp_enhance_test_") as tmpdir:
            root = Path(tmpdir)
            source = root / "sample.docx"
            doc = Document()
            doc.add_paragraph("一、总体要求")
            doc.add_paragraph("这是第一段正文内容。")
            pic_para = doc.add_paragraph()
            _add_drawing(pic_para)
            doc.save(source)

            config = DEFAULT_CONFIG.copy()
            config.update({
                'convert_floating_objects': True,
                'media_single_line_spacing': True,
                'normalize_heading_levels': True,
            })
            processor = WordProcessor(config)
            output = root / "sample_formatted.docx"
            try:
                processor.format_document(str(source), str(output))
            finally:
                processor._cleanup_temp_files()

            formatted = Document(output)
            drawings = list(formatted.element.body.iter(qn('w:drawing')))
            self.assertTrue(drawings, "浮动图片应保留在输出文档中")
            for drawing in drawings:
                self.assertEqual(
                    [WordProcessor._local_tag(child.tag) for child in drawing], ['inline'],
                )

            pic_paragraphs = [
                para for para in formatted.paragraphs
                if WordProcessor._has_drawing_or_pict(para)
            ]
            self.assertTrue(pic_paragraphs)
            for para in pic_paragraphs:
                spacing = para._p.get_or_add_pPr().get_or_add_spacing()
                self.assertEqual(spacing.get(qn('w:line')), '240')
                self.assertEqual(spacing.get(qn('w:lineRule')), 'auto')

            heading = next(
                para for para in formatted.paragraphs if para.text.strip().startswith('一、')
            )
            # w:outlineLvl 以 0 表示一级，日志与显示时 +1
            self.assertEqual(processor._get_outline_level(heading), 0)


def main(argv=None):
    suite = unittest.defaultTestLoader.loadTestsFromModule(__import__(__name__))
    import wfp_numbering_tests
    import wfp_numbering_integration_tests
    for module in (wfp_numbering_tests, wfp_numbering_integration_tests):
        suite.addTests(unittest.defaultTestLoader.loadTestsFromModule(module))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    if result.wasSuccessful():
        print("所有单元测试通过")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
