"""Chinese desktop client for the local resume diagnosis prototype."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QGuiApplication
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
    QPushButton, QScrollArea, QSizePolicy, QStackedWidget, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

from reporting import build_report
from resume_engine import (
    JobProfile, analyze_resume, job_from_description, load_demo_jobs,
    load_jobs_csv, read_document, to_json_ready,
)


ROOT = Path(__file__).resolve().parent
STYLE = """
QWidget { font-family: "Microsoft YaHei UI", "Segoe UI"; font-size: 13px; color: #1d1d1f; }
QWidget#window { background: #f5f5f7; }
QWidget#mainArea, QWidget#content { background: transparent; }
QFrame#sidebar { background: #f0f0f2; border-right: 1px solid #dedee3; }
QFrame#card { background: white; border: 1px solid #e7e7eb; border-radius: 18px; }
QFrame#metric { background: white; border: 1px solid #e7e7eb; border-radius: 16px; }
QLabel#appName { font-size: 18px; font-weight: 700; color: #17171a; }
QLabel#pageTitle { font-size: 27px; font-weight: 700; color: #1d1d1f; }
QLabel#sectionTitle { font-size: 16px; font-weight: 700; }
QLabel#muted { color: #73737b; }
QLabel#small { color: #777780; font-size: 11px; }
QLabel#metricValue { font-size: 27px; font-weight: 700; color: #007aff; }
QLabel#good { color: #168449; }
QLabel#attention { color: #bd6a00; }
QPushButton { background: #ffffff; border: 1px solid #dedee4; border-radius: 10px;
              padding: 9px 16px; color: #1d1d1f; font-weight: 600; }
QPushButton:hover { background: #f7f7fa; border-color: #b9b9c1; }
QPushButton:pressed { background: #ededf2; }
QPushButton#primary { background: #007aff; border-color: #007aff; color: white; }
QPushButton#primary:hover { background: #006bea; border-color: #006bea; }
QPushButton#nav { text-align: left; background: transparent; border: 0; border-radius: 10px;
                  padding: 13px 15px; color: #4d4d55; }
QPushButton#nav:hover { background: #e5e5ea; }
QPushButton#nav[active="true"] { background: #dcecff; color: #0066cc; font-weight: 700; }
QLineEdit, QPlainTextEdit, QComboBox { background: white; border: 1px solid #d9d9df;
                                      border-radius: 10px; padding: 9px; selection-background-color: #b9dcff; }
QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus { border: 1px solid #007aff; }
QPlainTextEdit { line-height: 1.5; }
QTableWidget { background: white; border: 1px solid #e7e7eb; border-radius: 10px;
               alternate-background-color: #fafafc; gridline-color: #ececf0; }
QTableWidget::item { padding: 6px; }
QTableWidget::item:selected { background: #dcecff; color: #1d1d1f; }
QHeaderView::section { background: #f6f6f8; color: #56565e; font-weight: 600;
                       border: 0; border-bottom: 1px solid #e7e7eb; padding: 10px; }
QCheckBox { spacing: 8px; }
QScrollArea { border: 0; background: transparent; }
"""


def label(text: str, kind: str = "", wrap: bool = False) -> QLabel:
    item = QLabel(text)
    if kind:
        item.setObjectName(kind)
    item.setWordWrap(wrap)
    return item


def card() -> tuple[QFrame, QVBoxLayout]:
    frame = QFrame()
    frame.setObjectName("card")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(22, 20, 22, 20)
    layout.setSpacing(12)
    return frame, layout


def scroll_page() -> tuple[QScrollArea, QVBoxLayout]:
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setMinimumSize(0, 0)
    scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Ignored)
    content = QWidget()
    content.setObjectName("content")
    layout = QVBoxLayout(content)
    layout.setContentsMargins(4, 4, 4, 20)
    layout.setSpacing(16)
    layout.setAlignment(Qt.AlignTop)
    scroll.setWidget(content)
    return scroll, layout


def read_only_box(text: str = "", height: int = 160) -> QPlainTextEdit:
    box = QPlainTextEdit(text)
    box.setReadOnly(True)
    box.setMinimumHeight(height)
    return box


class ResumeDesktopApp(QWidget):
    PAGE_NAMES = ("简历上传", "岗位要求", "匹配分析", "修改建议", "报告导出")

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("window")
        self.setWindowTitle("机械类学生简历诊断与岗位匹配系统")
        self.resize(1320, 850)
        self.setMinimumSize(1050, 700)
        self.demo_jobs = load_demo_jobs(ROOT / "data" / "demo_jobs.json")
        self.imported_jobs: list[JobProfile] = []
        self.manual_jobs: list[JobProfile] = []
        self.jobs: list[JobProfile] = []
        self.result: dict | None = None
        self.report_bytes: bytes | None = None
        self._build_shell()
        self.resume_text.textChanged.connect(self._invalidate_analysis)
        self.student_input.textChanged.connect(self._invalidate_analysis)
        self.jobs_table.itemSelectionChanged.connect(self._invalidate_analysis)
        self.refresh_jobs()

    def _invalidate_analysis(self) -> None:
        if self.result is not None:
            self.result = None
            self.report_bytes = None
            for value in (self.score_value, self.strength_value, self.gap_value):
                value.setText("—")
            self.target_label.setText("输入已变化，请重新分析。")
            self.dim_table.setRowCount(0)
            self.rank_table.setRowCount(0)
            self.strength_text.setPlainText("重新分析后显示。")
            self.weakness_text.setPlainText("重新分析后显示。")
            self.evidence_text.setPlainText("重新分析后显示。")
            self.report_preview.setPlainText("输入已变化，请重新生成报告。")
            self._clear_suggestions()
            self.status.setText("简历、称呼或目标岗位已变化，请重新点击“开始分析并生成报告”。")

    def _build_shell(self) -> None:
        whole = QHBoxLayout(self)
        whole.setContentsMargins(0, 0, 0, 0)
        whole.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(222)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(16, 30, 16, 22)
        side.setSpacing(6)
        side.addWidget(label("简历诊断", "appName"))
        side.addWidget(label("就业能力分析工作台", "small"))
        side.addSpacing(30)
        self.nav_buttons: list[QPushButton] = []
        for index, title in enumerate(self.PAGE_NAMES):
            button = QPushButton(f"{index + 1:02d}   {title}")
            button.setObjectName("nav")
            button.setCursor(Qt.PointingHandCursor)
            button.clicked.connect(lambda checked=False, i=index: self.show_page(i))
            self.nav_buttons.append(button)
            side.addWidget(button)
        side.addStretch()
        side.addWidget(label("本地运行 · 简历不主动上传", "small", True))
        whole.addWidget(sidebar)

        main = QWidget()
        main.setObjectName("mainArea")
        column = QVBoxLayout(main)
        column.setContentsMargins(32, 25, 32, 15)
        column.setSpacing(18)
        top = QHBoxLayout()
        title_stack = QVBoxLayout()
        self.page_title = label("简历上传", "pageTitle")
        self.page_subtitle = label("添加学生简历，核对系统提取的文字。", "muted")
        title_stack.addWidget(self.page_title)
        title_stack.addWidget(self.page_subtitle)
        top.addLayout(title_stack, 1)
        self.analyze_button = QPushButton("开始分析并生成报告")
        self.analyze_button.setObjectName("primary")
        self.analyze_button.setCursor(Qt.PointingHandCursor)
        self.analyze_button.clicked.connect(self.analyze)
        top.addWidget(self.analyze_button, 0, Qt.AlignTop)
        column.addLayout(top)
        self.pages = QStackedWidget()
        self.pages.setMinimumSize(0, 0)
        self.pages.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Ignored)
        self.pages.addWidget(self._resume_page())
        self.pages.addWidget(self._jobs_page())
        self.pages.addWidget(self._analysis_page())
        self.pages.addWidget(self._suggestions_page())
        self.pages.addWidget(self._report_page())
        column.addWidget(self.pages, 1)
        self.status = label("准备就绪。先添加简历，再选择目标岗位。", "small")
        column.addWidget(self.status)
        whole.addWidget(main, 1)
        self.show_page(0)

    def show_page(self, index: int) -> None:
        self.pages.setCurrentIndex(index)
        self.page_title.setText(self.PAGE_NAMES[index])
        descriptions = (
            "添加学生简历，核对系统提取的文字。",
            "选择目标岗位，或导入真实企业的岗位要求。",
            "查看分维度指标、证据原句和岗位排序。",
            "逐句比较原文、问题和建议替换句。",
            "查看报告摘要并保存可编辑的 Word 报告。",
        )
        self.page_subtitle.setText(descriptions[index])
        for number, button in enumerate(self.nav_buttons):
            button.setProperty("active", number == index)
            button.style().unpolish(button)
            button.style().polish(button)

    def _resume_page(self) -> QScrollArea:
        page, layout = scroll_page()
        frame, body = card()
        body.addWidget(label("学生简历", "sectionTitle"))
        row = QHBoxLayout()
        upload = QPushButton("选择 PDF / DOCX / TXT 文件")
        upload.clicked.connect(self.choose_resume)
        sample = QPushButton("载入虚构演示简历")
        sample.clicked.connect(self.load_sample_resume)
        row.addWidget(upload)
        row.addWidget(sample)
        row.addStretch()
        body.addLayout(row)
        self.resume_path = label("尚未选择文件，也可以直接在下方粘贴简历。", "muted", True)
        body.addWidget(self.resume_path)
        name_row = QHBoxLayout()
        name_row.addWidget(label("报告中的学生称呼"))
        self.student_input = QLineEdit("学生（匿名）")
        self.student_input.setMaximumWidth(290)
        name_row.addWidget(self.student_input)
        name_row.addStretch()
        body.addLayout(name_row)
        layout.addWidget(frame)

        frame, body = card()
        body.addWidget(label("简历文本", "sectionTitle"))
        body.addWidget(label("上传后可检查并直接修改文字；点击右上角按钮开始分析。", "muted"))
        self.resume_text = QPlainTextEdit()
        self.resume_text.setPlaceholderText("在这里粘贴简历内容，或点击上方按钮上传文件……")
        self.resume_text.setMinimumHeight(380)
        body.addWidget(self.resume_text)
        body.addWidget(label("扫描版 PDF 需先 OCR。测试时建议使用匿名编号，不填写身份证号等非必要信息。", "small", True))
        layout.addWidget(frame)
        return page

    def _jobs_page(self) -> QScrollArea:
        page, layout = scroll_page()
        frame, body = card()
        row = QHBoxLayout()
        row.addWidget(label("目标岗位", "sectionTitle"))
        row.addStretch()
        self.use_demo = QCheckBox("显示 5 个演示岗位")
        self.use_demo.setChecked(True)
        self.use_demo.toggled.connect(self.refresh_jobs)
        row.addWidget(self.use_demo)
        body.addLayout(row)
        body.addWidget(label("点击表格中的一行作为重点分析岗位。演示岗位仅用于试用，不是真实招聘。", "muted", True))
        button_row = QHBoxLayout()
        for text, callback in (("导入企业岗位 CSV", self.import_csv), ("清空导入岗位", self.clear_imported),
                               ("清空手动岗位", self.clear_manual)):
            button = QPushButton(text)
            button.clicked.connect(callback)
            button_row.addWidget(button)
        button_row.addStretch()
        body.addLayout(button_row)
        self.jobs_table = self._table(["岗位名称", "企业", "要求数", "来源"], 220)
        self.jobs_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.jobs_table.setSelectionMode(QTableWidget.SingleSelection)
        body.addWidget(self.jobs_table)
        layout.addWidget(frame)

        frame, body = card()
        body.addWidget(label("手动添加一个岗位", "sectionTitle"))
        body.addWidget(label("粘贴岗位说明，也可以填写具体技能、工具、职责和项目关键词。", "muted"))
        grid = QGridLayout()
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(12)
        self.title_input = QLineEdit()
        self.employer_input = QLineEdit()
        self.degree_input = QComboBox()
        self.degree_input.addItems(["", "中专", "大专", "本科", "硕士", "博士"])
        self.majors_input = QLineEdit()
        self.skills_input = QLineEdit()
        self.tools_input = QLineEdit()
        self.duties_input = QLineEdit()
        self.projects_input = QLineEdit()
        fields = (
            ("岗位名称 *", self.title_input), ("企业名称", self.employer_input),
            ("最低学历", self.degree_input), ("相关专业", self.majors_input),
            ("专业技能", self.skills_input), ("工具软件", self.tools_input),
            ("工作职责关键词", self.duties_input), ("项目经验关键词", self.projects_input),
        )
        for index, (title, widget) in enumerate(fields):
            row, col = divmod(index, 2)
            cell = QVBoxLayout()
            cell.addWidget(label(title, "small"))
            cell.addWidget(widget)
            grid.addLayout(cell, row, col)
        body.addLayout(grid)
        body.addWidget(label("岗位说明原文", "small"))
        self.jd_text = QPlainTextEdit()
        self.jd_text.setPlaceholderText("在这里粘贴企业岗位说明……")
        self.jd_text.setMinimumHeight(110)
        body.addWidget(self.jd_text)
        bottom = QHBoxLayout()
        bottom.addWidget(label("多个要求可用顿号、逗号或分号分隔。", "small"))
        bottom.addStretch()
        add_button = QPushButton("添加到岗位列表")
        add_button.setObjectName("primary")
        add_button.clicked.connect(self.add_manual_job)
        bottom.addWidget(add_button)
        body.addLayout(bottom)
        layout.addWidget(frame)
        return page

    @staticmethod
    def _table(headers: list[str], height: int) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.verticalHeader().hide()
        table.horizontalHeader().setStretchLastSection(True)
        table.setMinimumHeight(height)
        return table

    @staticmethod
    def _fill_table(table: QTableWidget, rows: list[list[str]]) -> None:
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for column, value in enumerate(row):
                table.setItem(row_index, column, QTableWidgetItem(str(value)))
        table.resizeColumnsToContents()
        for row_index in range(len(rows)):
            table.setRowHeight(row_index, 43)

    def _analysis_page(self) -> QScrollArea:
        page, layout = scroll_page()
        metrics = QHBoxLayout()
        self.score_value = label("—", "metricValue")
        self.strength_value = label("—", "metricValue")
        self.gap_value = label("—", "metricValue")
        for caption, value in (("证据匹配度", self.score_value), ("优势线索", self.strength_value),
                               ("待改进线索", self.gap_value)):
            frame = QFrame()
            frame.setObjectName("metric")
            box = QVBoxLayout(frame)
            box.setContentsMargins(20, 16, 20, 16)
            box.addWidget(label(caption, "muted"))
            box.addWidget(value)
            metrics.addWidget(frame, 1)
        layout.addLayout(metrics)
        self.target_label = label("请先上传简历、选择目标岗位并点击“开始分析并生成报告”。", "muted", True)
        layout.addWidget(self.target_label)
        frame, body = card()
        body.addWidget(label("分维度指标", "sectionTitle"))
        body.addWidget(label("覆盖率反映简历已展示的岗位要求比例；加权得分由岗位适用维度的权重计算。", "muted", True))
        self.dim_table = self._table(["维度", "覆盖率", "权重", "加权得分", "判断", "未找到的要求"], 280)
        body.addWidget(self.dim_table)
        layout.addWidget(frame)
        pair = QHBoxLayout()
        for title, attr in (("已展示的优势", "strength_text"), ("待改进与证据缺口", "weakness_text")):
            frame, body = card()
            body.addWidget(label(title, "sectionTitle"))
            box = read_only_box("分析完成后显示。", 185)
            setattr(self, attr, box)
            body.addWidget(box)
            pair.addWidget(frame, 1)
        layout.addLayout(pair)
        frame, body = card()
        body.addWidget(label("证据原句与经历概览", "sectionTitle"))
        self.evidence_text = read_only_box("分析完成后显示。", 210)
        body.addWidget(self.evidence_text)
        layout.addWidget(frame)
        frame, body = card()
        body.addWidget(label("岗位排序", "sectionTitle"))
        self.rank_table = self._table(["排序", "岗位", "企业", "证据匹配度", "来源"], 230)
        body.addWidget(self.rank_table)
        layout.addWidget(frame)
        return page

    def _suggestions_page(self) -> QScrollArea:
        page, layout = scroll_page()
        intro, body = card()
        body.addWidget(label("原句 → 问题 → 建议替换句", "sectionTitle"))
        body.addWidget(label("替换句只重排原文中已写出的事实。信息不足时会标明需补充的证据，不会编造职责、成果或百分比。", "muted", True))
        layout.addWidget(intro)
        self.suggestions_container = QVBoxLayout()
        self.suggestions_container.setSpacing(14)
        layout.addLayout(self.suggestions_container)
        self._clear_suggestions()
        return page

    def _clear_suggestions(self) -> None:
        while self.suggestions_container.count():
            item = self.suggestions_container.takeAt(0)
            widget = item.widget()
            if widget:
                widget.setParent(None)
                widget.deleteLater()
        self.suggestions_container.addWidget(label("分析完成后，逐句修改建议会显示在这里。", "muted"))

    def _report_page(self) -> QScrollArea:
        page, layout = scroll_page()
        frame, body = card()
        body.addWidget(label("简历诊断报告", "sectionTitle"))
        body.addWidget(label("点击分析后报告会立即生成。报告含学历、专业、岗位需求、分维度指标、优势短板、原句改写与岗位排序。", "muted", True))
        row = QHBoxLayout()
        word = QPushButton("保存 Word 报告")
        word.setObjectName("primary")
        word.clicked.connect(self.save_report)
        structured = QPushButton("保存结构化 JSON")
        structured.clicked.connect(self.save_json)
        row.addWidget(word)
        row.addWidget(structured)
        row.addStretch()
        body.addLayout(row)
        layout.addWidget(frame)
        frame, body = card()
        body.addWidget(label("报告摘要", "sectionTitle"))
        self.report_preview = read_only_box("分析完成后显示报告摘要。", 330)
        body.addWidget(self.report_preview)
        layout.addWidget(frame)
        return page

    def choose_resume(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择学生简历", "", "简历文件 (*.pdf *.docx *.txt)")
        if not path:
            return
        try:
            text = read_document(path, Path(path).read_bytes())
        except (OSError, ValueError) as exc:
            QMessageBox.critical(self, "简历读取失败", str(exc))
            return
        self.resume_text.setPlainText(text)
        self.resume_path.setText(path)
        self.status.setText("简历已读取。请核对提取的文字，再开始分析。")

    def load_sample_resume(self) -> None:
        self.resume_text.setPlainText((ROOT / "samples" / "演示简历.txt").read_text(encoding="utf-8"))
        self.resume_path.setText("虚构演示简历，仅用于功能体验")
        self.student_input.setText("虚构演示学生")
        self.status.setText("已载入虚构演示简历。")

    def refresh_jobs(self) -> None:
        self.jobs = (self.demo_jobs if self.use_demo.isChecked() else []) + self.imported_jobs + self.manual_jobs
        rows = []
        for job in self.jobs:
            count = len(set(job.skills + job.tools + job.responsibilities + job.projects))
            rows.append([job.title, job.employer, str(count), job.source])
        self._fill_table(self.jobs_table, rows)
        if rows:
            self.jobs_table.selectRow(0)
        self.status.setText(f"当前有 {len(self.jobs)} 个岗位。请选择一行作为目标岗位。")

    def import_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "导入企业岗位 CSV", "", "CSV 文件 (*.csv)")
        if not path:
            return
        try:
            jobs = load_jobs_csv(Path(path).read_bytes())
        except (OSError, ValueError) as exc:
            QMessageBox.critical(self, "岗位导入失败", str(exc))
            return
        self.imported_jobs.extend(jobs)
        self.refresh_jobs()
        self.status.setText(f"已导入 {len(jobs)} 个岗位，请选择目标岗位。")

    def clear_imported(self) -> None:
        self.imported_jobs.clear()
        self.refresh_jobs()

    def clear_manual(self) -> None:
        self.manual_jobs.clear()
        self.refresh_jobs()

    def add_manual_job(self) -> None:
        try:
            job = job_from_description(
                title=self.title_input.text(), description=self.jd_text.toPlainText(),
                employer=self.employer_input.text(), min_degree=self.degree_input.currentText(),
                majors=self.majors_input.text(), skills=self.skills_input.text(),
                tools=self.tools_input.text(), responsibilities=self.duties_input.text(),
                projects=self.projects_input.text(),
            )
        except ValueError as exc:
            QMessageBox.warning(self, "岗位信息不完整", str(exc))
            return
        self.manual_jobs.append(job)
        self.refresh_jobs()
        self.jobs_table.selectRow(len(self.jobs) - 1)
        self.status.setText(f"已添加岗位：{job.employer} · {job.title}。")

    def analyze(self) -> None:
        text = self.resume_text.toPlainText().strip()
        if len(text) < 20:
            self.show_page(0)
            QMessageBox.warning(self, "缺少简历", "请先上传简历、载入演示简历或粘贴至少 20 个字的简历内容。")
            return
        index = self.jobs_table.currentRow()
        if not self.jobs or index < 0:
            self.show_page(1)
            QMessageBox.warning(self, "缺少目标岗位", "请先在岗位列表选中一个目标岗位。")
            return
        try:
            result = analyze_resume(text, self.jobs, index)
            report = build_report(result, self.student_input.text().strip() or "学生（匿名）")
        except Exception as exc:
            QMessageBox.critical(self, "分析失败", f"{type(exc).__name__}：{exc}")
            return
        self.result = result
        self.report_bytes = report
        self._show_result()
        self.show_page(2)
        self.status.setText("分析与报告已生成。可查看原句修改建议，并在报告页面保存 Word 文件。")

    def _show_result(self) -> None:
        assert self.result is not None
        result = self.result
        target = result["target"]
        profile = result["profile"]
        assessment = result["assessment"]
        self.score_value.setText(f"{target['score']} 分" if target["score"] is not None else "未评分")
        self.strength_value.setText(str(len(assessment["优势"])))
        self.gap_value.setText(str(len(assessment["待改进"])))
        self.target_label.setText(f"目标岗位：{target['employer']} · {target['title']}  ｜  来源：{target['source']}")
        self._fill_table(self.dim_table, [[
            item["维度"], f"{item['覆盖率']}%", f"{item['权重']}%", f"{item['加权得分']} 分",
            item["状态"], "、".join(item["未找到"]) or "无",
        ] for item in assessment["指标"]])
        self.strength_text.setPlainText("\n\n".join(assessment["优势"]) or "当前简历尚无达到‘文本覆盖较高’阈值的岗位维度。")
        self.weakness_text.setPlainText("\n\n".join(assessment["待改进"]) or "当前规则未识别到明显证据缺口。")
        evidence = [
            f"学历：{profile.degree or '未识别'}",
            f"专业：{'、'.join(profile.majors) or '未识别'}",
            f"专业技能：{'、'.join(profile.skills) or '未识别'}",
            f"工具应用：{'、'.join(profile.tools) or '未识别'}", "",
        ]
        for item in assessment["指标"]:
            evidence.append(f"【{item['维度']}】")
            evidence.extend(f"{term}：{line}" for term, line in list(item["证据原句"].items())[:6])
            if not item["证据原句"]:
                evidence.append("未找到直接证据原句。")
            evidence.append("")
        evidence.extend(["工作/实习经历：", *(profile.work_lines[:8] or ["未在对应栏目识别到文字。"]), "",
                         "项目经历：", *(profile.project_lines[:8] or ["未在对应栏目识别到文字。"]), "",
                         "成果线索：", *(profile.outcome_lines[:8] or ["未识别到成果描述，请人工核对。"])])
        self.evidence_text.setPlainText("\n".join(evidence))
        self._fill_table(self.rank_table, [[
            str(number), job["title"], job["employer"],
            f"{job['score']} 分" if job["score"] is not None else "未评分",
            "演示样例" if job["source"].startswith("演示") else job["source"],
        ] for number, job in enumerate(result["ranked"], 1)])
        self._show_suggestions(result["suggestions"])
        rewrite_count = sum(bool(item["建议替换句"]) for item in result["suggestions"])
        preview = [
            f"学生称呼：{self.student_input.text().strip() or '学生（匿名）'}",
            f"目标岗位：{target['employer']} · {target['title']}",
            f"简历证据匹配度：{self.score_value.text()}",
            f"学历：{profile.degree or '未识别'}；专业：{'、'.join(profile.majors) or '未识别'}",
            f"已展示优势：{len(assessment['优势'])} 条；待改进：{len(assessment['待改进'])} 条",
            f"逐项建议：{len(result['suggestions'])} 条；可直接核对的替换句：{rewrite_count} 条", "",
            "分维度指标：",
            *(f"• {item['维度']}：覆盖率 {item['覆盖率']}%，加权得分 {item['加权得分']} 分；{item['状态']}"
              for item in assessment["指标"]), "",
            "优势：", *(assessment["优势"] or ["暂无达到‘文本覆盖较高’阈值的维度。"]), "",
            "待改进：", *(assessment["待改进"] or ["暂无明显缺口。"]), "",
            result["disclaimer"],
        ]
        self.report_preview.setPlainText("\n".join(preview))

    def _show_suggestions(self, suggestions: list[dict[str, str]]) -> None:
        while self.suggestions_container.count():
            item = self.suggestions_container.takeAt(0)
            widget = item.widget()
            if widget:
                widget.setParent(None)
                widget.deleteLater()
        if not suggestions:
            self.suggestions_container.addWidget(label("当前规则未发现明确话术问题，请仍由人工核对事实与表达。", "muted"))
            return
        for number, suggestion in enumerate(suggestions, 1):
            frame, body = card()
            body.addWidget(label(f"{number:02d}  {suggestion['位置']} · {suggestion['类型']}", "sectionTitle"))
            body.addWidget(label("原句 / 现状", "small"))
            body.addWidget(label(suggestion["原文"], wrap=True))
            body.addWidget(label("问题", "small"))
            body.addWidget(label(suggestion["问题"], "attention", True))
            replacement = suggestion["建议替换句"]
            if replacement:
                body.addWidget(label("建议替换句", "small"))
                body.addWidget(label(replacement, "good", True))
                copy = QPushButton("复制替换句")
                copy.clicked.connect(lambda checked=False, text=replacement: QGuiApplication.clipboard().setText(text))
                body.addWidget(copy, 0, Qt.AlignLeft)
            else:
                body.addWidget(label("暂不能安全生成替换句：原句缺少可核实的具体事实。", "muted", True))
            body.addWidget(label(f"核实与补充：{suggestion['需核实']}", "muted", True))
            self.suggestions_container.addWidget(frame)

    def save_report(self) -> None:
        if self.report_bytes is None:
            QMessageBox.information(self, "尚无报告", "请先完成一次简历分析。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "保存 Word 诊断报告", "机械类学生简历诊断与岗位匹配报告.docx", "Word 文档 (*.docx)")
        if not path:
            return
        try:
            Path(path).write_bytes(self.report_bytes)
        except OSError as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            return
        self.status.setText(f"Word 报告已保存：{path}")
        QMessageBox.information(self, "保存成功", f"Word 报告已保存到：\n{path}")

    def save_json(self) -> None:
        if self.result is None:
            QMessageBox.information(self, "尚无分析数据", "请先完成一次简历分析。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "保存结构化分析", "简历诊断分析.json", "JSON 文件 (*.json)")
        if not path:
            return
        try:
            Path(path).write_text(json.dumps(to_json_ready(self.result), ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            return
        self.status.setText(f"JSON 已保存：{path}")
        QMessageBox.information(self, "保存成功", f"结构化分析已保存到：\n{path}")


def main() -> None:
    application = QApplication(sys.argv)
    application.setFont(QFont("Microsoft YaHei UI", 10))
    application.setStyleSheet(STYLE)
    window = ResumeDesktopApp()
    if "--smoke" in sys.argv:
        window.load_sample_resume()
        window.analyze()
        assert window.result is not None and window.report_bytes is not None
        assert any(item["建议替换句"] for item in window.result["suggestions"])
        assert window.result["assessment"]["指标"]
        window.resume_text.insertPlainText("\n补充待核对信息")
        assert window.result is None and window.report_bytes is None
        print("Desktop UI smoke test passed")
        return
    window.show()
    window.raise_()
    window.activateWindow()
    sys.exit(application.exec())


if __name__ == "__main__":
    main()
