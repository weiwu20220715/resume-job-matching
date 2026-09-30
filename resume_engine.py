"""Local, explainable resume analysis for the employment guidance prototype.

The score measures evidence *written in the resume* against a supplied job
profile. It must never be interpreted as a hiring probability or a person's
actual ability. No resume content is sent to an external service.
"""

from __future__ import annotations

import csv
import io
import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from pypdf import PdfReader


MAX_UPLOAD_BYTES = 10 * 1024 * 1024
ALLOWED_SUFFIXES = {".pdf", ".docx", ".txt"}

# Synonyms are deliberately limited and visible, so a false match can be audited.
ALIASES: dict[str, tuple[str, ...]] = {
    "AutoCAD": ("autocad", "auto cad", "cad制图", "cad绘图"),
    "SolidWorks": ("solidworks", "solid works", "solid-work", "sw建模"),
    "Creo": ("creo", "pro/e", "proe"),
    "NX": ("ug/nx", "ug nx", "siemens nx", "ug软件", "nx软件"),
    "CATIA": ("catia",),
    "ANSYS": ("ansys",),
    "MATLAB": ("matlab",),
    "Python": ("python",),
    "PLC": ("plc", "s7-1200", "s7-1500", "西门子s7"),
    "MES": ("mes系统", "mes平台", "制造执行系统", "mes"),
    "工业机器人": ("工业机器人", "机器人示教", "机器人编程"),
    "机械设计": ("机械设计", "结构设计", "零部件设计"),
    "三维建模": ("三维建模", "3d建模", "三维设计"),
    "工程制图": ("工程制图", "机械制图", "二维制图", "cad制图"),
    "公差分析": ("公差分析", "尺寸链", "形位公差"),
    "有限元分析": ("有限元分析", "有限元仿真", "fea分析"),
    "工艺规划": ("工艺规划", "工艺设计", "工艺编制", "工艺路线"),
    "产线调试": ("产线调试", "生产线调试", "设备调试", "联调联试"),
    "设备维护": ("设备维护", "设备保养", "故障排查", "故障诊断"),
    "质量分析": ("质量分析", "质量改进", "质量控制", "质量管理"),
    "数据分析": ("数据分析", "数据处理", "数据可视化"),
    "跨部门沟通": ("跨部门沟通", "跨团队沟通", "跨部门协作", "跨团队协作"),
    "项目管理": ("项目管理", "进度管理", "项目计划", "项目协调"),
    "技术汇报": ("技术汇报", "方案汇报", "技术方案汇报", "项目汇报"),
    "数字孪生": ("数字孪生",),
    "智能制造": ("智能制造", "智能产线", "智能生产"),
}

DEGREE_RANK = {"中专": 0, "大专": 1, "本科": 2, "硕士": 3, "博士": 4}
DEGREE_PATTERN = re.compile(r"博士|硕士|研究生|本科|学士|大专|专科|中专")
KNOWN_MAJORS = (
    "机械设计制造及其自动化", "智能制造工程", "机械工程", "机械电子工程",
    "材料成型及控制工程", "工业工程", "自动化", "电气工程及其自动化",
    "机器人工程", "车辆工程", "工业设计", "计算机科学与技术",
)
MAJOR_FAMILIES = {
    "机械类": ("机械设计制造及其自动化", "机械工程", "机械电子工程", "材料成型及控制工程", "车辆工程"),
    "自动化类": ("自动化", "电气工程及其自动化", "机器人工程"),
}
SECTION_PATTERN = re.compile(
    r"^(教育背景|教育经历|学历教育|个人信息|基本信息|求职意向|专业技能|技能证书|"
    r"技能特长|技能掌握|项目经历|项目经验|科研项目|实习经历|工作经历|实践经历|"
    r"校园经历|获奖经历|荣誉奖励|竞赛经历|自我评价|个人总结)\s*[:：]?$"
)
SECTION_MAP = {
    "教育背景": "教育", "教育经历": "教育", "学历教育": "教育",
    "个人信息": "基本信息", "基本信息": "基本信息", "求职意向": "求职意向",
    "专业技能": "技能", "技能证书": "技能", "技能特长": "技能", "技能掌握": "技能",
    "项目经历": "项目", "项目经验": "项目", "科研项目": "项目", "竞赛经历": "项目",
    "实习经历": "工作", "工作经历": "工作", "实践经历": "工作", "校园经历": "工作",
    "获奖经历": "成果", "荣誉奖励": "成果", "自我评价": "自评", "个人总结": "自评",
}
WEIGHTS = {"学历": 10, "专业": 10, "专业技能": 30, "工具应用": 20, "工作职责": 15, "项目经验": 15}


@dataclass
class ResumeProfile:
    degree: str | None
    majors: list[str]
    skills: list[str]
    tools: list[str]
    sections: dict[str, list[str]]
    project_lines: list[str]
    work_lines: list[str]
    outcome_lines: list[str]
    evidence: dict[str, str]
    warnings: list[str] = field(default_factory=list)


@dataclass
class JobProfile:
    title: str
    employer: str = "未指定企业"
    min_degree: str = ""
    majors: list[str] = field(default_factory=list)
    skills: list[str] = field(default_factory=list)
    tools: list[str] = field(default_factory=list)
    responsibilities: list[str] = field(default_factory=list)
    projects: list[str] = field(default_factory=list)
    description: str = ""
    source: str = "用户输入"


def _clean(text: str) -> str:
    return re.sub(r"\s+", "", text).lower()


def canonical(term: str) -> str:
    candidate = term.strip()
    compact = _clean(candidate)
    if not compact:
        return ""
    for name, variants in ALIASES.items():
        if compact == _clean(name) or any(compact == _clean(v) for v in variants):
            return name
    return candidate


def split_terms(value: str | list[str] | None) -> list[str]:
    if isinstance(value, list):
        parts = value
    elif value:
        parts = re.split(r"[,，;；、\n]+", str(value))
    else:
        parts = []
    output: list[str] = []
    for part in parts:
        term = canonical(str(part))
        if term and term not in output:
            output.append(term)
    return output


def read_document(filename: str, data: bytes) -> str:
    """Extract selectable text. Scanned PDFs are reported as unsupported."""
    suffix = Path(filename).suffix.lower()
    if suffix not in ALLOWED_SUFFIXES:
        raise ValueError("仅支持 PDF、DOCX、TXT 文件。")
    if not data or len(data) > MAX_UPLOAD_BYTES:
        raise ValueError("文件为空或大于 10 MB。")
    if suffix == ".txt":
        for encoding in ("utf-8-sig", "gb18030"):
            try:
                result = data.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("TXT 编码无法识别，请保存为 UTF-8。")
    elif suffix == ".docx":
        try:
            doc = Document(io.BytesIO(data))
            parts = [p.text for p in doc.paragraphs if p.text.strip()]
            parts += [c.text for table in doc.tables for row in table.rows for c in row.cells if c.text.strip()]
            result = "\n".join(parts)
        except Exception as exc:
            raise ValueError("DOCX 读取失败，请检查文件是否损坏。") from exc
    else:
        try:
            pdf = PdfReader(io.BytesIO(data))
            if len(pdf.pages) > 40:
                raise ValueError("PDF 超过 40 页，请上传简历或精简文件。")
            result = "\n".join(page.extract_text() or "" for page in pdf.pages)
        except ValueError:
            raise
        except Exception as exc:
            raise ValueError("PDF 读取失败，请检查文件是否损坏或已加密。") from exc
    if len(result.strip()) < 20:
        raise ValueError("未提取到足够文字。扫描版 PDF 或图片简历请先 OCR 后重新上传。")
    return result[:100_000]


def _lines(text: str) -> list[str]:
    return [re.sub(r"\s+", " ", line).strip(" -•●\t") for line in text.splitlines() if line.strip()]


def _find_evidence(text: str, term: str) -> str | None:
    variants = (term,) + ALIASES.get(term, ())
    for line in _lines(text):
        normal = _clean(line)
        if any(_clean(v) in normal for v in variants if v):
            return line[:240]
    return None


def _find_degree(text: str) -> str | None:
    found: list[str] = []
    for match in DEGREE_PATTERN.findall(text):
        level = {"研究生": "硕士", "学士": "本科", "专科": "大专"}.get(match, match)
        found.append(level)
    return max(found, key=lambda x: DEGREE_RANK[x]) if found else None


def _detect_majors(text: str) -> list[str]:
    found = [major for major in KNOWN_MAJORS if _find_evidence(text, major)]
    return [major for major in found if not any(major != other and major in other for other in found)]


def parse_resume(text: str) -> ResumeProfile:
    sections: dict[str, list[str]] = {name: [] for name in set(SECTION_MAP.values())}
    active = "其他"
    sections[active] = []
    for line in _lines(text):
        head = SECTION_PATTERN.fullmatch(line)
        if head:
            active = SECTION_MAP[head.group(1)]
        else:
            sections[active].append(line)
    identity_text = "\n".join(sections["教育"] + sections["基本信息"] + sections["其他"][:5])
    majors = _detect_majors(identity_text)
    extra_major = re.search(r"(?:所学专业|专业)\s*[:：]\s*([^\s,，;；]{2,24})", identity_text)
    if extra_major and extra_major.group(1) not in majors:
        majors.append(extra_major.group(1))
    tools_vocab = {"AutoCAD", "SolidWorks", "Creo", "NX", "CATIA", "ANSYS", "MATLAB", "Python", "PLC", "MES"}
    ability_text = "\n".join(line for name in ("技能", "项目", "工作", "成果", "自评", "其他") for line in sections[name])
    # The education line alone must not count as proof of a skill or tool.
    evidence = {key: found for key in ALIASES if (found := _find_evidence(ability_text, key))}
    tools = [key for key in evidence if key in tools_vocab]
    skills = [key for key in evidence if key not in tools_vocab]
    projects = sections.get("项目", [])
    work = sections.get("工作", [])
    outcome_pattern = re.compile(r"\d+(?:\.\d+)?\s*(?:%|％|项|台|个|份|件|人|小时|分钟|万元|元|次)|专利|获奖|一等奖|二等奖|三等奖|提升|降低|减少")
    outcomes = [line for line in projects + work + sections.get("成果", []) if outcome_pattern.search(line)]
    warnings = []
    if not projects:
        warnings.append("未识别出“项目经历”栏目；项目能力可能被低估。")
    if not work:
        warnings.append("未识别出“实习/工作经历”栏目；职责证据可能被低估。")
    if not _find_degree(text):
        warnings.append("未识别到学历，请核对教育经历是否写明。")
    return ResumeProfile(_find_degree(text), majors, skills, tools, sections, projects, work, outcomes, evidence, warnings)


def job_from_dict(item: dict[str, Any], source: str = "用户输入") -> JobProfile:
    def pick(*names: str) -> Any:
        return next((item[name] for name in names if name in item and item[name] is not None), "")
    title = str(pick("岗位名称", "title")).strip()
    if not title:
        raise ValueError("岗位缺少“岗位名称”列。")
    degree = str(pick("最低学历", "min_degree")).strip()
    if degree and degree not in DEGREE_RANK:
        raise ValueError(f"岗位“{title}”的最低学历无效：{degree}")
    return JobProfile(
        title=title,
        employer=str(pick("企业名称", "employer") or "未指定企业").strip(),
        min_degree=degree,
        majors=split_terms(pick("相关专业", "majors")),
        skills=split_terms(pick("专业技能", "skills")),
        tools=split_terms(pick("工具软件", "tools")),
        responsibilities=split_terms(pick("工作职责", "responsibilities")),
        projects=split_terms(pick("项目关键词", "projects")),
        description=str(pick("岗位描述", "description")),
        source=source,
    )


def load_demo_jobs(path: str | Path) -> list[JobProfile]:
    with open(path, "r", encoding="utf-8") as handle:
        return [job_from_dict(item, "演示样例，非企业真实招聘") for item in json.load(handle)]


def load_jobs_csv(data: bytes) -> list[JobProfile]:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            content = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("CSV 编码无法识别，请保存为 UTF-8。")
    rows = list(csv.DictReader(io.StringIO(content)))
    if not rows:
        raise ValueError("CSV 没有岗位记录。")
    if len(rows) > 500:
        raise ValueError("一次最多导入 500 个岗位。")
    return [job_from_dict(row, "用户导入岗位") for row in rows]


def job_from_description(
    title: str, description: str, employer: str = "未指定企业",
    min_degree: str = "", majors: str = "", skills: str = "", tools: str = "",
    responsibilities: str = "", projects: str = "",
) -> JobProfile:
    """Turn pasted JD into a profile; explicit fields override dictionary extraction."""
    if not title.strip():
        raise ValueError("请填写岗位名称。")
    if not description.strip() and not any((majors, skills, tools, responsibilities, projects)):
        raise ValueError("请粘贴岗位说明或填写岗位要求。")
    degrees = [_find_degree(value) for value in re.findall(r"博士|硕士|研究生|本科|学士|大专|专科|中专", description)]
    detected_degree = min((value for value in degrees if value), key=lambda value: DEGREE_RANK[value], default=None)
    detected_majors = _detect_majors(description)
    for family in MAJOR_FAMILIES:
        if (family in description or family.replace("类", "相关专业") in description) and family not in detected_majors:
            detected_majors.append(family)
    tool_names = {"AutoCAD", "SolidWorks", "Creo", "NX", "CATIA", "ANSYS", "MATLAB", "Python", "PLC", "MES"}
    detected_tools = [term for term in ALIASES if term in tool_names and _find_evidence(description, term)]
    detected_skills = [term for term in ALIASES if term not in tool_names and _find_evidence(description, term)]
    item = {
        "岗位名称": title.strip(), "企业名称": employer.strip() or "未指定企业",
        "最低学历": min_degree or detected_degree or "",
        "相关专业": majors or detected_majors,
        "专业技能": skills or detected_skills,
        "工具软件": tools or detected_tools,
        "工作职责": responsibilities,
        "项目关键词": projects,
        "岗位描述": description,
    }
    return job_from_dict(item, "用户粘贴岗位说明")


def _major_match(profile: ResumeProfile, options: list[str]) -> tuple[float, list[str], list[str]]:
    if not options:
        return 0.0, [], []
    if not profile.majors:
        return 0.0, [], options
    for actual in profile.majors:
        for target in options:
            family = next((members for label, members in MAJOR_FAMILIES.items()
                           if label in target or label.replace("类", "相关专业") in target), None)
            if family and actual in family:
                return 1.0, [target], []
            if _clean(actual) in _clean(target) or _clean(target) in _clean(actual):
                return 1.0, [target], []
    return 0.0, [], options


def _term_match(text: str, terms: list[str]) -> tuple[float, list[str], list[str], dict[str, str]]:
    if not terms:
        return 0.0, [], [], {}
    matched, missing, evidence = [], [], {}
    for term in terms:
        line = _find_evidence(text, canonical(term))
        if line:
            matched.append(term)
            evidence[term] = line
        else:
            missing.append(term)
    return len(matched) / len(terms), matched, missing, evidence


def match_job(text: str, profile: ResumeProfile, job: JobProfile) -> dict[str, Any]:
    """Return auditable evidence coverage, not a validated hiring score."""
    categories: dict[str, dict[str, Any]] = {}
    if job.min_degree:
        value = float(profile.degree is not None and DEGREE_RANK[profile.degree] >= DEGREE_RANK[job.min_degree])
        categories["学历"] = {"ratio": value, "matched": [profile.degree] if value else [],
                           "missing": [] if value else [f"最低{job.min_degree}"],
                           "evidence": {profile.degree: _find_evidence(text, profile.degree) or ""} if value else {}}
    if job.majors:
        ratio, found, missing = _major_match(profile, job.majors)
        categories["专业"] = {"ratio": ratio, "matched": found, "missing": missing,
                           "evidence": {found[0]: _find_evidence(text, profile.majors[0]) or ""} if found else {}}
    section_text = {k: "\n".join(v) for k, v in profile.sections.items()}
    ability_text = "\n".join(section_text.get(name, "") for name in ("技能", "项目", "工作", "成果", "自评", "其他"))
    # Skills/tools can be documented anywhere in the resume. Duties and
    # projects need contextual experience evidence to avoid keyword stuffing.
    sources = {
        "专业技能": (ability_text, job.skills),
        "工具应用": (ability_text, job.tools),
        "工作职责": (section_text.get("工作", "") + "\n" + section_text.get("项目", ""), job.responsibilities),
        "项目经验": (section_text.get("项目", ""), job.projects),
    }
    for category, (source_text, terms) in sources.items():
        if terms:
            ratio, found, missing, evidence = _term_match(source_text, terms)
            categories[category] = {"ratio": ratio, "matched": found, "missing": missing, "evidence": evidence}
    requirement_count = len(set(job.skills + job.tools + job.responsibilities + job.projects))
    available_weight = sum(WEIGHTS[name] for name in categories)
    score = (round(sum(WEIGHTS[name] * item["ratio"] for name, item in categories.items()) / available_weight * 100, 1)
             if available_weight and requirement_count >= 3 else None)
    for name, item in categories.items():
        item["weight"] = round(WEIGHTS[name] / available_weight * 100, 1) if available_weight else 0
        item["points"] = round(item["weight"] * item["ratio"], 1)
    gaps = [f"{name}：{', '.join(info['missing'][:4])}" for name, info in categories.items() if info["missing"]]
    return {"title": job.title, "employer": job.employer, "source": job.source,
            "description": job.description[:2000],
            "requirements": {"最低学历": job.min_degree, "相关专业": job.majors,
                             "专业技能": job.skills, "工具软件": job.tools,
                             "工作职责": job.responsibilities, "项目关键词": job.projects},
            "score": score, "categories": categories, "gaps": gaps,
            "requirement_count": requirement_count,
            "warning": "岗位要求少于 3 个具体技能、工具、职责或项目条目，暂不评分。" if score is None else ""}


def _sentence_rewrite(line: str, section: str, heading: str = "") -> tuple[str, str, str] | None:
    """Rewrite only facts already present in a line; never add a metric or duty."""
    original = line.strip()
    clean = original.rstrip("。；; ")
    if len(clean) < 5 or len(clean) > 180:
        return None
    looks_like_heading = (heading and len(heading) <= 40 and not re.search(r"[。；;，,]", heading)
                          and not re.match(r"^(?:参与|负责|协助|使用|完成|输出|形成|编制|熟悉|了解|掌握)", heading))
    prefix = f"在{heading}中，" if looks_like_heading else ""
    # A multi-clause project sentence already states the actual tool, task and
    # deliverable. Rearranging those clauses makes the contribution readable.
    pattern = re.match(r"^参与([^，,。；;]+)[，,]\s*使用\s*([^，,。；;]+)[，,]\s*(输出|形成|完成|编制)(.+)$", clean)
    if pattern:
        task, action, verb, outcome = (part.strip() for part in pattern.groups())
        replacement = f"{prefix}参与{task}；使用 {action}，并{verb} {outcome}。"
        return "多个动作和成果堆在同一句，具体任务和产出不够醒目。", replacement, "请核实‘使用’和‘输出/完成’确由本人承担。"
    pattern = re.match(r"^协助([^，,。；;]+)[，,]\s*参与([^，,。；;]+)$", clean)
    if pattern:
        first, second = (part.strip() for part in pattern.groups())
        return "两个动作并列但职责边界不清，建议分别写明。", f"{prefix}协助{first}，并参与{second}。", "如有本人具体操作或交付物，请继续补充；不要把协助写成独立负责。"
    pattern = re.match(r"^参与([^，,。；;]+)[，,]\s*(完成|输出|形成|编制)(.+)$", clean)
    if pattern:
        task, verb, result = (part.strip() for part in pattern.groups())
        return "具体产出位于句尾，读者不易迅速看出贡献。", f"{prefix}参与{task}，并{verb}{result}。", "请确认该产出确由本人完成；若只是协作，应保留‘参与’。"
    if re.match(r"^(?:熟悉|了解|掌握|精通|学习过)", clean):
        return ("仅写掌握程度，缺少使用场景和作品，不能判断实际应用深度。", "",
                "请提供真实课程、项目或实习中的操作与交付物；在核实前不生成可能夸大能力的替换句。")
    if re.match(r"^(?:负责|参与|协助)", clean):
        return ("只写职责动词，缺少明确的本人动作、工具或结果。", "",
                "请补充本人做了什么、用了什么、留下什么可核实的成果；事实不足时不能安全改写为‘独立完成’。")
    return None


def make_suggestions(text: str, profile: ResumeProfile, result: dict[str, Any]) -> list[dict[str, str]]:
    suggestions: list[dict[str, str]] = []
    for section, lines in (("专业技能", profile.sections.get("技能", [])),
                           ("项目经历", profile.project_lines), ("实习/工作经历", profile.work_lines),
                           ("自我评价", profile.sections.get("自评", []))):
        for position, line in enumerate(lines):
            previous = lines[position - 1] if position > 0 else ""
            finding = _sentence_rewrite(line, section, previous)
            if finding is None:
                continue
            problem, replacement, check = finding
            suggestions.append({"类型": "逐句改写", "位置": section, "原文": line,
                                "问题": problem, "建议替换句": replacement,
                                "需核实": check, "建议": replacement or check})
            if len(suggestions) >= 15:
                break
        if len(suggestions) >= 15:
            break
    if not profile.degree or not profile.majors:
        suggestions.append({"类型": "信息缺口", "位置": "教育经历", "原文": "未识别到完整学历或专业信息",
                            "问题": "岗位筛选条件缺少可核对的依据。", "建议替换句": "",
                            "需核实": "按真实情况补充学校、专业、学历层次和起止时间；在读时写明预计毕业时间。",
                            "建议": "按真实情况补充学校、专业、学历层次和起止时间。"})
    for category in ("专业技能", "工具应用", "工作职责", "项目经验"):
        gap = result["categories"].get(category, {}).get("missing", [])
        if gap:
            note = "若有真实经历，在相应项目或实习条目补充本人任务、实际工具和可核实成果；没有则列为学习计划。"
            suggestions.append({"类型": "岗位证据缺口", "位置": category, "原文": "简历中未找到直接证据",
                                "问题": f"目标岗位要求：{'、'.join(gap[:5])}。未写出不等于不会，请先核实经历。",
                                "建议替换句": "", "需核实": note, "建议": note})
    if not profile.outcome_lines:
        note = "补充真实的图纸、程序、报告、样机、检测记录或竞赛成果；有原始记录时再写数量或改进幅度。"
        suggestions.append({"类型": "成果证据缺口", "位置": "项目成果", "原文": "未识别到成果或交付物描述",
                            "问题": "经历描述中缺少结果证据。", "建议替换句": "", "需核实": note, "建议": note})
    return suggestions[:24]


def assess_strengths_weaknesses(profile: ResumeProfile, result: dict[str, Any]) -> dict[str, Any]:
    """Describe documented evidence, not the student's unobserved ability."""
    metrics: list[dict[str, Any]] = []
    strengths: list[str] = []
    weaknesses: list[str] = []
    for name, info in result["categories"].items():
        ratio = info["ratio"]
        matched, missing = info["matched"], info["missing"]
        if ratio >= 0.75:
            level = "文本覆盖较高"
            strengths.append(f"{name}：简历文本覆盖岗位要求的 {ratio * 100:.0f}%；已识别{'、'.join(matched[:4]) or '相关信息'}。")
        elif ratio > 0:
            level = "部分覆盖"
            weaknesses.append(f"{name}：已覆盖 {ratio * 100:.0f}%，仍未在简历中找到{'、'.join(missing[:4]) or '部分要求'}。")
        else:
            level = "未找到直接证据"
            weaknesses.append(f"{name}：当前简历未展示该岗位要求的直接证据；需核实{'、'.join(missing[:4]) or '相关经历'}。")
        metrics.append({"维度": name, "覆盖率": round(ratio * 100, 1), "权重": info["weight"],
                        "加权得分": info["points"], "状态": level,
                        "已找到": matched, "未找到": missing, "证据原句": info["evidence"]})
        if name in {"专业技能", "工具应用"} and matched and info["evidence"]:
            lines = list(info["evidence"].values())
            if all(re.match(r"^(?:学习过|熟悉|了解|掌握|精通|工具[:：])", line) for line in lines):
                weaknesses.append(f"{name}应用深度：目前主要是能力或工具罗列，缺少在项目/实习中如何使用的具体原句。")
    if profile.outcome_lines:
        strengths.append(f"成果表达：识别到 {len(profile.outcome_lines)} 条可能包含数量或结果的原句，需核对真实性与本人贡献。")
    else:
        weaknesses.append("成果表达：未识别到明确的数量或交付物结果，项目价值不易判断。")
    if not profile.work_lines:
        weaknesses.append("工作/实习经历：未在对应栏目识别到职责描述，无法从简历判断实践深度。")
    if not profile.project_lines:
        weaknesses.append("项目经历：未在对应栏目识别到项目描述，无法从简历判断项目贡献。")
    return {"指标": metrics, "优势": strengths, "待改进": weaknesses}


def analyze_resume(text: str, jobs: list[JobProfile], selected_index: int | None = None) -> dict[str, Any]:
    if not jobs:
        raise ValueError("请先添加至少一个岗位。")
    profile = parse_resume(text)
    results = [match_job(text, profile, job) for job in jobs]
    ranked = sorted(results,
                    key=lambda item: item["score"] if item["score"] is not None else -1, reverse=True)
    target = results[selected_index] if selected_index is not None and 0 <= selected_index < len(results) else ranked[0]
    return {"profile": profile, "ranked": ranked, "target": target,
            "assessment": assess_strengths_weaknesses(profile, target),
            "suggestions": make_suggestions(text, profile, target),
            "disclaimer": "分数表示简历对岗位要求的书面证据覆盖度，未经过企业招聘结果验证，不代表录用概率或学生真实能力。"}


def to_json_ready(analysis: dict[str, Any]) -> dict[str, Any]:
    return {**analysis, "profile": asdict(analysis["profile"])}
