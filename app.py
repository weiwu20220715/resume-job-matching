"""Streamlit front end for the local resume and job matching prototype."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from reporting import build_report
from resume_engine import (
    analyze_resume, job_from_description, load_demo_jobs, load_jobs_csv,
    read_document, to_json_ready,
)


ROOT = Path(__file__).parent
st.set_page_config(page_title="机械类学生简历诊断与岗位匹配", page_icon="📄", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1280px; padding-top: 2rem;}
.stMetric {background:#f4f7fa; border:1px solid #e5ebf0; border-radius:10px; padding:12px;}
div[data-testid="stAlert"] {border-radius:8px;}
</style>
""", unsafe_allow_html=True)

st.title("机械类学生简历诊断与岗位匹配")
st.caption("本地可运行原型 · 简历修改建议 · 岗位要求对照 · 可编辑诊断报告")
st.info("本版的分数是“简历中可见证据对岗位要求的覆盖度”，并非招聘录用概率。演示岗位是样例；正式使用前请导入真实、经核对的企业岗位要求。")

with st.sidebar:
    st.header("使用步骤")
    st.markdown("1. 上传简历或粘贴文本\n2. 导入或填写企业岗位要求\n3. 选择目标岗位并分析\n4. 下载 Word 报告")
    st.divider()
    st.subheader("隐私与范围")
    st.write("简历只在当前本地进程内解析，程序不会主动上传到云端，也不会写入数据库。关闭会话后请自行管理下载的报告。")
    st.write("支持可提取文字的 PDF、DOCX、TXT；扫描版 PDF 需先 OCR。")
    st.write("当前岗位词典与权重是演示配置，尚待企业调研和专家论证。")

resume_col, job_col = st.columns([1, 1], gap="large")

with resume_col:
    st.header("1 上传学生简历")
    resume_file = st.file_uploader("选择简历文件", type=["pdf", "docx", "txt"], help="上限 10 MB；建议先去除身份证号、手机号等非必要个人信息。")
    pasted_resume = st.text_area("或直接粘贴简历文字", height=155, placeholder="教育经历、专业技能、项目经历、实习经历……")
    student_label = st.text_input("报告中的学生称呼", value="学生（匿名）", help="建议用匿名编号，不填真实姓名。")

with job_col:
    st.header("2 提供企业岗位要求")
    use_demo = st.checkbox("加入 5 个演示岗位", value=True)
    csv_file = st.file_uploader("批量导入岗位 CSV", type=["csv"], help="字段模板位于 data/jobs_template.csv。可包含多个真实岗位。")
    with open(ROOT / "data" / "jobs_template.csv", "rb") as handle:
        st.download_button("下载岗位 CSV 模板", handle.read(), file_name="岗位导入模板.csv", mime="text/csv")
    with st.expander("手动粘贴一个企业岗位说明", expanded=False):
        employer = st.text_input("企业名称", key="employer", placeholder="例如：某机械制造企业")
        job_title = st.text_input("岗位名称", key="job_title", placeholder="例如：机械设计工程师")
        jd = st.text_area("岗位说明原文", key="jd", height=110, placeholder="粘贴岗位职责与任职要求。系统自动识别词典内的学历、专业、技能和工具；建议核对并在下方补足字段。")
        degree = st.selectbox("最低学历（可选）", ["", "中专", "大专", "本科", "硕士", "博士"], key="degree")
        majors = st.text_input("相关专业（可选，顿号分隔）", key="majors")
        skills = st.text_input("专业技能（可选）", key="skills")
        tools = st.text_input("工具软件（可选）", key="tools")
        duties = st.text_input("工作职责关键词（建议填写）", key="duties")
        projects = st.text_input("项目经验关键词（建议填写）", key="projects")

jobs = []
if use_demo:
    jobs.extend(load_demo_jobs(ROOT / "data" / "demo_jobs.json"))
if csv_file is not None:
    try:
        jobs.extend(load_jobs_csv(csv_file.getvalue()))
    except ValueError as exc:
        st.error(f"岗位 CSV 导入失败：{exc}")
if job_title.strip():
    try:
        jobs.append(job_from_description(job_title, jd, employer, degree, majors, skills, tools, duties, projects))
    except ValueError as exc:
        st.warning(f"手动岗位暂未加入：{exc}")

st.divider()
st.header("3 选择岗位并开始分析")
if jobs:
    labels = [f"{job.employer}｜{job.title}（{job.source}）" for job in jobs]
    selected_index = st.selectbox("重点分析的目标岗位", range(len(jobs)), format_func=lambda i: labels[i])
    st.caption(f"当前有 {len(jobs)} 个岗位。岗位说明来源会显示在报告中。")
    with st.expander("查看目标岗位的结构化要求"):
        selected = jobs[selected_index]
        st.json({"岗位": selected.title, "企业": selected.employer, "最低学历": selected.min_degree,
                 "相关专业": selected.majors, "专业技能": selected.skills, "工具软件": selected.tools,
                 "工作职责": selected.responsibilities, "项目关键词": selected.projects}, expanded=False)
else:
    selected_index = None
    st.warning("请勾选演示岗位、导入 CSV 或填写一个岗位。")

if st.button("分析简历并生成报告", type="primary", disabled=not jobs, width="stretch"):
    try:
        if resume_file is not None:
            resume_text = read_document(resume_file.name, resume_file.getvalue())
        elif pasted_resume.strip():
            resume_text = pasted_resume.strip()
        else:
            raise ValueError("请先上传简历或粘贴简历文字。")
        if len(resume_text) < 20:
            raise ValueError("简历文字太少，无法进行可靠分析。")
        analysis = analyze_resume(resume_text, jobs, selected_index)
        if analysis["target"]["score"] is None:
            raise ValueError("目标岗位至少需要 3 个不同的技能、工具、职责或项目关键词；请补充岗位要求。")
        report_bytes = build_report(analysis, student_label.strip() or "学生（匿名）")
        st.session_state["result"] = analysis
        st.session_state["report"] = report_bytes
        st.session_state["student_label"] = student_label
    except ValueError as exc:
        st.error(f"分析未完成：{exc}")
    except Exception as exc:
        st.error(f"分析发生异常：{type(exc).__name__}。请检查文件格式或岗位字段。")
        st.exception(exc)

if "result" in st.session_state:
    analysis = st.session_state["result"]
    target = analysis["target"]
    profile = analysis["profile"]
    st.divider()
    st.header("分析结果")
    a, b, c, d = st.columns(4)
    a.metric("目标岗位证据匹配度", f"{target['score']} 分")
    b.metric("已识别学历", profile.degree or "未识别")
    c.metric("专业技能条目", len(profile.skills))
    d.metric("项目成果线索", len(profile.outcome_lines))
    st.caption(analysis["disclaimer"])

    portrait_tab, match_tab, improve_tab, report_tab = st.tabs(["学生画像", "岗位匹配与证据", "逐项修改建议", "报告下载"])
    with portrait_tab:
        st.subheader("简历信息提取")
        st.write(f"**学历层次：** {profile.degree or '未识别'}")
        st.write(f"**专业：** {'、'.join(profile.majors) or '未识别'}")
        st.write(f"**专业技能：** {'、'.join(profile.skills) or '未识别'}")
        st.write(f"**工具应用：** {'、'.join(profile.tools) or '未识别'}")
        for label, items in (("工作或实习职责", profile.work_lines), ("项目经验", profile.project_lines), ("成果线索", profile.outcome_lines)):
            st.markdown(f"**{label}**")
            if items:
                for line in items[:12]:
                    st.markdown(f"- {line}")
            else:
                st.write("未在对应栏目识别到文字。")
        for warning in profile.warnings:
            st.warning(warning)

    with match_tab:
        st.subheader(f"目标岗位：{target['employer']} · {target['title']}")
        chart_data = pd.DataFrame([{"维度": name, "覆盖率": round(info["ratio"] * 100, 1), "权重": info["weight"]}
                                   for name, info in target["categories"].items()])
        if not chart_data.empty:
            st.bar_chart(chart_data.set_index("维度")["覆盖率"], horizontal=True)
            if len(chart_data) >= 3:
                names = chart_data["维度"].tolist()
                values = chart_data["覆盖率"].tolist()
                fig = go.Figure(data=go.Scatterpolar(r=values + values[:1], theta=names + names[:1], fill="toself", name="简历证据"))
                fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), showlegend=False, height=390,
                                  margin=dict(l=30, r=30, t=15, b=15))
                st.plotly_chart(fig, width="stretch")
        st.dataframe(chart_data, hide_index=True, width="stretch")
        for name, info in target["categories"].items():
            with st.expander(f"{name} · 覆盖率 {info['ratio'] * 100:.0f}%"):
                st.write("**已找到：** " + ("、".join(info["matched"]) or "暂无"))
                st.write("**未在简历找到：** " + ("、".join(info["missing"]) or "暂无"))
                for term, sentence in info["evidence"].items():
                    st.write(f"- {term}：{sentence}")
        st.subheader("岗位推荐排序")
        st.dataframe(pd.DataFrame([{"岗位": item["title"], "企业": item["employer"], "证据匹配度": item["score"], "来源": item["source"]}
                                   for item in analysis["ranked"]]), hide_index=True, width="stretch")
        st.caption("少于 3 个具体要求的岗位不评分，会排在列表末尾。不同岗位要求的完整程度会影响分数可比性。")

    with improve_tab:
        if not analysis["suggestions"]:
            st.success("规则未发现明确问题，仍建议人工核对事实和岗位表达。")
        for index, item in enumerate(analysis["suggestions"], 1):
            with st.expander(f"建议 {index} · {item['位置']}", expanded=index <= 2):
                st.write("**原文或现状：** " + item["原文"])
                st.write("**问题：** " + item["问题"])
                st.write("**怎么改：** " + item["建议"])

    with report_tab:
        st.write("下载后可在 Word 中编辑，核对岗位原文、学生真实经历和每一条修改建议。")
        st.download_button("下载 Word 诊断报告", st.session_state["report"],
                           file_name="机械类学生简历诊断与岗位匹配报告.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           width="stretch")
        st.download_button("下载结构化分析 JSON", json.dumps(to_json_ready(analysis), ensure_ascii=False, indent=2).encode("utf-8"),
                           file_name="简历诊断分析.json", mime="application/json", width="stretch")
        st.caption("报告包含学生画像、岗位需求、技能与工具、职责与项目、成果线索、修改建议及岗位推荐。")
