# -*- coding: utf-8 -*-
"""
多人周报填报系统（免登录版）
- 填报：姓名 + 本周工作 + 下周计划 + 风险 + 本周指标（指标名/年度目标/完成值，差距自动算）
- 管理员：查看全部 / 编辑 / 删除单条 / 导出 Excel（含指标）+ Word（按部门周报版式）
运行：streamlit run weekly_report.py
"""
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

# ============ 配置区（上线前改这里） ============
ADMIN_PASSWORD = "admin@2026"
DEPT_NAME = "协同运营部"
# 固定指标：(序号, 目标指标, 单位, 年度目标)。首页直接展示，只填"完成值"。
INDICATORS = [
    (1,  "物资编码维度集中物资服务覆盖率", "%", 65),
    (2,  "生产物资电商化下单率", "%", 50),
    (3,  "集中供应生产物资即时供应率", "%", 95),
    (4,  "实现在线协同的关键供应商个数（含渠道商）", "家", 80),
    (5,  "累计投运中心仓数量", "个", 12),
    (6,  "集团公司发电企业单位库存金额最高值", "元/kW", 4.8),
    (7,  "实施生产物资仓储管理的电厂总数", "个", 105),
    (8,  "集中仓储模式覆盖区域内期末整体单位库存金额降低比例", "%", 10),
    (9,  "中心仓库存周转率", "次", 14),
    (10, "“一码到底”功能应用品类数", "个", 30),
    (11, "集中物流框架协议可视监管覆盖率（按物资采购金额占比统计）", "%", 50),
    (12, "中心仓新能源配送车辆占比", "%", 60),
    (13, "废旧物资集中处置成交金额", "万元", 35000),
    (14, "闲置物资调剂金额", "万元", 5000),
    (15, "在华能商城绿色低碳专区上架绿色商品项数", "项", 800),
    (16, "在华能商城自主可控专区上架自主可控信息化相关产品品类项数", "项", 800),
    (17, "品类物资需求预测准确率", "%", 65),
    (18, "集中供应物资抽检率（供应商及品类抽检数量与集中供应物资供应商及品类占比）", "%", 2),
]
DB_FILE = "report.db"
# ================================================

def week_key(today: date = None):
    today = today or date.today()
    iso = today.isocalendar()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return f"{iso[0]}-W{iso[1]:02d}", monday.isoformat(), sunday.isoformat()

def cn_week_num(wk):
    return int(wk.split("-W")[1])

def init_db():
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS week_report(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, week_key TEXT, week_start TEXT,
        work_this TEXT, work_next TEXT, risk TEXT, updated_at TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS indicator(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER,
        indicator_name TEXT,
        annual_target REAL,
        actual_value REAL
    )''')
    conn.commit(); conn.close()

def upsert_report(name, wk, ws, wt, wn, risk):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT id FROM week_report WHERE name=? AND week_key=?", (name, wk))
    row = c.fetchone()
    now = datetime.now().isoformat(timespec="seconds")
    if row:
        rid = row[0]
        c.execute("""UPDATE week_report SET work_this=?,work_next=?,risk=?,updated_at=? WHERE id=?""",
                  (wt, wn, risk, now, rid))
    else:
        c.execute("""INSERT INTO week_report(name,week_key,week_start,work_this,work_next,risk,updated_at)
                     VALUES(?,?,?,?,?,?,?)""", (name, wk, ws, wt, wn, risk, now))
        rid = c.lastrowid
    conn.commit(); conn.close()
    return rid

def replace_indicators(report_id, df_ind):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("DELETE FROM indicator WHERE report_id=?", (report_id,))
    for _, r in df_ind.iterrows():
        name = str(r.get("指标名称", "")).strip()
        if not name or name == "nan":
            continue
        try:
            target = float(r.get("年度目标", 0) or 0)
        except Exception:
            target = 0.0
        try:
            actual = float(r.get("完成值", 0) or 0)
        except Exception:
            actual = 0.0
        c.execute("""INSERT INTO indicator(report_id,indicator_name,annual_target,actual_value)
                     VALUES(?,?,?,?)""", (report_id, name, target, actual))
    conn.commit(); conn.close()

def get_my_report(name, wk):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT id,work_this,work_next,risk FROM week_report WHERE name=? AND week_key=?", (name, wk))
    r = c.fetchone(); conn.close(); return r

def get_indicators(report_id):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql("""SELECT indicator_name AS 指标名称,
                               annual_target AS 年度目标,
                               actual_value AS 完成值
                        FROM indicator WHERE report_id=?""",
                     conn, params=(report_id,))
    conn.close()
    return df if not df.empty else pd.DataFrame(columns=["指标名称", "年度目标", "完成值"])

def get_rows_by_week(wk):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql("""SELECT id,name,work_this,work_next,risk,updated_at
                        FROM week_report WHERE week_key=? ORDER BY name""",
                     conn, params=(wk,))
    conn.close(); return df

def get_row_by_id(rid):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT id,name,week_key,work_this,work_next,risk,updated_at FROM week_report WHERE id=?", (rid,))
    r = c.fetchone(); conn.close(); return r

def update_report_by_id(rid, name, wt, wn, risk):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    now = datetime.now().isoformat(timespec="seconds")
    c.execute("""UPDATE week_report SET name=?,work_this=?,work_next=?,risk=?,updated_at=? WHERE id=?""",
              (name, wt, wn, risk, now, rid))
    conn.commit(); conn.close()

def delete_report_by_id(rid):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("DELETE FROM indicator WHERE report_id=?", (rid,))
    c.execute("DELETE FROM week_report WHERE id=?", (rid,))
    conn.commit(); conn.close()

def gap_text(target, actual):
    """差距 = 年度目标 - 完成值；<=0 显示'已完成'"""
    try:
        gap = float(target) - float(actual)
    except Exception:
        return ""
    if gap <= 0:
        return "已完成"
    return f"{gap:.2f}"

# ---------- Word 导出（按部门周报版式） ----------
def build_docx(rows_with_ind, week_label, monday, sunday):
    """rows_with_ind: list of dict(name, work_this, work_next, risk, df_ind)"""
    doc = Document()
    # 默认字体
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(11)

    # 标题
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = h.add_run(f"{DEPT_NAME}{date.today().year}年第{cn_week_num(week_label)}周周报")
    run.bold = True
    run.font.size = Pt(16)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(f"（{monday}～{sunday}）").font.size = Pt(11)

    # 1. 本周指标完成情况
    doc.add_paragraph("1.本周指标完成情况").runs[0].bold = True
    doc.add_paragraph("一、指标完成情况")
    # 合并本周所有人填的完成值：按指标名取最新提交人的值
    merged = {}
    for r in rows_with_ind:
        df = r["df_ind"]
        if df.empty:
            continue
        for _, ir in df.iterrows():
            merged[ir["指标名称"]] = ir["完成值"]
    # 6 列：序号/目标指标/单位/年度目标/完成值/差距
    t = doc.add_table(rows=1, cols=6)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    for i, hname in enumerate(["序号", "目标指标", "单位", "年度目标", "完成值", "差距"]):
        hdr[i].text = hname
        for p in hdr[i].paragraphs:
            for run in p.runs:
                run.bold = True
    for sno, iname, unit, target in INDICATORS:
        actual = merged.get(iname, "")
        cells = t.add_row().cells
        cells[0].text = str(sno)
        cells[1].text = iname
        cells[2].text = unit
        cells[3].text = str(target)
        cells[4].text = str(actual) if actual != "" else ""
        cells[5].text = gap_text(target, actual) if actual != "" else ""

    # 2. 各板块重点工作进展与下周计划
    doc.add_paragraph("2.各板块重点工作进展与下周计划").runs[0].bold = True
    for r in rows_with_ind:
        p = doc.add_paragraph()
        p.add_run(f"【{r['name']}】").bold = True
        doc.add_paragraph(f"本周完成：{r['work_this'] or ''}")
        doc.add_paragraph(f"下周计划：{r['work_next'] or ''}")
        if r.get("risk"):
            doc.add_paragraph(f"问题与风险：{r['risk']}")
        doc.add_paragraph("")

    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf

# ============== 主流程 ==============
init_db()
wk, monday, sunday = week_key()

if "editing_id" not in st.session_state:
    st.session_state["editing_id"] = None

# 全局样式：表格铺满全宽、字体小一号、长字段自动折行
st.markdown("""
<style>
    .stApp { font-size: 13px; }
    /* 表格容器铺满宽度 */
    [data-testid="stDataFrameResizable"], [data-testid="stDataEditorResizable"] {
        width: 100% !important;
        overflow-x: hidden !important;
    }
    /* 单元格长文本折行 */
    .stDataEditor table, .stDataFrame table {
        width: 100% !important;
        table-layout: fixed !important;
    }
    .stDataEditor th, .stDataEditor td,
    .stDataFrame  th, .stDataFrame  td {
        white-space: normal !important;
        word-break: break-word !important;
        overflow-wrap: anywhere !important;
        font-size: 12px !important;
        padding: 3px 6px !important;
        line-height: 1.3 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📝 周报填报系统（免登录）")
st.caption(f"本周：**{wk}**（{monday}～{sunday}）　打开即填，本周重复提交会覆盖已有内容。")

name = st.text_input("你的姓名", placeholder="输入姓名后填写下方周报")

tab_fill, tab_admin = st.tabs(["📝 填报 / 修改周报", "📊 管理员面板"])

# ---------- 填报页 ----------
with tab_fill:
    existing = get_my_report(name.strip(), wk) if name.strip() else None
    rid = existing[0] if existing else None
    default_this = existing[1] if existing else ""
    default_next = existing[2] if existing else ""
    default_risk = existing[3] if existing else ""
    if existing:
        st.info(f"已检测到你本周（{wk}）的周报，可直接修改后重新提交。")

    # 指标表（在本周完成工作上方）
    st.markdown("**本周指标完成情况**")
    saved = get_indicators(rid).set_index("指标名称")["完成值"].to_dict() if rid else {}
    rows = []
    for sno, iname, unit, target in INDICATORS:
        actual = float(saved.get(iname, 0) or 0)
        rows.append({
            "序号": sno, "目标指标": iname, "单位": unit,
            "年度目标": float(target), "完成值": actual,
            "差距": gap_text(target, actual),
        })
    df_ind = st.data_editor(
        pd.DataFrame(rows), num_rows="fixed", use_container_width=True, hide_index=True,
        disabled=["序号", "目标指标", "单位", "年度目标", "差距"],
        column_config={
            "序号": st.column_config.NumberColumn(width="small"),
            "目标指标": st.column_config.TextColumn(width="large"),
            "单位": st.column_config.TextColumn(width="small"),
            "年度目标": st.column_config.NumberColumn(format="%.2f"),
            "完成值": st.column_config.NumberColumn(format="%.2f"),
            "差距": st.column_config.TextColumn(width="small"),
        },
        key="ind_editor"
    )

    wt = st.text_area("本周完成工作", value=default_this, height=130, placeholder="可多条，每行一条")
    wn = st.text_area("下周工作计划", value=default_next, height=130)
    risk = st.text_area("存在问题 / 需要协调的风险", value=default_risk, height=80)

    if st.button("提交周报", type="primary"):
        if not name.strip():
            st.warning("⚠️ 姓名为必填项，请先填写姓名")
        elif not (wt.strip() and wn.strip()):
            st.warning("本周工作和下周计划不能为空")
        else:
            new_rid = upsert_report(name.strip(), wk, monday,
                                   wt.strip(), wn.strip(), risk.strip())
            replace_indicators(new_rid, df_ind)
            st.success("✅ 周报已提交！本周内可随时回来修改，再次提交会覆盖。")

# ---------- 管理员面板 ----------
with tab_admin:
    pwd = st.text_input("管理员密码", type="password")
    if pwd != ADMIN_PASSWORD:
        if pwd: st.error("密码错误")
        st.stop()

    all_weeks = sqlite3.connect(DB_FILE).execute(
        "SELECT DISTINCT week_key FROM week_report ORDER BY week_key DESC").fetchall()
    week_options = [w[0] for w in all_weeks] or [wk]
    sel_week = st.selectbox("选择周次", week_options, index=0)

    df = get_rows_by_week(sel_week)
    submitted_names = sorted(df["name"].tolist()) if not df.empty else []

    c1, c2 = st.columns(2)
    c1.metric("本周已提交人数", len(submitted_names))
    c2.metric("本周记录条数", len(df))
    if submitted_names:
        st.success("已提交人员：" + "、".join(submitted_names))
    else:
        st.info("本周还没有人提交")

    # 编辑表单
    eid = st.session_state["editing_id"]
    if eid is not None:
        row = get_row_by_id(eid)
        if row is None:
            st.session_state["editing_id"] = None
        else:
            rid, rname, rwk, rwt, rwn, rrisk, rupd = row
            st.warning(f"正在编辑：{rname}（{rwk}）　原更新时间：{rupd}")
            e_name = st.text_input("姓名", value=rname, key=f"ename_{rid}")
            e_wt = st.text_area("本周完成工作", value=rwt or "", height=110, key=f"ewt_{rid}")
            e_wn = st.text_area("下周计划", value=rwn or "", height=110, key=f"ewn_{rid}")
            e_risk = st.text_area("问题与风险", value=rrisk or "", height=70, key=f"erisk_{rid}")
            saved_ind = get_indicators(rid).set_index("指标名称")["完成值"].to_dict()
            rows = [{"序号": sno, "目标指标": n, "单位": u, "年度目标": float(t),
                     "完成值": float(saved_ind.get(n, 0) or 0)}
                    for sno, n, u, t in INDICATORS]
            e_ind = st.data_editor(
                pd.DataFrame(rows), num_rows="fixed", use_container_width=True, hide_index=True,
                disabled=["序号", "目标指标", "单位", "年度目标"],
                key=f"eind_{rid}")
            b1, b2 = st.columns(2)
            if b1.button("💾 保存修改", key=f"save_{rid}", type="primary"):
                update_report_by_id(rid, e_name.strip(), e_wt.strip(), e_wn.strip(), e_risk.strip())
                replace_indicators(rid, e_ind)
                st.success("✅ 已保存修改")
                st.session_state["editing_id"] = None
                st.rerun()
            if b2.button("取消", key=f"cancel_{rid}"):
                st.session_state["editing_id"] = None
                st.rerun()
        st.divider()

    # 逐条记录
    st.subheader(f"{sel_week} 全部周报（{len(df)} 条）")
    if df.empty:
        st.info("本周还没有人提交")
    else:
        rows_with_ind = []
        for _, r in df.iterrows():
            rid = r["id"]
            df_ind = get_indicators(rid)
            rows_with_ind.append({
                "name": r["name"], "work_this": r["work_this"],
                "work_next": r["work_next"], "risk": r["risk"], "df_ind": df_ind
            })
            with st.container(border=True):
                cols = st.columns([3, 1, 1])
                cols[0].write(f"**{r['name']}**　·　更新于 {r['updated_at']}")
                cols[0].caption(f"本周：{(r['work_this'] or '')[:50]}　|　下周：{(r['work_next'] or '')[:40]}")
                if not df_ind.empty:
                    with cols[0].expander(f"查看指标（{len(df_ind)} 条）"):
                        preview = df_ind.copy()
                        preview["差距"] = preview.apply(lambda x: gap_text(x["年度目标"], x["完成值"]), axis=1)
                        st.dataframe(preview, use_container_width=True, hide_index=True)
                if cols[1].button("✏️ 编辑", key=f"edit_{rid}"):
                    st.session_state["editing_id"] = rid
                    st.rerun()
                if st.session_state.get(f"del_confirm_{rid}", False):
                    if cols[2].button("⚠️ 确认删除", key=f"del_yes_{rid}", type="primary"):
                        delete_report_by_id(rid)
                        st.warning(f"已删除 {r['name']} 的周报")
                        st.session_state[f"del_confirm_{rid}"] = False
                        st.rerun()
                else:
                    if cols[2].button("🗑 删除", key=f"del_{rid}"):
                        st.session_state[f"del_confirm_{rid}"] = True
                        st.rerun()

        st.divider()
        # 导出
        df_main = df[["name", "work_this", "work_next", "risk", "updated_at"]].rename(columns={
            "name": "姓名", "work_this": "本周完成", "work_next": "下周计划",
            "risk": "问题与风险", "updated_at": "更新时间"})
        # 指标 sheet：18 行固定骨架，合并本周所有人填的完成值
        merged = {}
        for r in rows_with_ind:
            for _, ir in r["df_ind"].iterrows():
                merged[ir["指标名称"]] = ir["完成值"]
        ind_rows = []
        for sno, iname, unit, target in INDICATORS:
            actual = merged.get(iname, "")
            ind_rows.append({
                "序号": sno, "目标指标": iname, "单位": unit,
                "年度目标": target, "完成值": actual,
                "差距": gap_text(target, actual) if actual != "" else ""
            })
        df_ind_all = pd.DataFrame(ind_rows)

        buf_x = BytesIO()
        with pd.ExcelWriter(buf_x, engine="openpyxl") as xw:
            df_main.to_excel(xw, index=False, sheet_name="周报内容")
            df_ind_all.to_excel(xw, index=False, sheet_name="指标完成情况")
        buf_x.seek(0)

        buf_w = build_docx(rows_with_ind, sel_week, monday, sunday)

        c1, c2 = st.columns(2)
        c1.download_button("📥 导出 Excel（含指标）", buf_x,
                          file_name=f"周报_{sel_week}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c2.download_button("📥 导出 Word", buf_w,
                           file_name=f"周报_{sel_week}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
