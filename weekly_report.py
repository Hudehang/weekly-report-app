# -*- coding: utf-8 -*-
"""
多人周报填报系统（10人规模，免登录版）
任何人打开即可填写；填姓名 → 提交 / 修改本周周报；管理员密码查看全部 & 未填报统计 & 导出 Excel
运行：streamlit run weekly_report.py
"""
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO

# ============ 配置区（上线前改这里） ============
STAFF = ["张三", "李四", "王五", "赵六", "钱七",
         "孙八", "周九", "吴十", "郑十一", "王十二"]   # 改成你团队10人名字（用于未填报统计）
ADMIN_PASSWORD = "admin@2026"                          # 管理员密码
DB_FILE = "report.db"
# ================================================

def week_key(today: date = None):
    today = today or date.today()
    iso = today.isocalendar()
    monday = today - timedelta(days=today.weekday())
    return f"{iso[0]}-W{iso[1]:02d}", monday.isoformat()

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS week_report(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, week_key TEXT, week_start TEXT,
        work_this TEXT, work_next TEXT, risk TEXT, updated_at TEXT
    )''')
    conn.commit(); conn.close()

def upsert_report(name, wk, ws, wt, wn, risk):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT id FROM week_report WHERE name=? AND week_key=?", (name, wk))
    row = c.fetchone()
    now = datetime.now().isoformat(timespec="seconds")
    if row:
        c.execute("""UPDATE week_report SET work_this=?,work_next=?,risk=?,updated_at=? WHERE id=?""",
                  (wt, wn, risk, now, row[0]))
    else:
        c.execute("""INSERT INTO week_report(name,week_key,week_start,work_this,work_next,risk,updated_at)
                     VALUES(?,?,?,?,?,?,?)""", (name, wk, ws, wt, wn, risk, now))
    conn.commit(); conn.close()

def get_my_report(name, wk):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT work_this,work_next,risk FROM week_report WHERE name=? AND week_key=?", (name, wk))
    r = c.fetchone(); conn.close(); return r

def get_all_by_week(wk):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql("""SELECT name,work_this,work_next,risk,updated_at FROM week_report
                        WHERE week_key=? ORDER BY name""", conn, params=(wk,))
    conn.close(); return df

init_db()
wk, ws = week_key()

st.title("📝 周报填报系统（免登录）")
st.caption(f"本周：**{wk}**　周一日期：{ws}　直接填写即可，本周重复提交会覆盖已有内容。")

# 姓名输入（顶部固定）
name = st.text_input("你的姓名", placeholder="输入姓名后填写下方周报")

tab_fill, tab_admin = st.tabs(["📝 填报 / 修改周报", "📊 管理员面板"])

with tab_fill:
    if not name.strip():
        st.info("⬆ 先在上方输入你的姓名")
    else:
        existing = get_my_report(name.strip(), wk)
        default_this = existing[0] if existing else ""
        default_next = existing[1] if existing else ""
        default_risk = existing[2] if existing else ""
        if existing:
            st.info(f"已检测到你本周（{wk}）的周报，可直接修改后重新提交。")

        wt = st.text_area("本周完成工作", value=default_this, height=150, placeholder="可多条，每行一条")
        wn = st.text_area("下周工作计划", value=default_next, height=150)
        risk = st.text_area("存在问题 / 需要协调的风险", value=default_risk, height=100)

        if st.button("提交周报", type="primary"):
            if not (wt.strip() and wn.strip()):
                st.warning("本周工作和下周计划不能为空")
            else:
                upsert_report(name.strip(), wk, ws, wt.strip(), wn.strip(), risk.strip())
                st.success("✅ 周报已提交！本周内可随时回来修改，再次提交会覆盖。")

with tab_admin:
    pwd = st.text_input("管理员密码", type="password")
    if pwd != ADMIN_PASSWORD:
        if pwd: st.error("密码错误")
        st.stop()

    all_weeks = sqlite3.connect(DB_FILE).execute(
        "SELECT DISTINCT week_key FROM week_report ORDER BY week_key DESC").fetchall()
    week_options = [w[0] for w in all_weeks] or [wk]
    sel_week = st.selectbox("选择周次", week_options, index=0)

    df = get_all_by_week(sel_week)
    submitted = set(df["name"].tolist()) if not df.empty else set()
    missing = [n for n in STAFF if n not in submitted]

    c1, c2, c3 = st.columns(3)
    c1.metric("应提交", len(STAFF))
    c2.metric("已提交", len(submitted))
    c3.metric("未提交", len(missing))

    if missing:
        st.warning("未提交名单：" + "、".join(missing))
    else:
        st.success("🎉 本周全部提交完成")

    st.divider()
    st.subheader(f"{sel_week} 全部周报")
    if df.empty:
        st.info("本周还没有人提交")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name=sel_week)
            pd.DataFrame({"未提交人员": missing}).to_excel(xw, index=False, sheet_name="未提交名单")
        buf.seek(0)
        st.download_button("📥 导出本周周报 Excel", buf,
                           file_name=f"周报_{sel_week}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
