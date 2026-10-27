import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os
import shutil
from io import BytesIO
import zipfile
import json
from datetime import datetime

# ====================== 配置常量（Cert写入位置已配置完成） ======================
APP_TITLE = "EHP MainCon Claim 自动计算工具"
CERT_TPL_FOLDER = "cert_templates"
SNAPSHOT_FOLDER = "snapshots"

# 三份Cert模板与回填单元格C28
CERT_MAP = {
    "main": {
        "tpl_name": "PC26127- Equatorial Hotel Penang Claim 9.xlsx",
        "target_cell": "C28",
        "desc": "主客房部分 PC26127"
    },
    "corridor": {
        "tpl_name": "PC26128- Equatorial Hotel Penang (Guestroom Corridor) Claim 6.xlsx",
        "target_cell": "C28",
        "desc": "走廊部分 PC26128"
    },
    "public": {
        "tpl_name": "PC26129- Equatorial Hotel Penang (Public Area) Claim 3.xlsx",
        "target_cell": "C28",
        "desc": "公共区域 PC26129"
    }
}

# 输出sheet名称，和原Excel保持一致
OUTPUT_SHEETS = [
    "MC",
    "MC(Foyer)",
    "MC‑Corridor",
    "Public Area Cement Partition",
    "MainCon_Claim_Summary"
]

# 创建文件夹
os.makedirs(CERT_TPL_FOLDER, exist_ok=True)
os.makedirs(SNAPSHOT_FOLDER, exist_ok=True)

# ====================== 页面初始化 ======================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title("🏗️ " + APP_TITLE)
st.markdown("""
> 📌工作流：
> 1. 上传标准化月度进度Excel（Level / Wing / Room_Type / Item_Work / Percent_Complete）
> 2. 系统复刻原MC Calculation全部计算逻辑
> 3. 网页预览结果，支持手动调整申报%（Override）
> 4. 导出完整MC结果Excel + **自动回填3份Payment Cert(C28)**
> 5. 自动生成本月快照，用于下个月计算本期申报，避免Over‑claim
""")

# ====================== 工具函数 ======================
def read_upload_excel(upload_file):
    """读取上传的进度输入表"""
    df = pd.read_excel(upload_file)
    must_cols = ["Level","Wing","Room_Type","Item_Work","Percent_Complete","Remarks"]
    for c in must_cols:
        if c not in df.columns:
            raise Exception(f"输入模板缺失列：{c}，请使用input_progress_template.xlsx")
    df["Percent_Complete"] = pd.to_numeric(df["Percent_Complete"], errors="coerce")
    df = df[(df["Percent_Complete"] >=0) & (df["Percent_Complete"] <=100)].copy()
    return df

def mock_mc_calc_engine(df_input):
    """
    模拟原Excel MC Calculation引擎
    注意：此处为框架，你后续把真实房间单价、Corner Room特殊系数填进base_master_data字典；
    当前输出结构对齐原4个sheet + Summary；
    返回各模块累计Workdone金额：main_total, foyer_total, corridor_total, public_total
    """
    # ==========【这里需要填入你的真实主数据：房间类型、房间数量、单房Rate，公共区域L/H等】==========
    base_master_data = {
        "room_rates": {
            "SK1": {"qty":38, "rate":1400},
            "SK2": {"qty":194, "rate":1400},
            "SK3": {"qty":13, "rate":1400},
            "ST1": {"qty":72, "rate":1400},
            "ST2": {"qty":14, "rate":1400},
            "CK": {"qty":29, "rate":1900},
            "CDD": {"qty":21, "rate":2500},
            "CT": {"qty":8, "rate":2000},
            "PR1": {"qty":5, "rate":4900},
            "PR2": {"qty":8, "rate":3500},
            "FS": {"qty":4, "rate":4700}
        }
    }

    # -----------------客房MC计算逻辑演示框架----------------
    room_summary_list = []
    for rt, info in base_master_data["room_rates"].items():
        sub = df_input[df_input["Room_Type"]==rt].copy()
        avg_pct = sub["Percent_Complete"].mean() if len(sub)>0 else 0
        contract_amt = info["qty"] * info["rate"]
        claim_amt = contract_amt * avg_pct / 100
        room_summary_list.append({
            "Room_Type": rt,
            "Qty": info["qty"],
            "Rate_RM": info["rate"],
            "Avg_Percent": round(avg_pct,2),
            "Contract_Amount": contract_amt,
            "Cumulative_Claim": round(claim_amt,2)
        })
    df_mc = pd.DataFrame(room_summary_list)
    main_total = df_mc["Cumulative_Claim"].sum()

    # Foyer，走廊，公共区域演示占位（后续把你原Excel的RoundUP、L×H等公式1:1移植）
    foyer_total = main_total * 0.08
    corridor_total = main_total * 0.22
    public_total = main_total * 0.15

    # 汇总表 MainCon_Claim_Summary
    sum_dict = {
        "Item":["GuestRoom(MC)","Foyer","Corridor","Public Area"],
        "Cumulative_Workdone":[main_total, foyer_total, corridor_total, public_total]
    }
    df_summary = pd.DataFrame(sum_dict)
    total_all = df_summary["Cumulative_Workdone"].sum()
    df_summary.loc[len(df_summary)] = ["GRAND TOTAL", total_all]

    return {
        "df_mc": df_mc,
        "foyer_total": foyer_total,
        "corridor_total": corridor_total,
        "public_total": public_total,
        "df_summary": df_summary,
        "grand_total": total_all
    }

def write_cert_template(tpl_filename, target_cell, workdone_value):
    """读取cert模板，写入C28单元格，返回BytesIO对象"""
    tpl_path = os.path.join(CERT_TPL_FOLDER, tpl_filename)
    if not os.path.exists(tpl_path):
        raise FileNotFoundError(f"找不到模板文件：{tpl_path}，请放到cert_templates文件夹")
    wb = load_workbook(tpl_path)
    ws = wb.active
    ws[target_cell] = round(workdone_value,2)
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def export_mc_excel(calc_result):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        calc_result["df_mc"].to_excel(writer, sheet_name="MC", index=False)
        pd.DataFrame([{"Foyer_Cumulative":calc_result["foyer_total"]}]).to_excel(writer, sheet_name="MC(Foyer)", index=False)
        pd.DataFrame([{"Corridor_Cumulative":calc_result["corridor_total"]}]).to_excel(writer, sheet_name="MC‑Corridor", index=False)
        pd.DataFrame([{"PublicArea_Cumulative":calc_result["public_total"]}]).to_excel(writer, sheet_name="Public Area Cement Partition", index=False)
        calc_result["df_summary"].to_excel(writer, sheet_name="MainCon_Claim_Summary", index=False)
    bio.seek(0)
    return bio

def save_snapshot(calc_result, claim_no):
    snap = {
        "claim_no": claim_no,
        "calc":{
            "main": calc_result["df_mc"].to_dict("records"),
            "foyer_total":calc_result["foyer_total"],
            "corridor_total":calc_result["corridor_total"],
            "public_total":calc_result["public_total"],
            "grand_total":calc_result["grand_total"]
        },
        "timestamp": datetime.now().isoformat()
    }
    fname = os.path.join(SNAPSHOT_FOLDER, f"snap_{claim_no}.json")
    with open(fname,"w",encoding="utf‑8") as f:
        json.dump(snap, f, ensure_ascii=False, indent=2)
    return fname

# ====================== 网页主体交互 ======================
st.subheader("①上传月度进度输入Excel")
uploaded_file = st.file_uploader("上传 input_progress_template.xlsx 格式文件", type=["xlsx"])
claim_no_input = st.text_input("本期Claim编号（用于快照命名，例如：Claim‑9）", value="Claim‑9")

if uploaded_file is not None:
    try:
        df_input = read_upload_excel(uploaded_file)
        st.success("✅输入文件读取成功")
        st.dataframe(df_input, height=300)

        # 运行计算引擎
        calc_res = mock_mc_calc_engine(df_input)

        st.subheader("②计算结果预览")
        col1,col2,col3,col4 = st.columns(4)
        col1.metric("客房MC累计(RM)", f'{calc_res["df_mc"]["Cumulative_Claim"].sum():,.2f}')
        col2.metric("Foyer累计(RM)", f'{calc_res["foyer_total"]:,.2f}')
        col3.metric("Corridor走廊累计(RM)", f'{calc_res["corridor_total"]:,.2f}')
        col4.metric("Public公共区域累计(RM)", f'{calc_res["public_total"]:,.2f}')

        st.markdown("##### 客房明细 MC sheet")
        st.dataframe(calc_res["df_mc"])
        st.markdown("##### MainCon 汇总表 MainCon_Claim_Summary")
        st.dataframe(calc_res["df_summary"])

        st.info("💡注意：当前mock_mc_calc_engine为演示框架，**你需要把原Excel真实公式、Corner Room系数、公共区域L/H、RoundUP逻辑补全到此函数内，输出数字才会和你手工Excel完全一致**")

        st.subheader("③导出全部文件包（MC结果Excel + 3份回填C28的Cert）")
        if st.button("🔨生成全部输出包ZIP"):
            # 1 导出MC结果excel
            mc_bio = export_mc_excel(calc_res)
            # 2 分别写三份cert
            bio_main = write_cert_template(CERT_MAP["main"]["tpl_name"], CERT_MAP["main"]["target_cell"], calc_res["df_mc"]["Cumulative_Claim"].sum())
            bio_corridor = write_cert_template(CERT_MAP["corridor"]["tpl_name"], CERT_MAP["corridor"]["target_cell"], calc_res["corridor_total"])
            bio_public = write_cert_template(CERT_MAP["public"]["tpl_name"], CERT_MAP["public"]["target_cell"], calc_res["public_total"])
            # 3 保存快照
            snap_path = save_snapshot(calc_res, claim_no_input)
            st.success(f"快照已保存：{snap_path}")

            # 打包zip
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("EHP_MainCon_Result.xlsx", mc_bio.getvalue())
                zf.writestr("PC26127_已回填C28.xlsx", bio_main.getvalue())
                zf.writestr("PC26128_已回填C28.xlsx", bio_corridor.getvalue())
                zf.writestr("PC26129_已回填C28.xlsx", bio_public.getvalue())
            zip_buffer.seek(0)
            st.download_button(
                label="📥下载全部输出ZIP包",
                data=zip_buffer,
                file_name=f"EHP_Output_{claim_no_input}.zip",
                mime="application/zip"
            )
            st.success("✅生成完成！解压zip得到4个文件，三份Cert的C28已自动填入对应模块Workdone金额")

    except Exception as e:
        st.error(f"❌处理异常：{str(e)}")
        st.exception(e)

st.markdown("---")
st.subheader("📋使用注意事项")
st.markdown("""
1. 三份原始Cert模板务必放在子文件夹 `cert_templates`，文件名一字不差，否则程序找不到模板；
2. **mock_mc_calc_engine函数是框架，现在里面是模拟演示数字！！**
> ⚠重要：必须把你原EHP ROOM_2026_10.xlsx的真实计算规则：SAME/DIFFERENT/Corner Room特殊系数、走廊L×H向上取整RoundUP、公共区域计算全部移植进`mock_mc_calc_engine`函数，输出数字才和你手工Excel一模一样；
3. 每月快照保存在`snapshots`文件夹json，后续扩展：读取上月快照自动计算【本期申报=本月累计‑上期已申报】；
4. 不要打开正在读写的Excel文件，会造成占用报错；全部关闭Excel再上传/导出。
""")

