import streamlit as st
import pandas as pd
import openpyxl
import io
import os

# # 设置页面标题和图标
# st.set_page_config(page_title="A406表审核", page_icon="📊")
# st.title("A406表审核小程序网页版")
# st.caption("国家统计局铜仁调查队")

st.set_page_config(page_title="A406表审核", page_icon="📊", layout="wide")

st.title("A406表审核小程序网页版")
st.caption("国家统计局铜仁调查队")

# 显示品牌信息（与原 tkinter 界面中的背景文字和图标对应）
st.markdown("欢迎使用A406表审核小程序网页版！")
# st.image("统计标识.png", width=100)  # 如果你有图标图片，放在同一目录下

def process_files(uploaded_file):
    """
    处理上传的 Excel 文件，返回处理后的 Excel 文件的 bytes 对象
    """
    # 读取上传的 Excel 文件
    df = pd.read_excel(uploaded_file, sheet_name='Sheet1', engine='openpyxl')

    # 创建一个内存缓冲区，用于保存结果
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        # ------------------- 猪数据处理 -------------------
        pig_jin_weight = pd.to_numeric(df['猪畜禽产品产量-净重（公斤）'], errors='coerce')
        pig_quantity = pd.to_numeric(df['猪畜禽出栏量（头/只）'], errors='coerce').replace(0, pd.NA)
        df['猪胴体重'] = pig_jin_weight / pig_quantity

        # 异常筛选
        pig_filtered = df[(df['猪胴体重'] <= 85) | (df['猪胴体重'] >= 115) & (df['猪胴体重'].notna())]
        if not pig_filtered.empty:
            pig_filtered = pig_filtered[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                         '行政区划代码', '期别(年)', '状态', '猪畜禽存栏量（头/只）',
                                         '猪畜禽出栏量（头/只）', '猪畜禽产品产量-毛重（公斤）',
                                         '猪畜禽产品产量-净重（公斤）', '猪胴体重']]
            pd.DataFrame(['猪胴体重异常数据（胴体重<=85 或 胴体重>=115）']).to_excel(
                writer, sheet_name='猪胴体重异常', index=False, header=False, startrow=0
            )
            pig_filtered.to_excel(writer, sheet_name='猪胴体重异常', index=False, startrow=1)

        # 重复查找
        df_clean = df.dropna(subset=['猪胴体重', '状态']).copy()
        if len(df_clean['猪胴体重']) != len(set(df_clean['猪胴体重'])):
            df_clean['count'] = df_clean.groupby(['乡(镇、街道)', '猪胴体重'])['猪胴体重'].transform('size')
            result = df_clean[df_clean['count'] >= 3]
            if not result.empty:
                result = result[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                 '行政区划代码', '期别(年)', '状态', '猪畜禽存栏量（头/只）',
                                 '猪畜禽出栏量（头/只）', '猪畜禽产品产量-毛重（公斤）',
                                 '猪畜禽产品产量-净重（公斤）', '猪胴体重']]
                pd.DataFrame(['猪胴体重一致(>=3次)数据']).to_excel(
                    writer, sheet_name='猪胴体重一致', index=False, header=False, startrow=0
                )
                result.to_excel(writer, sheet_name='猪胴体重一致', index=False, startrow=1)

        # ------------------- 牛数据处理 -------------------
        cow_jin_weight = pd.to_numeric(df['牛畜禽产品产量-净重（公斤）'], errors='coerce')
        cow_quantity = pd.to_numeric(df['牛畜禽出栏量（头/只）'], errors='coerce').replace(0, pd.NA)
        df['牛胴体重'] = cow_jin_weight / cow_quantity

        cow_filtered = df[(df['牛胴体重'] <= 100)]
        if len(cow_filtered) == 0:
            cow_filtered = df[(df['牛胴体重'] <= 110) | (df['牛胴体重'] >= 170) & (df['牛胴体重'].notna())]
        else:
            cow_filtered = df[(df['牛胴体重'] <= 100) | (df['牛胴体重'] >= 170) & (df['牛胴体重'].notna())]

        if not cow_filtered.empty:
            cow_filtered = cow_filtered[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                         '行政区划代码', '期别(年)', '状态', '牛畜禽存栏量（头/只）',
                                         '牛畜禽出栏量（头/只）', '牛畜禽产品产量-毛重（公斤）',
                                         '牛畜禽产品产量-净重（公斤）', '牛胴体重']]
            pd.DataFrame(['牛胴体重异常数据（胴体重<=100 或 胴体重>=170，若胴体重<=100无数据，则筛选范围为胴体重<=110 或 胴体重>=170）']).to_excel(
                writer, sheet_name='牛胴体重异常', index=False, header=False, startrow=0
            )
            cow_filtered.to_excel(writer, sheet_name='牛胴体重异常', index=False, startrow=1)

        # 重复查找
        df_clean = df.dropna(subset=['牛胴体重', '状态']).copy()
        if len(df_clean['牛胴体重']) != len(set(df_clean['牛胴体重'])):
            df_clean['count'] = df_clean.groupby(['乡(镇、街道)', '牛胴体重'])['牛胴体重'].transform('size')
            result = df_clean[df_clean['count'] >= 3]
            if not result.empty:
                result = result[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                 '行政区划代码', '期别(年)', '状态', '牛畜禽存栏量（头/只）',
                                 '牛畜禽出栏量（头/只）', '牛畜禽产品产量-毛重（公斤）',
                                 '牛畜禽产品产量-净重（公斤）', '牛胴体重']]
                pd.DataFrame(['牛胴体重一致(>=3次)数据']).to_excel(
                    writer, sheet_name='牛胴体重一致', index=False, header=False, startrow=0
                )
                result.to_excel(writer, sheet_name='牛胴体重一致', index=False, startrow=1)

        # ------------------- 羊数据处理 -------------------
        sheep_jin_weight = pd.to_numeric(df['羊畜禽产品产量-净重（公斤）'], errors='coerce')
        sheep_quantity = pd.to_numeric(df['羊畜禽出栏量（头/只）'], errors='coerce').replace(0, pd.NA)
        df['羊胴体重'] = sheep_jin_weight / sheep_quantity

        sheep_filtered = df[(df['羊胴体重'] <= 15) | (df['羊胴体重'] >= 23) & (df['羊胴体重'].notna())]
        if not sheep_filtered.empty:
            sheep_filtered = sheep_filtered[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                             '行政区划代码', '期别(年)', '状态', '羊畜禽存栏量（头/只）',
                                             '羊畜禽出栏量（头/只）', '羊畜禽产品产量-毛重（公斤）',
                                             '羊畜禽产品产量-净重（公斤）', '羊胴体重']]
            pd.DataFrame(['羊胴体重异常数据（胴体重<=15 或 胴体重>=23）']).to_excel(
                writer, sheet_name='羊胴体重异常', index=False, header=False, startrow=0
            )
            sheep_filtered.to_excel(writer, sheet_name='羊胴体重异常', index=False, startrow=1)

        df_clean = df.dropna(subset=['羊胴体重', '状态']).copy()
        if len(df_clean['羊胴体重']) != len(set(df_clean['羊胴体重'])):
            df_clean['count'] = df_clean.groupby(['乡(镇、街道)', '羊胴体重'])['羊胴体重'].transform('size')
            result = df_clean[df_clean['count'] >= 3]
            if not result.empty:
                result = result[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                 '行政区划代码', '期别(年)', '状态', '羊畜禽存栏量（头/只）',
                                 '羊畜禽出栏量（头/只）', '羊畜禽产品产量-毛重（公斤）',
                                 '羊畜禽产品产量-净重（公斤）', '羊胴体重']]
                pd.DataFrame(['羊胴体重一致(>=3次)数据']).to_excel(
                    writer, sheet_name='羊胴体重一致', index=False, header=False, startrow=0
                )
                result.to_excel(writer, sheet_name='羊胴体重一致', index=False, startrow=1)

        # ------------------- 活家禽数据处理 -------------------
        qin_jin_weight = pd.to_numeric(df['活家禽畜禽产品产量-净重（公斤）'], errors='coerce')
        qin_quantity = pd.to_numeric(df['活家禽畜禽出栏量（头/只）'], errors='coerce').replace(0, pd.NA)
        df['活家禽胴体重'] = qin_jin_weight / qin_quantity

        qin_filtered = df[(df['活家禽胴体重'] <= 1.5) | (df['活家禽胴体重'] >= 2.2) & (df['活家禽胴体重'].notna())]
        if not qin_filtered.empty:
            qin_filtered = qin_filtered[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                         '行政区划代码', '期别(年)', '状态', '活家禽畜禽存栏量（头/只）',
                                         '活家禽畜禽出栏量（头/只）', '活家禽畜禽产品产量-毛重（公斤）',
                                         '活家禽畜禽产品产量-净重（公斤）', '活家禽胴体重']]
            pd.DataFrame(['活家禽胴体重异常数据（胴体重<=1.5 或 胴体重>=2.2）']).to_excel(
                writer, sheet_name='活家禽胴体重异常', index=False, header=False, startrow=0
            )
            qin_filtered.to_excel(writer, sheet_name='活家禽胴体重异常', index=False, startrow=1)

        df_clean = df.dropna(subset=['活家禽胴体重', '状态']).copy()
        if len(df_clean['活家禽胴体重']) != len(set(df_clean['活家禽胴体重'])):
            df_clean['count'] = df_clean.groupby(['乡(镇、街道)', '活家禽胴体重'])['活家禽胴体重'].transform('size')
            result = df_clean[df_clean['count'] >= 3]
            if not result.empty:
                result = result[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                 '行政区划代码', '期别(年)', '状态', '活家禽畜禽存栏量（头/只）',
                                 '活家禽畜禽出栏量（头/只）', '活家禽畜禽产品产量-毛重（公斤）',
                                 '活家禽畜禽产品产量-净重（公斤）', '活家禽胴体重']]
                pd.DataFrame(['活家禽胴体重一致(>=3次)数据']).to_excel(
                    writer, sheet_name='活家禽胴体重一致', index=False, header=False, startrow=0
                )
                result.to_excel(writer, sheet_name='活家禽胴体重一致', index=False, startrow=1)

        # ------------------- 鸡蛋产量除以鸡存栏 -------------------
        egg1_weight = pd.to_numeric(df['鸡蛋畜禽产品产量-净重（公斤）'], errors='coerce')
        chicken_quantity = pd.to_numeric(df['鸡畜禽存栏量（头/只）'], errors='coerce').replace(0, pd.NA)
        df['鸡蛋产量除以鸡存栏'] = egg1_weight / chicken_quantity
        ji1_filtered = df[(df['鸡蛋产量除以鸡存栏'] >= 3.2) & (df['鸡蛋产量除以鸡存栏'].notna())]
        if not ji1_filtered.empty:
            ji1_filtered = ji1_filtered[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                         '行政区划代码', '期别(年)', '状态', '鸡畜禽存栏量（头/只）',
                                         '鸡畜禽出栏量（头/只）', '鸡畜禽产品产量-毛重（公斤）',
                                         '鸡畜禽产品产量-净重（公斤）', '鸡蛋畜禽存栏量（头/只）',
                                         '鸡蛋畜禽出栏量（头/只）', '鸡蛋畜禽产品产量-毛重（公斤）',
                                         '鸡蛋畜禽产品产量-净重（公斤）', '鸡蛋产量除以鸡存栏']]
            pd.DataFrame(['鸡蛋产量除以鸡存栏异常数据(>= 3.2)']).to_excel(
                writer, sheet_name='鸡蛋产量除以鸡存栏异常', index=False, header=False, startrow=0
            )
            ji1_filtered.to_excel(writer, sheet_name='鸡蛋产量除以鸡存栏异常', index=False, startrow=1)

        # ------------------- 鸡蛋产量除以蛋鸡存栏 -------------------
        egg2_weight = pd.to_numeric(df['鸡蛋畜禽产品产量-净重（公斤）'], errors='coerce')
        egg_chicken_quantity = pd.to_numeric(df['蛋鸡畜禽存栏量（头/只）'], errors='coerce').replace(0, pd.NA)
        df['鸡蛋产量除以蛋鸡存栏'] = egg2_weight / egg_chicken_quantity
        ji2_filtered = df[(df['鸡蛋产量除以蛋鸡存栏'] >= 4) & (df['鸡蛋产量除以蛋鸡存栏'].notna())]
        if not ji2_filtered.empty:
            ji2_filtered = ji2_filtered[['序号', '市(州)', '县(区、市)', '乡(镇、街道)', '村(据、社区)',
                                         '行政区划代码', '期别(年)', '状态', '蛋鸡畜禽存栏量（头/只）',
                                         '蛋鸡畜禽出栏量（头/只）', '蛋鸡畜禽产品产量-毛重（公斤）',
                                         '蛋鸡畜禽产品产量-净重（公斤）', '鸡蛋畜禽存栏量（头/只）',
                                         '鸡蛋畜禽出栏量（头/只）', '鸡蛋畜禽产品产量-毛重（公斤）',
                                         '鸡蛋畜禽产品产量-净重（公斤）', '鸡蛋产量除以蛋鸡存栏']]
            pd.DataFrame(['鸡蛋产量除以蛋鸡存栏异常数据(>= 4)']).to_excel(
                writer, sheet_name='鸡蛋产量除以蛋鸡存栏异常', index=False, header=False, startrow=0
            )
            ji2_filtered.to_excel(writer, sheet_name='鸡蛋产量除以蛋鸡存栏异常', index=False, startrow=1)

    # 将缓冲区指针移到开头
    output_buffer.seek(0)
    return output_buffer.getvalue()


# ---------- Streamlit 界面 ----------
uploaded_file = st.file_uploader("请选择要处理的 Excel 文件", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # 显示文件名
    st.write(f"已上传：{uploaded_file.name}")

    if st.button("开始审核"):
        with st.spinner("正在处理数据，请稍候..."):
            try:
                result_bytes = process_files(uploaded_file)

                # 提供下载按钮
                st.download_button(
                    label="下载审核结果",
                    data=result_bytes,
                    file_name="A406审核结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("处理完成！点击上方按钮下载结果。")
            except Exception as e:
                st.error(f"处理过程中发生错误：{e}")
                st.stop()