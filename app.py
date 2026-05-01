import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel 自动生成工具", layout="centered")

st.title("📦 Excel 自动化生成工具")
st.write("上传原始 Excel 文件（例如 `1.xlsx`），自动生成带 U 编号的导入模板。")

uploaded_file = st.file_uploader("请上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ 上传成功！预览数据：")
        st.dataframe(df.head())

        if {"Shipment", "Location", "NewCode"}.issubset(df.columns):
            result_rows = []
            for _, row in df.iterrows():
                shipment = str(row["Shipment"]).strip()
                location = str(row["Location"]).strip()
                try:
                    count = int(row["NewCode"])
                except:
                    continue

                for i in range(1, count + 1):
                    result_rows.append({
                        "服务*": location,
                        "店铺": "",
                        "客户订单号": shipment,
                        "地址库编码": "",
                        "收件人姓名*": "/",
                        "收件人公司": "",
                        "收件人地址一*": "/",
                        "收件人地址二": "",
                        "收件人地址三": "",
                        "收件人城市*": "/",
                        "收件人省份*": "",
                        "收件人邮编*": "/",
                        "收件人国家*": "CA",
                        "收件人电话*": "/",
                        "收件人邮箱": "",
                        "参考号一": "",
                        "参考号二": "",
                        "申报币种": "",
                        "备注": "",
                        "货箱编号": f"{shipment}-{str(i).zfill(4)}",
                        "货箱重量(KG)": 1,
                        "货箱长度(CM)": 1,
                        "货箱宽度(CM)": 1,
                        "货箱高度(CM)": 1,
                        "产品SKU": "",
                        "产品英文品名": "/",
                        "产品中文品名": "/",
                        "产品申报单价": 1,
                        "产品申报数量": 1,
                        "产品材质": "/",
                        "产品海关编码": "",
                        "产品用途": "",
                        "产品品牌": "",
                        "产品型号": "",
                        "产品销售链接": "",
                        "产品销售价格": "",
                        "产品图片链接": "",
                        "产品重量(kg)": "",
                        "产品ASIN": "",
                        "产品FNSKU": "",
                        "带电": "",
                        "带磁": "",
                        "危险品": "",
                        "报关方式": "",
                        "清关方式": "",
                        "交税方式": "",
                        "交货条款": "",
                        "派送方式": "",
                        "VAT号": "",
                        "购买保险": "",
                        "保价": "",
                        "投保币种": "",

                    })

            result_df = pd.DataFrame(result_rows)

            st.write("✅ 已生成结果预览：")
            st.dataframe(result_df.head(10))

            # 下载功能
            output = BytesIO()
            result_df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                label="📥 下载生成结果 Excel",
                data=output,
                file_name="生成结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.error("Excel 缺少必须的列：Shipment、Location、NewCode")

    except Exception as e:
        st.error(f"❌ 解析文件失败：{e}")
