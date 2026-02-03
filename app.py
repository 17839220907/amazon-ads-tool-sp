import streamlit as st
import pandas as pd
import io
import datetime

# ================= 1. 网页基础设置 =================
st.set_page_config(page_title="亚马逊SP广告自动生成工具", page_icon="🚀", layout="wide")

st.title("🚀 亚马逊SP广告自动生成工具")
st.markdown("""
**使用说明：**
1. 准备好你的 Excel 配置文件（格式需与 **广告自动生成工具.xlsx** 一致）。
2. 点击下方按钮上传文件。
3. 系统将自动处理并提供 CSV 下载。
""")


# ================= 2. 核心工具函数 (完全复用原逻辑) =================

def clean_df(df):
    """清洗DataFrame：去列名空格、去空行"""
    if df is None: return None
    df.columns = [str(c).strip().replace('\ufeff', '') for c in df.columns]
    df.dropna(how='all', inplace=True)
    return df


def find_sheet_strict(xls, target_name):
    """
    【核心修正】精准查找 Sheet
    逻辑：必须完全相等（忽略大小写和首尾空格）
    绝不允许 "Pro" 匹配到 "Pro Max"
    """
    target = target_name.lower().strip()
    for sheet in xls.sheet_names:
        # 只有当名字完全一样时才返回
        if sheet.lower().strip() == target:
            return sheet
    return None


def get_col(row, col_list):
    """获取数值"""
    for col in col_list:
        if col in row and pd.notna(row[col]):
            return float(row[col])
    return 0.0


def get_str(row, col_list):
    """获取字符串"""
    for col in col_list:
        if col in row and pd.notna(row[col]):
            val = str(row[col]).strip()
            if val.endswith('.0'): return val[:-2]
            return val
    return None


# ================= 3. 主程序逻辑 (适配网页端) =================

uploaded_file = st.file_uploader("请拖拽或选择 Excel 文件 (.xlsx)", type=['xlsx'])

if uploaded_file:
    # 显示加载状态
    with st.spinner('正在极速处理中，请稍候...'):
        try:
            # 1. 读取 Excel 对象
            xls = pd.ExcelFile(uploaded_file)


            # 2. 读取核心配置 Sheet
            def find_config_sheet(keyword):
                for s in xls.sheet_names:
                    if keyword in s: return s
                return None


            s_demand = find_config_sheet("广告需求")
            s_style = find_config_sheet("款式名")
            s_model = find_config_sheet("型号名")

            # 检查必要 Sheet
            if not all([s_demand, s_style, s_model]):
                st.error(f"❌ Excel 缺少核心 Sheet！\n检测结果：广告需求={s_demand}, 款式名={s_style}, 型号名={s_model}")
                st.stop()

            # 读取数据
            df_demand = clean_df(pd.read_excel(xls, sheet_name=s_demand))
            df_style = clean_df(pd.read_excel(xls, sheet_name=s_style))
            df_model = clean_df(pd.read_excel(xls, sheet_name=s_model))

            # 3. 构建数据映射
            # (1) 款式映射
            if '缩写' not in df_style.columns or '款式全称' not in df_style.columns:
                st.error("❌ [款式表] 缺少 '缩写' 或 '款式全称' 列")
                st.stop()
            style_abbr_map = dict(zip(df_style['缩写'].astype(str).str.strip(), df_style['款式全称']))

            style_root_map = {}
            if '核心词根' in df_style.columns:
                for _, row in df_style.iterrows():
                    if pd.notna(row['核心词根']):
                        val = str(row['核心词根']).replace('，', ',')
                        style_root_map[row['款式全称']] = [x.strip() for x in val.split(',')]
            else:
                st.error("❌ [款式表] 缺少 '核心词根' 列")
                st.stop()

            # (2) 型号映射
            model_abbr_map = dict(zip(df_model['缩写'].astype(str).str.strip(), df_model['型号全称']))
            model_name_to_abbr = dict(zip(df_model['型号全称'], df_model['缩写']))

            col_fid = '对应词表标识'
            if col_fid not in df_model.columns:
                for c in df_model.columns:
                    if '词表标识' in c: col_fid = c; break

            if col_fid not in df_model.columns:
                st.error("❌ [型号表] 缺少 '对应词表标识' 列")
                st.stop()

            model_file_id_map = dict(zip(df_model['型号全称'], df_model[col_fid]))

            # 4. 解析广告需求
            parsed_data = []
            logs = []  # 记录日志

            for idx, row in df_demand.iterrows():
                sku = str(row['SKU']).strip()
                found_model = None
                found_style = None

                # 优先匹配长字符串
                for abbr in sorted(model_abbr_map.keys(), key=len, reverse=True):
                    if abbr in sku:
                        found_model = model_abbr_map[abbr]
                        break

                for abbr in sorted(style_abbr_map.keys(), key=len, reverse=True):
                    if abbr in sku:
                        found_style = style_abbr_map[abbr]
                        break

                if found_model and found_style:
                    parsed_data.append({
                        'sku': sku,
                        'bid': get_col(row, ['竞价']),
                        'budget': get_col(row, ['每日预算']),
                        'start_date': get_str(row, ['开始日期', 'Start Date']),
                        'match': str(row.get('匹配模式', '精准')).strip(),
                        'top': get_col(row, ['首页位置溢价%', '首页溢价%']),
                        'prod': get_col(row, ['商品页溢价%', '商品页位置溢价%']),
                        'rest': get_col(row, ['其余位置溢价%', '其余溢价%']),
                        'model': found_model,
                        'style': found_style
                    })
                else:
                    logs.append(f"⚠️ 跳过无效SKU: {sku}")

            if not parsed_data:
                st.warning("❌ 未识别到任何有效任务，请检查表格内容。")
                st.stop()

            # 5. 生成广告数据
            df_p = pd.DataFrame(parsed_data)
            grouped = df_p.groupby(['model', 'style'])

            output_rows = []
            report_rows = []

            for (model, style), group in grouped:
                abbr = model_name_to_abbr.get(model)
                roots = style_root_map.get(style)
                file_id = model_file_id_map.get(model)

                if not roots or not file_id:
                    continue

                # --- 核心：精准读取关键词 Sheet ---
                target_sheet_name = f"关键词-{file_id}"
                kw_sheet = find_sheet_strict(xls, target_sheet_name)

                valid_keywords = []
                if kw_sheet:
                    df_kw = clean_df(pd.read_excel(xls, sheet_name=kw_sheet))
                    if '分类' in df_kw.columns and '关键词' in df_kw.columns:
                        # 排除竞品
                        df_kw = df_kw[~df_kw['分类'].astype(str).str.contains('品牌', na=False)]
                        # 筛选词根
                        for _, k_row in df_kw.iterrows():
                            cat = str(k_row['分类']).strip()
                            if cat in roots:
                                valid_keywords.append(k_row['关键词'])
                        valid_keywords = list(set(valid_keywords))
                    else:
                        logs.append(f"⚠️ Sheet [{kw_sheet}] 缺少 '分类' 或 '关键词' 列")
                else:
                    logs.append(f"🔸 警告: 找不到 Sheet -> [{target_sheet_name}]")

                first = group.iloc[0]
                camp_name = f"{abbr}-{style}-SP"

                # --- 写入中文行 ---
                # 1. 广告活动
                output_rows.append({
                    '产品': '商品推广', '实体层级': '广告活动', '操作': '创建',
                    '广告活动编号': camp_name, '广告活动名称': camp_name,
                    '投放类型': '手动', '状态': '已启用',
                    '每日预算': first['budget'],
                    '开始日期': first['start_date'],
                    '竞价方案': '固定竞价'
                })

                # 2. 竞价调整
                if first['top'] > 0:
                    output_rows.append(
                        {'产品': '商品推广', '实体层级': '竞价调整', '操作': '创建', '广告活动编号': camp_name,
                         '广告位': '广告位：搜索结果首页首位', '百分比': first['top']})
                if first['prod'] > 0:
                    output_rows.append(
                        {'产品': '商品推广', '实体层级': '竞价调整', '操作': '创建', '广告活动编号': camp_name,
                         '广告位': '广告位：商品页面', '百分比': first['prod']})
                if first['rest'] > 0:
                    output_rows.append(
                        {'产品': '商品推广', '实体层级': '竞价调整', '操作': '创建', '广告活动编号': camp_name,
                         '广告位': '广告位：搜索结果的其余位置', '百分比': first['rest']})

                # 3. 广告组 (固定竞价=1)
                output_rows.append({
                    '产品': '商品推广', '实体层级': '广告组', '操作': '创建',
                    '广告活动编号': camp_name, '广告组编号': camp_name,
                    '广告组名称': camp_name, '状态': '已启用',
                    '广告组默认竞价': 1
                })

                # 4. 商品广告
                skus_list = []
                for _, item in group.iterrows():
                    skus_list.append(item['sku'])
                    output_rows.append({
                        '产品': '商品推广', '实体层级': '商品广告', '操作': '创建',
                        '广告活动编号': camp_name, '广告组编号': camp_name,
                        'SKU': item['sku'], '状态': '已启用'
                    })

                # 5. 关键词 (仅此处填参考竞价=1)
                for kw in valid_keywords:
                    output_rows.append({
                        '产品': '商品推广', '实体层级': '关键词', '操作': '创建',
                        '广告活动编号': camp_name, '广告组编号': camp_name,
                        '关键词文本': kw,
                        '匹配类型': first['match'],
                        '竞价': group['bid'].max(),
                        '状态': '已启用',
                        '广告组默认竞价（仅供参考）': 1
                    })

                # 说明书
                report_rows.append({
                    '广告活动名称': camp_name,
                    '型号': model, '款式': style,
                    '目标Sheet': target_sheet_name,
                    '匹配状态': '✅ 成功' if kw_sheet else '❌ 失败',
                    '核心词根': str(roots),
                    '关键词数': len(valid_keywords),
                    'SKU列表': " | ".join(skus_list)
                })

                logs.append(f"🎉 生成: {camp_name} ({len(valid_keywords)} 词)")

            # 6. 输出结果
            if output_rows:
                st.success(f"✅ 成功生成 {len(output_rows)} 行数据！")

                # 定义列顺序
                cols = ['产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告活动名称', '广告组名称',
                        '投放类型', '状态', '每日预算', '开始日期', '竞价方案',
                        '广告组默认竞价', '广告组默认竞价（仅供参考）',
                        'SKU', '竞价', '匹配类型', '关键词文本', '广告位', '百分比']

                df_out = pd.DataFrame(output_rows)
                for c in cols:
                    if c not in df_out.columns: df_out[c] = None

                # 转换CSV格式 (UTF-8-SIG)
                csv_upload = df_out[cols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

                # 下载区
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 下载【上传表】(中文版)",
                        data=csv_upload,
                        file_name=f"【中文版】SP广告上传表_{timestamp}.csv",
                        mime='text/csv'
                    )

                with col2:
                    if report_rows:
                        csv_report = pd.DataFrame(report_rows).to_csv(index=False, encoding='utf-8-sig').encode(
                            'utf-8-sig')
                        st.download_button(
                            label="📄 下载【说明书】",
                            data=csv_report,
                            file_name=f"【说明书】广告生成详情_{timestamp}.csv",
                            mime='text/csv'
                        )
            else:
                st.warning("⚠️ 未生成任何数据。")

            # 折叠显示日志
            with st.expander("查看详细运行日志"):
                for log in logs:
                    st.text(log)

        except Exception as e:
            st.error(f"❌ 程序发生错误: {e}")
            import traceback

            st.text(traceback.format_exc())