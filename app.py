import datetime
import io

import pandas as pd
import streamlit as st


KEYWORDS_PER_AD_GROUP_LIMIT = 950


# ================= 1. 网页基础设置 =================
st.set_page_config(page_title="亚马逊广告自动生成工具", page_icon="🚀", layout="wide")

st.title("🚀 亚马逊广告自动生成工具")
st.markdown(
    """
**使用说明：**
1. 准备好你的 Excel 配置文件（格式需与 **广告自动生成工具.xlsx** 一致）。
2. 可同时填写 **广告需求-sp** 和 **广告需求-视频**。
3. 点击下方按钮上传文件，系统将自动处理并提供 Excel (.xlsx) 下载。
"""
)


# ================= 2. 通用工具函数 =================

def clean_df(df):
    """清洗 DataFrame：去列名空格、去空行。"""
    if df is None:
        return None
    df.columns = [str(c).strip().replace("\ufeff", "") for c in df.columns]
    df.dropna(how="all", inplace=True)
    return df


def find_sheet_strict(xls, target_name):
    """精准查找 sheet：忽略大小写和首尾空格，但必须完全相等。"""
    target = target_name.lower().strip()
    for sheet in xls.sheet_names:
        if sheet.lower().strip() == target:
            return sheet
    return None


def find_sheet_contains(xls, keyword):
    for sheet in xls.sheet_names:
        if keyword in sheet:
            return sheet
    return None


def get_col(row, col_list, default=0.0):
    for col in col_list:
        if col in row and pd.notna(row[col]):
            try:
                if isinstance(row[col], str):
                    value = row[col].strip().replace("%", "").replace("％", "")
                    return float(value)
                return float(row[col])
            except (TypeError, ValueError):
                return default
    return default


def get_str(row, col_list, default=None):
    for col in col_list:
        if col in row and pd.notna(row[col]):
            val = str(row[col]).strip()
            if val.endswith(".0"):
                return val[:-2]
            return val
    return default


def normalize_as_text(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.endswith(".0"):
        return text[:-2]
    return text or None


def ordered_unique(values):
    seen = set()
    result = []
    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text and text not in seen:
            seen.add(text)
            result.append(text)
    return result


def chunk_list(values, size):
    return [values[i:i + size] for i in range(0, len(values), size)]


def make_ad_group_name(base_name, group_index):
    if group_index == 1:
        return base_name
    return f"{base_name}-{group_index}"


def build_mappings(df_style, df_model):
    if "缩写" not in df_style.columns or "款式全称" not in df_style.columns:
        raise ValueError("❌ [款式表] 缺少 '缩写' 或 '款式全称' 列")
    if "核心词根" not in df_style.columns:
        raise ValueError("❌ [款式表] 缺少 '核心词根' 列")
    if "缩写" not in df_model.columns or "型号全称" not in df_model.columns:
        raise ValueError("❌ [型号表] 缺少 '缩写' 或 '型号全称' 列")

    style_abbr_map = dict(zip(df_style["缩写"].astype(str).str.strip(), df_style["款式全称"]))
    style_root_map = {}
    for _, row in df_style.iterrows():
        if pd.notna(row["核心词根"]):
            val = str(row["核心词根"]).replace("，", ",")
            roots = [x.strip() for x in val.split(",") if x.strip()]
            style_root_map[row["款式全称"]] = roots

    model_abbr_map = dict(zip(df_model["缩写"].astype(str).str.strip(), df_model["型号全称"]))
    model_name_to_abbr = dict(zip(df_model["型号全称"], df_model["缩写"]))

    col_fid = "对应词表标识"
    if col_fid not in df_model.columns:
        for col in df_model.columns:
            if "词表标识" in col:
                col_fid = col
                break
    if col_fid not in df_model.columns:
        raise ValueError("❌ [型号表] 缺少 '对应词表标识' 列")

    model_file_id_map = dict(zip(df_model["型号全称"], df_model[col_fid]))
    return style_abbr_map, style_root_map, model_abbr_map, model_name_to_abbr, model_file_id_map


def load_brand_settings(xls, logs):
    sheet = find_sheet_strict(xls, "品牌设置")
    if not sheet:
        return {}, None

    df_brand = clean_df(pd.read_excel(xls, sheet_name=sheet))
    if df_brand is None or df_brand.empty:
        return {}, None

    site_col = "站点" if "站点" in df_brand.columns else None
    brand_col = None
    for candidate in ["品牌实体编号", "Brand Entity ID", "brandEntityId", "Brand Entity Id"]:
        if candidate in df_brand.columns:
            brand_col = candidate
            break

    if not brand_col:
        logs.append("⚠️ [品牌设置] 缺少 '品牌实体编号' 列，视频广告会被跳过。")
        return {}, None

    brand_by_site = {}
    default_brand_id = None
    for _, row in df_brand.iterrows():
        brand_id = normalize_as_text(row.get(brand_col))
        if not brand_id:
            continue
        site = normalize_as_text(row.get(site_col)) if site_col else None
        if site:
            brand_by_site[site] = brand_id
        if not default_brand_id:
            default_brand_id = brand_id

    return brand_by_site, default_brand_id


def parse_sku_info(sku, style_abbr_map, model_abbr_map):
    found_model = None
    found_style = None

    for abbr in sorted(model_abbr_map.keys(), key=len, reverse=True):
        if abbr and abbr in sku:
            found_model = model_abbr_map[abbr]
            break

    for abbr in sorted(style_abbr_map.keys(), key=len, reverse=True):
        if abbr and abbr in sku:
            found_style = style_abbr_map[abbr]
            break

    return found_model, found_style


def load_keywords(xls, model, style, style_root_map, model_file_id_map, logs):
    roots = style_root_map.get(style)
    file_id = model_file_id_map.get(model)
    if not roots or not file_id:
        logs.append(f"⚠️ 跳过关键词：型号/款式缺少词根或词表标识 -> {model} / {style}")
        return [], f"关键词-{file_id}" if file_id else None, roots

    target_sheet_name = f"关键词-{file_id}"
    kw_sheet = find_sheet_strict(xls, target_sheet_name)
    if not kw_sheet:
        logs.append(f"🔸 警告: 找不到 Sheet -> [{target_sheet_name}]")
        return [], target_sheet_name, roots

    df_kw = clean_df(pd.read_excel(xls, sheet_name=kw_sheet))
    if "分类" not in df_kw.columns or "关键词" not in df_kw.columns:
        logs.append(f"⚠️ Sheet [{kw_sheet}] 缺少 '分类' 或 '关键词' 列")
        return [], target_sheet_name, roots

    df_kw = df_kw[~df_kw["分类"].astype(str).str.contains("品牌", na=False)]
    df_kw = df_kw[df_kw["分类"].astype(str).str.strip().isin(roots)]
    return ordered_unique(df_kw["关键词"]), target_sheet_name, roots


# ================= 3. SP 广告生成 =================

def generate_sp_rows(xls, df_demand, maps, logs):
    style_abbr_map, style_root_map, model_abbr_map, model_name_to_abbr, model_file_id_map = maps
    output_rows = []
    report_rows = []
    parsed_data = []

    if df_demand is None or df_demand.empty:
        return output_rows, report_rows
    if "SKU" not in df_demand.columns:
        logs.append("⚠️ [广告需求-sp] 缺少 SKU 列，跳过 SP。")
        return output_rows, report_rows

    for _, row in df_demand.iterrows():
        sku = normalize_as_text(row.get("SKU"))
        if not sku:
            continue
        found_model, found_style = parse_sku_info(sku, style_abbr_map, model_abbr_map)

        if found_model and found_style:
            parsed_data.append(
                {
                    "sku": sku,
                    "bid": get_col(row, ["竞价"]),
                    "budget": get_col(row, ["每日预算"]),
                    "start_date": get_str(row, ["开始日期", "Start Date"]),
                    "match": str(row.get("匹配模式", "精准")).strip(),
                    "top": get_col(row, ["首页位置溢价%", "首页溢价%"]),
                    "prod": get_col(row, ["商品页溢价%", "商品页位置溢价%"]),
                    "rest": get_col(row, ["其余位置溢价%", "其余溢价%"]),
                    "model": found_model,
                    "style": found_style,
                }
            )
        else:
            logs.append(f"⚠️ [SP] 跳过无法识别型号/款式的 SKU: {sku}")

    if not parsed_data:
        return output_rows, report_rows

    df_p = pd.DataFrame(parsed_data)
    grouped = df_p.groupby(["model", "style"])

    for (model, style), group in grouped:
        abbr = model_name_to_abbr.get(model)
        valid_keywords, target_sheet_name, roots = load_keywords(
            xls, model, style, style_root_map, model_file_id_map, logs
        )
        if not valid_keywords:
            continue

        first = group.iloc[0]
        camp_name = f"{abbr}-{style}-SP"

        output_rows.append(
            {
                "产品": "商品推广",
                "实体层级": "广告活动",
                "操作": "创建",
                "广告活动编号": camp_name,
                "广告活动名称": camp_name,
                "投放类型": "手动",
                "状态": "已启用",
                "每日预算": first["budget"],
                "开始日期": first["start_date"],
                "竞价方案": "固定竞价",
            }
        )

        if first["top"] > 0:
            output_rows.append(
                {
                    "产品": "商品推广",
                    "实体层级": "竞价调整",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告位": "广告位：搜索结果首页首位",
                    "百分比": first["top"],
                }
            )
        if first["prod"] > 0:
            output_rows.append(
                {
                    "产品": "商品推广",
                    "实体层级": "竞价调整",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告位": "广告位：商品页面",
                    "百分比": first["prod"],
                }
            )
        if first["rest"] > 0:
            output_rows.append(
                {
                    "产品": "商品推广",
                    "实体层级": "竞价调整",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告位": "广告位：搜索结果的其余位置",
                    "百分比": first["rest"],
                }
            )

        skus_list = [item["sku"] for _, item in group.iterrows()]
        keyword_groups = chunk_list(valid_keywords, KEYWORDS_PER_AD_GROUP_LIMIT)

        for group_index, keywords in enumerate(keyword_groups, start=1):
            ad_group_name = make_ad_group_name(camp_name, group_index)
            output_rows.append(
                {
                    "产品": "商品推广",
                    "实体层级": "广告组",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告组编号": ad_group_name,
                    "广告组名称": ad_group_name,
                    "状态": "已启用",
                    "广告组默认竞价": 1,
                }
            )

            for _, item in group.iterrows():
                output_rows.append(
                    {
                        "产品": "商品推广",
                        "实体层级": "商品广告",
                        "操作": "创建",
                        "广告活动编号": camp_name,
                        "广告组编号": ad_group_name,
                        "SKU": item["sku"],
                        "状态": "已启用",
                    }
                )

            for kw in keywords:
                output_rows.append(
                    {
                        "产品": "商品推广",
                        "实体层级": "关键词",
                        "操作": "创建",
                        "广告活动编号": camp_name,
                        "广告组编号": ad_group_name,
                        "关键词文本": kw,
                        "匹配类型": first["match"],
                        "竞价": group["bid"].max(),
                        "状态": "已启用",
                        "广告组默认竞价（仅供参考）": 1,
                    }
                )

        report_rows.append(
            {
                "广告类型": "SP",
                "广告活动名称": camp_name,
                "型号": model,
                "款式": style,
                "目标Sheet": target_sheet_name,
                "匹配状态": "✅ 成功",
                "核心词根": str(roots),
                "关键词数": len(valid_keywords),
                "广告组数": len(keyword_groups),
                "每组关键词上限": KEYWORDS_PER_AD_GROUP_LIMIT,
                "SKU列表": " | ".join(skus_list),
            }
        )
        logs.append(f"🎉 [SP] 生成: {camp_name} ({len(valid_keywords)} 词 / {len(keyword_groups)} 个广告组)")

    return output_rows, report_rows


# ================= 4. 视频广告生成 =================

def dedupe_campaign_name(base_name, used_names, asin):
    if base_name not in used_names:
        used_names.add(base_name)
        return base_name
    suffix = asin or len(used_names) + 1
    candidate = f"{base_name}-{suffix}"
    counter = 2
    while candidate in used_names:
        candidate = f"{base_name}-{suffix}-{counter}"
        counter += 1
    used_names.add(candidate)
    return candidate


def generate_video_rows(xls, df_demand, maps, brand_by_site, default_brand_id, logs):
    style_abbr_map, style_root_map, model_abbr_map, model_name_to_abbr, model_file_id_map = maps
    output_rows = []
    report_rows = []
    used_names = set()

    if df_demand is None or df_demand.empty:
        return output_rows, report_rows

    required_cols = ["SKU", "ASIN", "视频媒体编号", "竞价", "每日预算", "匹配模式", "开始日期"]
    missing = [col for col in required_cols if col not in df_demand.columns]
    if missing:
        logs.append(f"⚠️ [广告需求-视频] 缺少列：{', '.join(missing)}，跳过视频广告。")
        return output_rows, report_rows

    for _, row in df_demand.iterrows():
        sku = normalize_as_text(row.get("SKU"))
        asin = normalize_as_text(row.get("ASIN"))
        video_media_id = normalize_as_text(row.get("视频媒体编号"))
        if not sku and not asin and not video_media_id:
            continue
        if not sku or not asin or not video_media_id:
            logs.append(f"⚠️ [视频] 跳过缺少 SKU/ASIN/视频媒体编号 的行: SKU={sku}, ASIN={asin}")
            continue

        site = normalize_as_text(row.get("站点"))
        brand_entity_id = brand_by_site.get(site) or default_brand_id
        if not brand_entity_id:
            logs.append(f"⚠️ [视频] 跳过缺少品牌实体编号的行: 站点={site}, SKU={sku}, ASIN={asin}")
            continue

        found_model, found_style = parse_sku_info(sku, style_abbr_map, model_abbr_map)
        if not found_model or not found_style:
            logs.append(f"⚠️ [视频] 跳过无法识别型号/款式的 SKU: {sku}")
            continue

        valid_keywords, target_sheet_name, roots = load_keywords(
            xls, found_model, found_style, style_root_map, model_file_id_map, logs
        )
        if not valid_keywords:
            continue

        abbr = model_name_to_abbr.get(found_model)
        base_name = f"{abbr}-{found_style}-视频"
        camp_name = dedupe_campaign_name(base_name, used_names, asin)
        match_type = get_str(row, ["匹配模式"], "精准")
        placement_adjustments = [
            ("搜索结果首页首位", get_col(row, ["首页位置溢价%", "首页溢价%"])),
            ("商品详情页", get_col(row, ["商品页溢价%", "商品页位置溢价%"])),
            ("其他", get_col(row, ["其余位置溢价%", "其余溢价%"])),
        ]

        keyword_groups = chunk_list(valid_keywords, KEYWORDS_PER_AD_GROUP_LIMIT)
        output_rows.append(
            {
                "产品": "品牌推广",
                "实体层级": "广告活动",
                "操作": "创建",
                "广告活动编号": camp_name,
                "广告活动名称": camp_name,
                "开始日期": get_str(row, ["开始日期"]),
                "状态": "已启用",
                "品牌实体编号": brand_entity_id,
                "预算类型": "每日",
                "预算": get_col(row, ["每日预算"]),
                "竞价优化": False,
            }
        )
        for placement, percentage in placement_adjustments:
            output_rows.append(
                {
                    "产品": "品牌推广",
                    "实体层级": "按广告位调整竞价",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告位": placement,
                    "百分比": percentage,
                }
            )

        ad_group_keyword_parts = []
        for group_index, keywords in enumerate(keyword_groups, start=1):
            ad_group_name = make_ad_group_name(camp_name, group_index)
            ad_group_keyword_parts.append(f"{ad_group_name}:{len(keywords)}")
            output_rows.append(
                {
                    "产品": "品牌推广",
                    "实体层级": "广告组",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告组编号": ad_group_name,
                    "广告组名称": ad_group_name,
                    "状态": "已启用",
                }
            )
            output_rows.append(
                {
                    "产品": "品牌推广",
                    "实体层级": "视频广告",
                    "操作": "创建",
                    "广告活动编号": camp_name,
                    "广告组编号": ad_group_name,
                    "广告名称": ad_group_name,
                    "状态": "已启用",
                    "落地页类型": "商品详情页",
                    "同意翻译": False,
                    "创意素材 ASIN": asin,
                    "视频素材编号": video_media_id,
                }
            )
            for kw in keywords:
                output_rows.append(
                    {
                        "产品": "品牌推广",
                        "实体层级": "关键词",
                        "操作": "创建",
                        "广告活动编号": camp_name,
                        "广告组编号": ad_group_name,
                        "状态": "已启用",
                        "竞价": get_col(row, ["竞价"]),
                        "关键词文本": kw,
                        "匹配类型": match_type,
                    }
                )

        report_rows.append(
            {
                "广告类型": "视频",
                "广告活动名称": camp_name,
                "型号": found_model,
                "款式": found_style,
                "目标Sheet": target_sheet_name,
                "匹配状态": "✅ 成功",
                "核心词根": str(roots),
                "关键词数": len(valid_keywords),
                "广告组数": len(keyword_groups),
                "每组关键词上限": KEYWORDS_PER_AD_GROUP_LIMIT,
                "广告组关键词分配": " | ".join(ad_group_keyword_parts),
                "广告位调整": " | ".join(f"{placement}:{percentage}" for placement, percentage in placement_adjustments),
                "SKU列表": sku,
                "ASIN": asin,
                "视频媒体编号": video_media_id,
            }
        )
        logs.append(f"🎬 [视频] 生成: {camp_name} ({len(valid_keywords)} 词 / {len(keyword_groups)} 个广告组)")

    return output_rows, report_rows


# ================= 5. 输出列顺序 =================

SP_COLS = [
    "产品",
    "实体层级",
    "操作",
    "广告活动编号",
    "广告组编号",
    "广告活动名称",
    "广告组名称",
    "投放类型",
    "状态",
    "每日预算",
    "开始日期",
    "竞价方案",
    "广告组默认竞价",
    "广告组默认竞价（仅供参考）",
    "SKU",
    "竞价",
    "匹配类型",
    "关键词文本",
    "广告位",
    "百分比",
]

VIDEO_COLS = [
    "产品",
    "实体层级",
    "操作",
    "广告活动编号",
    "广告组合编号",
    "广告组编号",
    "广告编号",
    "关键词编号",
    "商品投放 ID",
    "广告活动名称",
    "广告组名称",
    "广告名称",
    "广告活动名称（仅供参考）",
    "广告组名称（仅供参考）",
    "广告组合名称（仅供参考）",
    "开始日期",
    "结束日期",
    "状态",
    "品牌实体编号",
    "广告活动状态（仅供参考）",
    "广告活动开展状态（仅供参考）",
    "广告活动投放状态详情（仅供参考）",
    "基于规则的预算正在处理（仅供参考）",
    "基于规则的预算名称（仅供参考）",
    "基于规则的预算值（仅供参考）",
    "基于规则的预算编号（仅供参考）",
    "广告组投放状态（仅供参考）",
    "广告组投放状态详情（仅供参考）",
    "预算类型",
    "预算",
    "竞价优化",
    "商品位置",
    "竞价",
    "广告位",
    "百分比",
    "受众编号",
    "购物者群体占比",
    "购物者群体类型",
    "站点名称（仅供参考）",
    "关键词文本",
    "匹配类型",
    "母语关键词",
    "母语区域",
    "拓展商品投放编号",
    "拓展商品投放名称（仅供参考）",
    "广告投放状态（仅供参考）",
    "广告投放状态详情（仅供参考）",
    "落地页 URL",
    "落地页 ASIN",
    "落地页类型",
    "品牌名称",
    "同意翻译",
    "品牌徽标素材编号",
    "品牌徽标 URL（仅供参考）",
    "品牌徽标裁剪",
    "自定义图片",
    "创意素材标题",
    "创意素材 ASIN",
    "视频素材编号",
    "原始视频素材编号（仅供参考）",
    "子页面",
    "商品排除项",
    "广告标题",
    "站点",
    "展示量",
    "点击量",
    "点击率",
    "花费",
    "销量",
    "订单数量",
    "商品数量",
    "转化率",
    "ACOS",
    "CPC",
    "ROAS",
]


def align_columns(rows, cols):
    df = pd.DataFrame(rows)
    for col in cols:
        if col not in df.columns:
            df[col] = None
    return df[cols]


# ================= 6. 主程序逻辑 =================

uploaded_file = st.file_uploader("请拖拽或选择 Excel 文件 (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("正在处理中，请稍候..."):
        try:
            xls = pd.ExcelFile(uploaded_file)

            s_sp = find_sheet_strict(xls, "广告需求-sp") or find_sheet_strict(xls, "广告需求")
            s_video = find_sheet_strict(xls, "广告需求-视频")
            s_style = find_sheet_contains(xls, "款式名")
            s_model = find_sheet_contains(xls, "型号名")

            if not s_style or not s_model:
                st.error(f"❌ Excel 缺少核心 Sheet！检测结果：款式名={s_style}, 型号名={s_model}")
                st.stop()
            if not s_sp and not s_video:
                st.error("❌ Excel 至少需要包含 [广告需求-sp] 或 [广告需求-视频] 其中一个 Sheet。")
                st.stop()

            df_style = clean_df(pd.read_excel(xls, sheet_name=s_style))
            df_model = clean_df(pd.read_excel(xls, sheet_name=s_model))
            maps = build_mappings(df_style, df_model)

            logs = []
            brand_by_site, default_brand_id = load_brand_settings(xls, logs)
            report_rows = []
            sp_rows = []
            video_rows = []

            if s_sp:
                df_sp = clean_df(pd.read_excel(xls, sheet_name=s_sp))
                sp_rows, sp_report = generate_sp_rows(xls, df_sp, maps, logs)
                report_rows.extend(sp_report)
            else:
                logs.append("ℹ️ 未找到 [广告需求-sp]，跳过 SP。")

            if s_video:
                df_video = clean_df(pd.read_excel(xls, sheet_name=s_video))
                video_rows, video_report = generate_video_rows(
                    xls, df_video, maps, brand_by_site, default_brand_id, logs
                )
                report_rows.extend(video_report)
            else:
                logs.append("ℹ️ 未找到 [广告需求-视频]，跳过视频广告。")

            if not sp_rows and not video_rows:
                st.warning("⚠️ 未生成任何数据，请检查需求表。")
                with st.expander("查看详细运行日志"):
                    for log in logs:
                        st.text(log)
                st.stop()

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_upload_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_upload_buffer, engine="openpyxl") as writer:
                if sp_rows:
                    align_columns(sp_rows, SP_COLS).to_excel(writer, index=False, sheet_name="商品推广活动")
                if video_rows:
                    align_columns(video_rows, VIDEO_COLS).to_excel(
                        writer, index=False, sheet_name="品牌推广多个广告组广告活动 1"
                    )
            excel_upload_data = excel_upload_buffer.getvalue()

            st.success(f"✅ 成功生成 SP {len(sp_rows)} 行，视频 {len(video_rows)} 行！")

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 下载【上传表】(.xlsx)",
                    data=excel_upload_data,
                    file_name=f"【中文版】广告上传表_SP加视频_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with col2:
                if report_rows:
                    excel_report_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_report_buffer, engine="openpyxl") as writer:
                        pd.DataFrame(report_rows).to_excel(writer, index=False, sheet_name="生成说明书")
                    st.download_button(
                        label="📄 下载【说明书】(.xlsx)",
                        data=excel_report_buffer.getvalue(),
                        file_name=f"【说明书】广告生成详情_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            with st.expander("查看详细运行日志"):
                for log in logs:
                    st.text(log)

        except Exception as e:
            st.error(f"❌ 程序发生错误: {e}")
            import traceback

            st.text(traceback.format_exc())
