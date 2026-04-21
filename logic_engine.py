"""
Data processing, image indexing, and Excel export.
"""
import io
import base64
import pandas as pd
from PIL import Image, ImageStat
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter


def excel_text_str(val):
    """Normalize cell text; avoid scientific notation for long numbers."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def qty_gt_one(val):
    try:
        return float(str(val).strip()) > 1
    except Exception:
        return False


def build_image_index(uploaded_images):
    """Index uploaded images by filename stem (lowercase)."""
    image_map = {}
    allowed_ext = {".jpg", ".jpeg", ".png"}
    for img in uploaded_images or []:
        file_name = (img.name or "").strip().replace("\\", "/").split("/")[-1]
        if not file_name:
            continue
        lower_name = file_name.lower()
        dot_idx = lower_name.rfind(".")
        if dot_idx == -1:
            continue
        ext = lower_name[dot_idx:]
        if ext not in allowed_ext:
            continue
        key = lower_name[:dot_idx]
        image_map[key] = {"bytes": img.getvalue(), "ext": ext}
    return image_map


def extract_size_parts(size_val):
    try:
        size_num = int(float(str(size_val).strip()))
    except Exception:
        return None, None
    digits = str(abs(size_num))
    if len(digits) == 4:
        return int(digits[:2]), int(digits[2:])
    if len(digits) == 3:
        return int(digits[:1]), int(digits[1:])
    if len(digits) >= 2:
        mid = len(digits) // 2
        return int(digits[:mid]), int(digits[mid:])
    if len(digits) == 1:
        val = int(digits)
        return val, val
    return None, None


def resolve_row_image_bytes(row, image_map):
    img_name_key = excel_text_str(row.get("图片名称", "")).lower()
    if img_name_key and img_name_key in image_map and image_map[img_name_key].get("bytes"):
        return image_map[img_name_key]["bytes"]
    order_key = f"{excel_text_str(row.get('purchase-date', ''))}_{excel_text_str(row.get('运单号', ''))}".lower()
    if order_key and order_key in image_map and image_map[order_key].get("bytes"):
        return image_map[order_key]["bytes"]
    return None


def detect_hanger_and_size(img_bytes, size_val):
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        img = img.resize((400, 400))
    except Exception:
        return "Canvas"

    width, height = img.size
    orientation = "portrait" if height >= width else "landscape"

    top_h = min(15, height)
    top_w = min(400, width)
    if top_h <= 0 or top_w <= 0:
        return "Canvas"
    top_region = img.crop((0, 0, top_w, top_h))
    top_std = ImageStat.Stat(top_region).stddev
    mean_std = sum(top_std) / len(top_std) if top_std else 0
    if mean_std < 8:
        return "Canvas"

    band_y0 = min(15, height)
    band_y1 = min(25, height)
    if band_y1 <= band_y0:
        return "Canvas"
    color_region = img.crop((0, band_y0, top_w, band_y1))
    avg = ImageStat.Stat(color_region).mean
    if len(avg) < 3:
        return "Canvas"
    r, g, b = avg[:3]
    hanger_type = "Canvas"
    if r < 60 and g < 60 and b < 60:
        hanger_type = "Black Hanger"
    elif r > 150 and g > 120 and b < 100:
        hanger_type = "Wood Hanger"
    if hanger_type == "Canvas":
        return "Canvas"

    first_len, second_len = extract_size_parts(size_val)
    if first_len is None or second_len is None:
        return hanger_type
    selected_len = first_len if orientation == "portrait" else second_len
    return f'{hanger_type} - {selected_len}"'


def apply_material_recognition(df, image_map):
    if df.empty:
        return df
    updated_df = df.copy()
    for i, row in updated_df.iterrows():
        img_bytes = resolve_row_image_bytes(row, image_map)
        updated_df.at[i, "材质"] = detect_hanger_and_size(img_bytes, row.get("尺寸")) if img_bytes else "Canvas"
    updated_df["图片名称"] = (
        updated_df["序号E"].astype(str) + "-" + updated_df["材质"] + "-" + updated_df["尺寸"].astype(str)
    )
    return updated_df


def process_data(source_df, ref_df):
    target = pd.DataFrame()
    target["运单号"] = source_df["order-item-id"].map(excel_text_str)
    target["purchase-date"] = source_df["order-id"].map(excel_text_str)
    target["SKU"] = source_df["sku"]
    target["图片"] = source_df.iloc[:, 2]
    target["数量"] = source_df["quantity-purchased"]
    target["姓名"] = source_df["recipient-name"]
    target["地址一"] = source_df["ship-address-1"]
    target["地址二"] = source_df["ship-address-2"].fillna("0")
    target["城市"] = source_df["ship-city"]
    target["州"] = source_df["ship-state"]
    target["邮编"] = source_df["ship-postal-code"]
    target["电话"] = source_df["ship-phone-number"].map(excel_text_str)
    target["Original Row Index"] = source_df.index + 2

    size_dict = dict(zip(ref_df["SKU"].astype(str), ref_df["尺寸"]))
    target["尺寸_原始"] = target["SKU"].astype(str).map(size_dict)
    target["尺寸_数值"] = pd.to_numeric(target["尺寸_原始"], errors="coerce").fillna(0).astype(int)
    target["尺寸"] = target["尺寸_数值"]

    target["Identity"] = target["姓名"].astype(str) + target["地址一"].astype(str)

    package_info = target.groupby("Identity").agg(
        row_count=("SKU", "count"),
        min_size=("尺寸_数值", "min"),
        max_size=("尺寸_数值", "max"),
        is_mixed_size=("尺寸_数值", lambda x: x.nunique() > 1),
    ).reset_index()

    def categorize(row):
        if row["is_mixed_size"]:
            return "B2"
        if row["row_count"] == 1 and row["max_size"] > 2436:
            return "A2"
        if row["max_size"] <= 2436:
            return "A1B1"
        return "A2"

    package_info["Category"] = package_info.apply(categorize, axis=1)
    target = target.merge(package_info[["Identity", "Category", "min_size", "max_size"]], on="Identity")

    tier1 = target[target["Category"] == "A1B1"].sort_values(
        ["max_size", "Identity", "尺寸_数值"], ascending=[False, True, False]
    )
    tier2 = target[target["Category"] == "B2"].sort_values(
        ["min_size", "Identity", "尺寸_数值"], ascending=[True, True, True]
    )
    tier3 = target[target["Category"] == "A2"].sort_values(
        ["尺寸_数值", "Identity"], ascending=[True, True]
    )

    final_df = pd.concat([tier1, tier2, tier3], ignore_index=True)
    final_df = final_df.merge(package_info[["Identity", "row_count"]], on="Identity", how="left")

    final_df["材质"] = "Canvas"
    final_df["材质"] = final_df["材质"].replace("画芯", "Canvas")

    current_idx = 0
    last_id = None
    d_column = []
    e_column = []
    id_counts = {}

    for _, row in final_df.iterrows():
        curr_id = row["Identity"]
        n_rows = int(row["row_count"])
        if curr_id != last_id:
            current_idx += 1
            id_counts[curr_id] = 1
            d_column.append(current_idx)
            if n_rows > 1:
                e_column.append(f"{current_idx}-1")
            else:
                e_column.append(current_idx)
        else:
            id_counts[curr_id] += 1
            d_column.append(current_idx)
            e_column.append(f"{current_idx}-{id_counts[curr_id]}")
        last_id = curr_id

    final_df["序号D"] = d_column
    final_df["序号E"] = e_column
    final_df["图片名称"] = final_df["序号E"].astype(str) + "-" + final_df["材质"] + "-" + final_df["尺寸"].astype(str)

    col_order = [
        "运单号",
        "purchase-date",
        "SKU",
        "序号D",
        "序号E",
        "图片",
        "材质",
        "尺寸",
        "数量",
        "姓名",
        "地址一",
        "地址二",
        "城市",
        "州",
        "邮编",
        "电话",
        "图片名称",
    ]
    return final_df[col_order + ["Original Row Index"]]


def save_to_excel_with_merge(df, source_file_bytes, image_map):
    output = io.BytesIO()
    wb = load_workbook(io.BytesIO(source_file_bytes), data_only=False, keep_vba=True)
    source_ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb[wb.sheetnames[0]]
    source_sheet_name = source_ws.title

    if "最终打印清单" in wb.sheetnames:
        del wb["最终打印清单"]
    ws = wb.create_sheet("最终打印清单")

    text_cols = {1, 2, 16}
    img_col = 6
    output_cols = [
        "运单号",
        "purchase-date",
        "SKU",
        "序号D",
        "序号E",
        "图片",
        "材质",
        "尺寸",
        "数量",
        "姓名",
        "地址一",
        "地址二",
        "城市",
        "州",
        "邮编",
        "电话",
        "图片名称",
    ]

    ws.append(output_cols)

    ws.column_dimensions["F"].width = 28
    source_ws.column_dimensions["C"].width = 35
    row_height_pt = 240
    body_font = Font(name="Microsoft YaHei", size=11)
    header_font = Font(name="Microsoft YaHei", size=11, bold=True)
    header_fill = PatternFill(fill_type="solid", fgColor="EDEDED")
    thin_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    base_widths = {
        1: 20,
        2: 20,
        3: 16,
        4: 8,
        5: 10,
        6: 28,
        7: 10,
        8: 8,
        9: 8,
        10: 16,
        11: 24,
        12: 24,
        13: 12,
        14: 10,
        15: 12,
        16: 18,
        17: 22,
    }
    core_cols = {1, 2, 4, 5, 6, 16}
    for col_idx, width in base_widths.items():
        if col_idx == 6:
            ws.column_dimensions["F"].width = 28
            continue
        final_width = width if col_idx in core_cols else round(width * 0.9, 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = final_width

    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        source_row = (
            int(row["Original Row Index"])
            if "Original Row Index" in row and pd.notna(row["Original Row Index"])
            else r_idx
        )
        img_key = f"{excel_text_str(row['purchase-date'])}_{excel_text_str(row['运单号'])}".lower()
        img_info = image_map.get(img_key)
        if img_info:
            img_stream = io.BytesIO(img_info["bytes"])
            xl_img = XLImage(img_stream)
            source_ws.row_dimensions[source_row].height = row_height_pt
            max_w, max_h = 236, 306
            w, h = float(xl_img.width), float(xl_img.height)
            if w > 0 and h > 0:
                scale = min(max_w / w, max_h / h, 1.0)
                xl_img.width = int(w * scale)
                xl_img.height = int(h * scale)
            source_ws.add_image(xl_img, f"C{source_row}")
            try:
                if hasattr(xl_img.anchor, "_from"):
                    x_offset = max(0, int((250 - xl_img.width) / 2))
                    y_offset = max(0, int((320 - xl_img.height) / 2))
                    xl_img.anchor._from.colOff = int(x_offset * 9525)
                    xl_img.anchor._from.rowOff = int(y_offset * 9525)
            except Exception:
                pass

    for row_idx in range(2, source_ws.max_row + 1):
        for col_idx in text_cols:
            src_cell = source_ws.cell(row=row_idx, column=col_idx)
            src_cell.value = excel_text_str(src_cell.value)
            src_cell.number_format = "@"

    safe_source_sheet = source_sheet_name.replace("'", "''")
    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        source_row = (
            int(row["Original Row Index"])
            if "Original Row Index" in row and pd.notna(row["Original Row Index"])
            else r_idx
        )
        ws.row_dimensions[r_idx].height = row_height_pt
        for c_idx, col_name in enumerate(output_cols, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            value = row[col_name]
            if c_idx in text_cols:
                cell.value = excel_text_str(value)
                cell.number_format = "@"
            elif c_idx == img_col:
                cell.value = f"='{safe_source_sheet}'!C{source_row}"
            else:
                cell.value = value
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = body_font
            cell.border = thin_border
            if c_idx == 9 and qty_gt_one(value):
                cell.fill = PatternFill(fill_type="solid", fgColor="FFFF00")
                cell.font = Font(name=body_font.name, size=body_font.size, bold=True)

    for c_idx in range(1, len(output_cols) + 1):
        hcell = ws.cell(row=1, column=c_idx)
        hcell.font = header_font
        hcell.fill = header_fill
        hcell.alignment = Alignment(horizontal="center", vertical="center")
        hcell.border = thin_border

    d_list = df["序号D"].tolist()
    n = len(d_list)
    start = 0
    while start < n:
        end = start
        while end + 1 < n and d_list[end + 1] == d_list[start]:
            end += 1
        if end > start:
            ws.merge_cells(
                start_row=start + 2,
                start_column=4,
                end_row=end + 2,
                end_column=4,
            )
        start = end + 1

    ws.page_setup.fitToWidth = 0
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale = 48
    ws.print_options.horizontalCentered = True

    wb.save(output)
    return output.getvalue()
