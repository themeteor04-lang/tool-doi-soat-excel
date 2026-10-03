import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path
import re


# =========================================================
# CẤU HÌNH
# =========================================================
st.set_page_config(
    page_title="Tool đối soát thiếu hàng",
    page_icon="📊",
    layout="wide"
)

st.title("📊 TOOL ĐỐI SOÁT THIẾU HÀNG")

st.write(
    "Đối chiếu số lượng giữa File gốc và File so sánh theo Part Number."
)

st.divider()


# =========================================================
# CHUẨN HÓA MÃ
# =========================================================
def normalize_code(value):

    if pd.isna(value):
        return ""

    if isinstance(value, (int, np.integer)):
        return str(value).strip()

    if isinstance(value, (float, np.floating)):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()

    value = str(value).strip()

    if re.fullmatch(r"-?\d+\.0+", value):
        return value.split(".")[0]

    return value


# =========================================================
# ĐỌC FILE
# =========================================================
def load_file(uploaded_file):

    if uploaded_file is None:
        raise ValueError("Chưa chọn file.")

    extension = Path(
        uploaded_file.name
    ).suffix.lower()

    if extension not in [".xlsx", ".xls"]:
        raise ValueError(
            "Chỉ hỗ trợ file .xlsx hoặc .xls"
        )

    uploaded_file.seek(0)

    try:

        if extension == ".xlsx":
            df = pd.read_excel(
                uploaded_file,
                engine="openpyxl"
            )
        else:
            df = pd.read_excel(
                uploaded_file,
                engine="xlrd"
            )

        df = df.dropna(how="all")

        df.columns = [
            str(col).strip()
            for col in df.columns
        ]

        if df.empty:
            raise ValueError(
                "File không có dữ liệu."
            )

        return df

    except Exception as e:

        raise ValueError(
            f"Không đọc được file: {e}"
        )


# =========================================================
# CHUYỂN SỐ LƯỢNG
# =========================================================
def convert_quantity(series):

    original = series.copy()

    converted = pd.to_numeric(
        series,
        errors="coerce"
    )

    invalid_mask = (
        converted.isna()
        & original.notna()
        & (
            original.astype(str)
            .str.strip() != ""
        )
    )

    invalid_count = int(
        invalid_mask.sum()
    )

    converted = converted.fillna(0)

    return converted, invalid_count


# =========================================================
# LẤY GIÁ TRỊ ĐẦU TIÊN
# =========================================================
def first_valid(series):

    for value in series:

        if pd.notna(value):

            text = str(value).strip()

            if text != "":
                return value

    return ""


# =========================================================
# SO SÁNH
# =========================================================
def compare_data(
    df_original,
    df_compare,
    page_col_original,
    part_col_original,
    weight_col_original,
    class_col_original,
    qty_col_original,
    part_col_compare,
    qty_col_compare
):

    original = df_original.copy()
    compare = df_compare.copy()

    # ==========================
    # FILE GỐC
    # ==========================
    original["_PART_"] = (
        original[part_col_original]
        .apply(normalize_code)
    )

    original = original[
        original["_PART_"] != ""
    ].copy()

    original["_QTY_"], invalid_original = (
        convert_quantity(
            original[qty_col_original]
        )
    )

    # ==========================
    # FILE SO SÁNH
    # ==========================
    compare["_PART_"] = (
        compare[part_col_compare]
        .apply(normalize_code)
    )

    compare = compare[
        compare["_PART_"] != ""
    ].copy()

    compare["_QTY_"], invalid_compare = (
        convert_quantity(
            compare[qty_col_compare]
        )
    )

    # ==========================
    # GROUP FILE GỐC
    # ==========================
    original_group = (
        original
        .groupby(
            "_PART_",
            as_index=False
        )
        .agg(
            {
                page_col_original:
                    first_valid,

                weight_col_original:
                    first_valid,

                class_col_original:
                    first_valid,

                "_QTY_":
                    "sum"
            }
        )
    )

    # ==========================
    # GROUP FILE SO SÁNH
    # ==========================
    compare_group = (
        compare
        .groupby(
            "_PART_",
            as_index=False
        )["_QTY_"]
        .sum()
    )

    compare_dict = dict(
        zip(
            compare_group["_PART_"],
            compare_group["_QTY_"]
        )
    )

    # ==========================
    # TÍNH HÀNG THIẾU
    # ==========================
    result = []

    for _, row in original_group.iterrows():

        part_number = row["_PART_"]

        qty_original = row["_QTY_"]

        qty_compare = compare_dict.get(
            part_number,
            0
        )

        qty_missing = (
            qty_original
            - qty_compare
        )

        if qty_missing > 0:

            result.append(
                {
                    "Trang":
                        row[
                            page_col_original
                        ],

                    "Part Number":
                        part_number,

                    "Weight (Kg)":
                        row[
                            weight_col_original
                        ],

                    "Classification":
                        row[
                            class_col_original
                        ],

                    "SL":
                        qty_missing
                }
            )

    result_df = pd.DataFrame(
        result,
        columns=[
            "Trang",
            "Part Number",
            "Weight (Kg)",
            "Classification",
            "SL"
        ]
    )

    if not result_df.empty:

        result_df["SL"] = (
            result_df["SL"]
            .apply(
                lambda x:
                int(x)
                if float(x).is_integer()
                else x
            )
        )

    warnings = {
        "original":
            invalid_original,

        "compare":
            invalid_compare
    }

    return result_df, warnings


# =========================================================
# XUẤT EXCEL
# =========================================================
def export_excel(result_df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        result_df.to_excel(
            writer,
            index=False,
            sheet_name="Hang thieu"
        )

        ws = writer.book[
            "Hang thieu"
        ]

        ws.freeze_panes = "A2"

        ws.auto_filter.ref = (
            ws.dimensions
        )

        ws.column_dimensions[
            "A"
        ].width = 15

        ws.column_dimensions[
            "B"
        ].width = 25

        ws.column_dimensions[
            "C"
        ].width = 18

        ws.column_dimensions[
            "D"
        ].width = 25

        ws.column_dimensions[
            "E"
        ].width = 12

    output.seek(0)

    return output.getvalue()


# =========================================================
# HÀM TỰ CHỌN CỘT
# =========================================================
def find_default(
    columns,
    name
):

    name_lower = (
        name
        .strip()
        .lower()
    )

    for i, col in enumerate(
        columns
    ):

        if (
            str(col)
            .strip()
            .lower()
            == name_lower
        ):
            return i + 1

    return 0


# =========================================================
# UPLOAD FILE
# =========================================================
col1, col2 = st.columns(2)


with col1:

    st.subheader(
        "1. File gốc"
    )

    file_original = (
        st.file_uploader(
            "Chọn File gốc",
            type=[
                "xlsx",
                "xls"
            ],
            key="file_original"
        )
    )


with col2:

    st.subheader(
        "2. File so sánh"
    )

    file_compare = (
        st.file_uploader(
            "Chọn File so sánh",
            type=[
                "xlsx",
                "xls"
            ],
            key="file_compare"
        )
    )


# =========================================================
# ĐỌC FILE SAU KHI UPLOAD
# =========================================================
df_original = None
df_compare = None


if file_original is not None:

    try:

        df_original = load_file(
            file_original
        )

        st.success(
            f"File gốc: "
            f"{len(df_original)} dòng"
        )

    except Exception as e:

        st.error(
            str(e)
        )


if file_compare is not None:

    try:

        df_compare = load_file(
            file_compare
        )

        st.success(
            f"File so sánh: "
            f"{len(df_compare)} dòng"
        )

    except Exception as e:

        st.error(
            str(e)
        )


# =========================================================
# CHỈ HIỆN CHỌN CỘT KHI ĐÃ CÓ CẢ 2 FILE
# =========================================================
if (
    df_original is not None
    and
    df_compare is not None
):

    st.divider()

    st.header(
        "3. Chọn các cột dữ liệu"
    )

    original_columns = list(
        df_original.columns
    )

    compare_columns = list(
        df_compare.columns
    )

    left, right = st.columns(2)


    # =====================================================
    # FILE GỐC
    # =====================================================
    with left:

        st.subheader(
            "File gốc"
        )

        page_col_original = (
            st.selectbox(
                "Trang",
                options=[
                    "-- Chọn cột --"
                ] + original_columns,

                index=find_default(
                    original_columns,
                    "Trang"
                ),

                key="page_original"
            )
        )

        part_col_original = (
            st.selectbox(
                "Part Number",
                options=[
                    "-- Chọn cột --"
                ] + original_columns,

                index=find_default(
                    original_columns,
                    "Part Number"
                ),

                key="part_original"
            )
        )

        weight_col_original = (
            st.selectbox(
                "Weight (Kg)",
                options=[
                    "-- Chọn cột --"
                ] + original_columns,

                index=find_default(
                    original_columns,
                    "Weight (Kg)"
                ),

                key="weight_original"
            )
        )

        class_col_original = (
            st.selectbox(
                "Classification",
                options=[
                    "-- Chọn cột --"
                ] + original_columns,

                index=find_default(
                    original_columns,
                    "Classification"
                ),

                key="class_original"
            )
        )

        qty_col_original = (
            st.selectbox(
                "SL",
                options=[
                    "-- Chọn cột --"
                ] + original_columns,

                index=find_default(
                    original_columns,
                    "SL"
                ),

                key="qty_original"
            )
        )


    # =====================================================
    # FILE SO SÁNH
    # =====================================================
    with right:

        st.subheader(
            "File so sánh"
        )

        page_col_compare = (
            st.selectbox(
                "Trang",
                options=[
                    "-- Chọn cột --"
                ] + compare_columns,

                index=find_default(
                    compare_columns,
                    "Trang"
                ),

                key="page_compare"
            )
        )

        part_col_compare = (
            st.selectbox(
                "Part Number",
                options=[
                    "-- Chọn cột --"
                ] + compare_columns,

                index=find_default(
                    compare_columns,
                    "Part Number"
                ),

                key="part_compare"
            )
        )

        weight_col_compare = (
            st.selectbox(
                "Weight (Kg)",
                options=[
                    "-- Chọn cột --"
                ] + compare_columns,

                index=find_default(
                    compare_columns,
                    "Weight (Kg)"
                ),

                key="weight_compare"
            )
        )

        class_col_compare = (
            st.selectbox(
                "Classification",
                options=[
                    "-- Chọn cột --"
                ] + compare_columns,

                index=find_default(
                    compare_columns,
                    "Classification"
                ),

                key="class_compare"
            )
        )

        qty_col_compare = (
            st.selectbox(
                "SL",
                options=[
                    "-- Chọn cột --"
                ] + compare_columns,

                index=find_default(
                    compare_columns,
                    "SL"
                ),

                key="qty_compare"
            )
        )


    st.divider()


    # =====================================================
    # NÚT SO SÁNH
    # =====================================================
    if st.button(
        "🔎 SO SÁNH",
        type="primary",
        use_container_width=True
    ):

        required = [
            page_col_original,
            part_col_original,
            weight_col_original,
            class_col_original,
            qty_col_original,
            part_col_compare,
            qty_col_compare
        ]

        if (
            "-- Chọn cột --"
            in required
        ):

            st.error(
                "Bạn chưa chọn đầy đủ các cột."
            )

        else:

            try:

                result_df, warnings = (
                    compare_data(
                        df_original,
                        df_compare,

                        page_col_original,
                        part_col_original,
                        weight_col_original,
                        class_col_original,
                        qty_col_original,

                        part_col_compare,
                        qty_col_compare
                    )
                )

                st.session_state[
                    "result"
                ] = result_df


                if (
                    warnings[
                        "original"
                    ] > 0
                ):

                    st.warning(
                        f"File gốc có "
                        f"{warnings['original']} "
                        f"giá trị SL không hợp lệ. "
                        f"Tool coi các giá trị đó là 0."
                    )


                if (
                    warnings[
                        "compare"
                    ] > 0
                ):

                    st.warning(
                        f"File so sánh có "
                        f"{warnings['compare']} "
                        f"giá trị SL không hợp lệ. "
                        f"Tool coi các giá trị đó là 0."
                    )


                if result_df.empty:

                    st.success(
                        "✅ Không có hàng thiếu."
                    )

                else:

                    st.success(
                        f"✅ Phát hiện "
                        f"{len(result_df)} "
                        f"Part Number đang thiếu."
                    )


            except Exception as e:

                st.error(
                    f"Lỗi khi so sánh: {e}"
                )


# =========================================================
# HIỂN THỊ KẾT QUẢ
# =========================================================
if "result" in st.session_state:

    result_df = (
        st.session_state[
            "result"
        ]
    )

    if not result_df.empty:

        st.divider()

        st.header(
            "4. Kết quả hàng thiếu"
        )

        st.dataframe(
            result_df,
            use_container_width=True,
            hide_index=True
        )

        total_missing = (
            result_df[
                "SL"
            ].sum()
        )

        st.metric(
            "Tổng SL thiếu",
            total_missing
        )

        excel_data = (
            export_excel(
                result_df
            )
        )

        st.download_button(
            label="📥 XUẤT EXCEL",

            data=excel_data,

            file_name=(
                "ket_qua_thieu_hang.xlsx"
            ),

            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),

            type="primary",

            use_container_width=True
        )