import io
from typing import List, Optional, Dict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Combine Excel Tool", page_icon="📑", layout="wide")
st.image("combineexcelfile.jpg")
st.title("📑 Combine Excel File ")
st.caption("Gộp dữ liệu Excel nhanh – chọn chế độ, tải lên, và tải về kết quả.")

with st.sidebar:
    st.header("⚙️ Tuỳ chọn")
    mode = st.radio(
        "Chọn chế độ gộp",
        (
            "Gộp TẤT CẢ sheet trong 1 file Excel",
            "Gộp NHIỀU file Excel (mỗi file 1 sheet)",
        ),
    )

    union_type = st.selectbox(
        "Kiểu hợp cột khi khác nhau",
        (
            "Hợp nhất theo TẬP HỢP (outer) – giữ tất cả cột",
            "Giao nhau (inner) – chỉ giữ cột chung",
        ),
        help="Nếu các sheet/file có cột khác nhau: outer giữ tất cả cột (thiếu sẽ là NaN), inner chỉ giữ cột xuất hiện ở tất cả bảng.",
    )
    join_how = "outer" if union_type.startswith("Hợp nhất") else "inner"

    add_source = st.checkbox(
        "Thêm cột nguồn (file/sheet)", value=True,
        help="Gắn cột _source để biết dữ liệu đến từ file/sheet nào."
    )

    preview_rows = st.number_input(
        "Số dòng xem trước", min_value=5, max_value=200, value=20, step=5
    )

    st.markdown("---")
    st.markdown(
        "👇THAM KHẢO THÊM CÁC TOOL HỮU ÍCH KHÁC!"
    )
    st.markdown("[Công Cụ Hữu Ích Miễn Phí](https://www.bpndgroup.com/cong-cu-mien-phi)")
    st.markdown(
        "👇LINK THAM GIA NHÓM ZALO MIỄN PHÍ"
    )
    st.markdown("[Nhóm AI Dữ Liệu Thực Chiến](https://zalo.me/g/lkouhv397)")
    st.markdown("[Nhóm Supply Chain Analysis](https://zalo.me/g/zxznwg212)")
    st.markdown(
        "👇THAM KHẢO THÊM CÁC KHÓA HỌC AI - DỮ LIỆU - SUPPLY CHAIN!"
    )
    
    st.markdown("[Khóa Học Đào Tạo Online Trực Tiếp](https://www.bpndgroup.com/djao-tao-ai-du-lieu)")
    st.markdown("[Khóa Học E-Learning Video](https://khoahoc.bpndgroup.com/)")
    st.image("founder.jpg",caption="Bản quyền bpndgroup.com - Lê Văn Đức AI Data Trainer")

def _safe_read_excel(file, sheet: Optional[str | int] = None) -> pd.DataFrame:
    """Đọc 1 sheet từ một đối tượng file-like của Streamlit.
    Trả về DataFrame; raise Exception nếu lỗi."""
    # Lưu vào buffer để có thể đọc nhiều lần nếu cần
    data = file.read()  
    bio = io.BytesIO(data)
    # pandas sẽ tự chọn engine phù hợp (openpyxl/xlrd)
    df = pd.read_excel(bio, sheet_name=sheet)
    # Đảm bảo reset pointer để dùng lại nếu cần
    file.seek(0)
    return df


def _concat_with_how(dfs: List[pd.DataFrame], how: str) -> pd.DataFrame:
    if not dfs:
        return pd.DataFrame()
    # Với inner: align cột chung
    if how == "inner":
        common_cols = set(dfs[0].columns)
        for d in dfs[1:]:
            common_cols &= set(d.columns)
        dfs = [d[list(common_cols)] for d in dfs]
    # pandas concat sẽ xử lý outer khi cột khác nhau
    return pd.concat(dfs, ignore_index=True, sort=False)


if mode == "Gộp TẤT CẢ sheet trong 1 file Excel":
    up = st.file_uploader(
        "Tải lên 1 file Excel", type=["xlsx", "xls"], accept_multiple_files=False
    )

    if up is not None:
        try:
            # Đọc tất cả sheet: dict[sheet_name -> DataFrame]
            up_bytes = io.BytesIO(up.read())
            up.seek(0)
            all_sheets: Dict[str, pd.DataFrame] = pd.read_excel(up_bytes, sheet_name=None)

            dfs: List[pd.DataFrame] = []
            for sheet_name, dfx in all_sheets.items():
                df = dfx.copy()
                if add_source:
                    df["_source_file"] = up.name
                    df["_source_sheet"] = sheet_name
                dfs.append(df)

            combined = _concat_with_how(dfs, join_how)

            st.success(f"Đã gộp {len(dfs)} sheet từ file: {up.name}")
            st.dataframe(combined.head(int(preview_rows)))

            # Tải về CSV
            csv_data = combined.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Tải về CSV",
                data=csv_data,
                file_name="combined.csv",
                mime="text/csv",
            )

            # Tải về Excel
            xlsx_buf = io.BytesIO()
            with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
                combined.to_excel(writer, index=False, sheet_name="combined")
            st.download_button(
                "⬇️ Tải về Excel",
                data=xlsx_buf.getvalue(),
                file_name="combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Không đọc được file: {e}")

else:  # Gộp NHIỀU file Excel (mỗi file 1 sheet)
    ups = st.file_uploader(
        "Tải lên NHIỀU file Excel", type=["xlsx", "xls"], accept_multiple_files=True
    )

    sheet_hint = st.text_input(
        "Tên sheet (tuỳ chọn, áp dụng cho TẤT CẢ file)",
        value="",
        placeholder="Để trống = sheet đầu tiên",
        help="Nếu nhập, chương trình sẽ đọc sheet này từ mỗi file. Nếu để trống, sẽ đọc sheet đầu tiên."
    )

    if ups:
        try:
            dfs: List[pd.DataFrame] = []
            for f in ups:
                # Mỗi file: 1 sheet – theo tên nhập, hoặc sheet đầu tiên (index 0)
                sheet_to_read: Optional[str | int] = sheet_hint if sheet_hint else 0
                df = _safe_read_excel(f, sheet=sheet_to_read)
                # Nếu người dùng nhập tên sheet không tồn tại và pandas trả về dict -> xử lý
                if isinstance(df, dict):
                    # Khi sheet=None sẽ trả về dict; nhưng ta không dùng case này ở đây
                    # Bảo vệ: chọn sheet đầu tiên
                    first_name = list(df.keys())[0]
                    df = df[first_name]
                if add_source:
                    df = df.copy()
                    df["_source_file"] = f.name
                    df["_source_sheet"] = sheet_to_read if sheet_hint else "<first>"
                dfs.append(df)

            combined = _concat_with_how(dfs, join_how)

            st.success(f"Đã gộp {len(dfs)} file")
            st.dataframe(combined.head(int(preview_rows)))

            # Tải về CSV
            csv_data = combined.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Tải về CSV",
                data=csv_data,
                file_name="combined.csv",
                mime="text/csv",
            )

            # Tải về Excel
            xlsx_buf = io.BytesIO()
            with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
                combined.to_excel(writer, index=False, sheet_name="combined")
            st.download_button(
                "⬇️ Tải về Excel",
                data=xlsx_buf.getvalue(),
                file_name="combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Có lỗi khi gộp file: {e}")

st.markdown("---")
st.subheader("🧭 Cách chạy")
st.code(
    """
    # 1) Chọn chế độ gộp

    # 2) Load một file excel hoặc nhiều file excel lên

    # 3) Chọn tải về Excel/CSV
  
    """,
    language="bash",
)

st.info(
    "Lưu ý: Công cụ xử lý và trả file kết quả về máy của bạn."
    " Nên hoàn toàn bảo mật data cho bạn/công ty của bạn nhé!")
