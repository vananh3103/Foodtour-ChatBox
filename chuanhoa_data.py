from pathlib import Path
import re
import unicodedata

import pandas as pd


# Files expected in the dataset folder beside this script.
DATA_DIR = Path(__file__).resolve().parent / "dataset"
RESTAURANT_FILE = DATA_DIR / "Food_tour_SG_chuan_hoa_dia_chi_v1.xlsx"
ADMIN_FILE = DATA_DIR / "tphcm-sap-nhap-phuong-xa.xlsx"

# These defaults match the workbooks currently in the dataset folder.
RESTAURANT_SHEET = "Food tour - chuan hoa"
ADMIN_SHEET = "Sheet1"
RESTAURANT_WARD_COLUMN = "Phường cũ nhận diện"
WARD_OUTPUT_COLUMN = "Phường hiện tại (đã đối chiếu)"
ADDRESS_COLUMN = "Địa chỉ chuẩn hóa (tạm thời)"
OLD_WARD_COLUMN = "Phường/Xã cũ"
ADMIN_DISTRICT_COLUMN = "Quận/Huyện cũ"
NEW_WARD_COLUMN = "Phường/Xã mới (từ 1/7/2025)"
ADMIN_PROVINCE_COLUMN = "Tỉnh/TP cũ (trước sáp nhập)"
ADMIN_PROVINCE_VALUE = "Thành phố Hồ Chí Minh"
RESTAURANT_DISTRICT_COLUMN = "Quận/huyện cũ"
OUTPUT_FILE = DATA_DIR / "Food_tour_SG_chuan_hoa_dia_chi.xlsx"
AUDIT_FILE = DATA_DIR / "Food_tour_SG_can_kiem_tra.xlsx"


def inspect_workbooks():
    """Print sheets, headers, and sample rows for the two input workbooks."""
    for path in (RESTAURANT_FILE, ADMIN_FILE):
        if not path.exists():
            raise FileNotFoundError(f"Không tìm thấy file: {path}")

        print(f"\n{'=' * 80}\n{path.name}")
        excel = pd.ExcelFile(path)
        print("Các sheet:", excel.sheet_names)
        for sheet in excel.sheet_names:
            sample = pd.read_excel(path, sheet_name=sheet, nrows=5)
            print(f"\nSheet: {sheet!r}")
            print("Tên cột:", sample.columns.tolist())
            print(sample.head(5).to_string(index=False))


def normalize_key(value):
    if pd.isna(value):
        return ""
    text = unicodedata.normalize("NFD", str(value).strip().lower())
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = text.replace("đ", "d")
    text = re.sub(r"^(phuong|xa|thi tran|quan|huyen|tp|thanh pho)\s+", "", text)
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def clean_unit_label(value):
    return re.sub(
        r"^(phường|xã|thị trấn|quận|huyện|thành phố|tỉnh)\s+",
        "",
        str(value).strip(),
        flags=re.IGNORECASE,
    )


def validate_column(frame, column, label):
    if column is None or column not in frame.columns:
        raise ValueError(
            f"{label} chưa được cấu hình đúng. Điền chính xác tên cột. "
            f"Các cột hiện có: {frame.columns.tolist()}"
        )


def make_location_mapping(frame):
    mapping = {}
    ambiguous = set()
    ward_only = {}
    ward_targets = {}
    for old_name, district, new_name in frame[
        [OLD_WARD_COLUMN, ADMIN_DISTRICT_COLUMN, NEW_WARD_COLUMN]
    ].dropna(subset=[OLD_WARD_COLUMN, NEW_WARD_COLUMN]).itertuples(
        index=False, name=None
    ):
        old_name = str(old_name).strip()
        new_name = str(new_name).strip()
        ward_key = normalize_key(old_name)
        district_key = normalize_key(district)
        if not ward_key or not old_name or not new_name:
            continue
        location_key = (ward_key, district_key)
        if (
            location_key in mapping
            and normalize_key(mapping[location_key][1]) != normalize_key(new_name)
        ):
            ambiguous.add(location_key)
        else:
            mapping[location_key] = (old_name, new_name)
        ward_targets.setdefault(ward_key, set()).add(normalize_key(new_name))
        ward_only.setdefault(ward_key, (old_name, new_name))

    for key in ambiguous:
        mapping.pop(key, None)

    unique_ward_mapping = {
        ward_key: value
        for ward_key, value in ward_only.items()
        if len(ward_targets[ward_key]) == 1
    }
    return mapping, unique_ward_mapping, ambiguous, ward_targets


def replace_name_in_address(value, old_name, new_name):
    if pd.isna(value):
        return value
    result = str(value)
    # Do not replace numeric-only ward names in free text; they may be house numbers.
    if not old_name or old_name.isdigit():
        return result
    pattern = re.compile(
        r"(?<!\w)" + re.escape(old_name) + r"(?!\w)",
        flags=re.IGNORECASE,
    )
    return pattern.sub(new_name, result)


def standardize():
    if not RESTAURANT_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file quán ăn: {RESTAURANT_FILE}")
    if not ADMIN_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy file đơn vị hành chính: {ADMIN_FILE}")

    restaurant_excel = pd.ExcelFile(RESTAURANT_FILE)
    selected_sheet = (
        restaurant_excel.sheet_names[RESTAURANT_SHEET]
        if isinstance(RESTAURANT_SHEET, int)
        else RESTAURANT_SHEET
    )
    if selected_sheet not in restaurant_excel.sheet_names:
        raise ValueError(
            f"Không tìm thấy sheet {selected_sheet!r}. "
            f"Các sheet hiện có: {restaurant_excel.sheet_names}"
        )
    restaurants = pd.read_excel(RESTAURANT_FILE, sheet_name=selected_sheet)
    admin = pd.read_excel(ADMIN_FILE, sheet_name=ADMIN_SHEET)

    validate_column(admin, OLD_WARD_COLUMN, "OLD_WARD_COLUMN (phường/xã cũ)")
    validate_column(admin, NEW_WARD_COLUMN, "NEW_WARD_COLUMN (phường/xã mới)")
    validate_column(admin, ADMIN_DISTRICT_COLUMN, "ADMIN_DISTRICT_COLUMN")
    validate_column(admin, ADMIN_PROVINCE_COLUMN, "ADMIN_PROVINCE_COLUMN")
    validate_column(
        restaurants, RESTAURANT_DISTRICT_COLUMN, "RESTAURANT_DISTRICT_COLUMN"
    )
    if RESTAURANT_WARD_COLUMN is None and ADDRESS_COLUMN is None:
        raise ValueError(
            "Hãy điền RESTAURANT_WARD_COLUMN (cột phường/xã riêng), "
            "ADDRESS_COLUMN (địa chỉ dạng văn bản), hoặc cả hai."
        )
    if RESTAURANT_WARD_COLUMN is not None:
        validate_column(restaurants, RESTAURANT_WARD_COLUMN, "RESTAURANT_WARD_COLUMN")
    if WARD_OUTPUT_COLUMN is not None and WARD_OUTPUT_COLUMN not in restaurants.columns:
        raise ValueError(
            f"WARD_OUTPUT_COLUMN không tồn tại: {WARD_OUTPUT_COLUMN!r}. "
            f"Các cột hiện có: {restaurants.columns.tolist()}"
        )
    if ADDRESS_COLUMN is not None:
        validate_column(restaurants, ADDRESS_COLUMN, "ADDRESS_COLUMN")

    admin = admin[
        admin[ADMIN_PROVINCE_COLUMN].map(normalize_key)
        == normalize_key(ADMIN_PROVINCE_VALUE)
    ]
    if admin.empty:
        raise ValueError(
            f"Không có dòng đối chiếu thuộc {ADMIN_PROVINCE_VALUE!r}; "
            "kiểm tra ADMIN_PROVINCE_VALUE."
        )
    ward_map, ward_only_map, ambiguous_locations, ward_targets = (
        make_location_mapping(admin)
    )
    print("Số cặp phường/xã cũ + quận/huyện bị mơ hồ:", len(ambiguous_locations))
    new_ward_names = {
        normalize_key(name) for name in admin[NEW_WARD_COLUMN].dropna()
    }
    result = restaurants.copy()
    audit = restaurants.copy()
    audit.insert(0, "dong_du_lieu_trong_sheet", restaurants.index + 2)

    def find_location(value, district):
        ward_key = normalize_key(value)
        district_key = normalize_key(district)
        if not ward_key:
            return None
        if district_key:
            return ward_map.get((ward_key, district_key))
        return ward_only_map.get(ward_key)

    if RESTAURANT_WARD_COLUMN is not None:
        original_ward = result[RESTAURANT_WARD_COLUMN].copy()
        original_district = result[RESTAURANT_DISTRICT_COLUMN].copy()
        existing_standardized = (
            result[WARD_OUTPUT_COLUMN].copy()
            if WARD_OUTPUT_COLUMN is not None
            else pd.Series(index=result.index, dtype=object)
        )

        def standardize_ward(value, district, existing_value):
            location = find_location(value, district)
            if location is not None:
                return clean_unit_label(location[1])
            if normalize_key(value) in new_ward_names:
                return clean_unit_label(value)
            if not pd.isna(existing_value) and str(existing_value).strip():
                return existing_value
            return pd.NA

        standardized_ward = pd.Series(
            [
                standardize_ward(old, district, current)
                for old, district, current in zip(
                    original_ward, original_district, existing_standardized
                )
            ],
            index=result.index,
        )
        destination_column = WARD_OUTPUT_COLUMN or RESTAURANT_WARD_COLUMN
        result[destination_column] = standardized_ward
        audit["phuong_xa_sau_chuan_hoa"] = standardized_ward
        audit["khong_khop_phuong_xa"] = [
            not pd.isna(old)
            and normalize_key(old) not in new_ward_names
            and find_location(old, district) is None
            and (pd.isna(current) or not str(current).strip())
            for old, district, current in zip(
                original_ward, original_district, existing_standardized
            )
        ]

    if ADDRESS_COLUMN is not None:
        result[ADDRESS_COLUMN] = [
            replace_name_in_address(
                address,
                (location[0] if location else ""),
                (location[1] if location else ""),
            )
            for address, old, district in zip(
                result[ADDRESS_COLUMN],
                restaurants[RESTAURANT_WARD_COLUMN],
                restaurants[RESTAURANT_DISTRICT_COLUMN],
            )
            for location in [find_location(old, district)]
        ]

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for sheet in restaurant_excel.sheet_names:
            frame = (
                result
                if sheet == selected_sheet
                else pd.read_excel(RESTAURANT_FILE, sheet_name=sheet)
            )
            frame.to_excel(writer, index=False, sheet_name=sheet[:31])

    if "khong_khop_phuong_xa" in audit.columns:
        unmatched = audit[audit["khong_khop_phuong_xa"]]
        with pd.ExcelWriter(AUDIT_FILE, engine="openpyxl") as writer:
            unmatched.to_excel(writer, index=False, sheet_name="Can_kiem_tra")
        print("Số dòng phường/xã chưa khớp:", len(unmatched))
        if not unmatched.empty:
            print("File các dòng cần kiểm tra:", AUDIT_FILE.resolve())

    print("Đã lưu file chuẩn hóa:", OUTPUT_FILE.resolve())
    print("Số dòng dữ liệu:", len(result))
    print(result.head(10).to_string(index=False))


if __name__ == "__main__":
    inspect_workbooks()
    if OLD_WARD_COLUMN is None or NEW_WARD_COLUMN is None:
        print(
            "\nBước tiếp theo: sửa các biến cấu hình ở đầu file theo tên cột vừa in "
            "ra, sau đó chạy lại."
        )
    else:
        standardize()
