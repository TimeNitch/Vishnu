from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Data_Vishnu.xlsx"

st.set_page_config(
    page_title="Vishnu Database",
    page_icon="🔎",
    layout="wide",
)

st.title("Vishnu Database")

if not EXCEL_FILE.exists():
    st.error(f"ไม่พบไฟล์ {EXCEL_FILE.name}")
    st.stop()

try:
    data = pd.read_excel(EXCEL_FILE, sheet_name="Data")
except Exception as error:
    st.error(f"อ่านไฟล์ Excel ไม่สำเร็จ: {error}")
    st.stop()

data = data.fillna("")


def available_values(column_name: str) -> list[str]:
    if column_name not in data.columns:
        return []

    return sorted(
        {
            str(value).strip()
            for value in data[column_name]
            if str(value).strip()
        }
    )


st.sidebar.header("ตัวกรอง")

search_text = st.sidebar.text_input(
    "ค้นหาทุกคอลัมน์",
    placeholder="ชื่อเล่น โรงเรียน จังหวัด...",
)

province_filter = st.sidebar.multiselect(
    "จังหวัด",
    available_values("จังหวัด"),
)

school_filter = st.sidebar.multiselect(
    "โรงเรียนเดิม",
    available_values("โรงเรียนเดิม"),
)

study_filter = st.sidebar.multiselect(
    "สายการเรียน",
    available_values("สายการเรียน"),
)

personality_filter = st.sidebar.multiselect(
    "แท็กบุคลิก",
    available_values("แท็กบุคลิก"),
)

filtered_data = data.copy()

if search_text:
    search_text = search_text.casefold()

    search_mask = filtered_data.astype(str).apply(
        lambda row: row.str.casefold().str.contains(
            search_text,
            regex=False,
        ).any(),
        axis=1,
    )

    filtered_data = filtered_data[search_mask]

if province_filter and "จังหวัด" in filtered_data.columns:
    filtered_data = filtered_data[
        filtered_data["จังหวัด"].isin(province_filter)
    ]

if school_filter and "โรงเรียนเดิม" in filtered_data.columns:
    filtered_data = filtered_data[
        filtered_data["โรงเรียนเดิม"].isin(school_filter)
    ]

if study_filter and "สายการเรียน" in filtered_data.columns:
    filtered_data = filtered_data[
        filtered_data["สายการเรียน"].isin(study_filter)
    ]

if personality_filter and "แท็กบุคลิก" in filtered_data.columns:
    filtered_data = filtered_data[
        filtered_data["แท็กบุคลิก"].isin(personality_filter)
    ]

st.write(f"พบข้อมูล **{len(filtered_data)} คน**")

for _, person in filtered_data.iterrows():
    profile_url = str(
        person.get("Profile Picture URL", "")
    ).strip()

    with st.container(border=True):
        image_column, detail_column = st.columns([1, 3])

        with image_column:
            if profile_url.startswith(("http://", "https://")):
                st.image(profile_url, use_container_width=True)
            else:
                st.info("ไม่มีรูปโปรไฟล์")

        with detail_column:
            st.subheader(
                str(person.get("ชื่อเล่น", "ไม่ระบุชื่อ"))
            )

            st.write(
                f"**ข้อ:** {person.get('Question', '-')}"
            )
            st.write(
                f"**จังหวัด:** {person.get('จังหวัด', '-')}"
            )
            st.write(
                f"**โรงเรียนเดิม:** "
                f"{person.get('โรงเรียนเดิม', '-')}"
            )
            st.write(
                f"**สายการเรียน:** "
                f"{person.get('สายการเรียน', '-')}"
            )
            st.write(
                f"**สิ่งที่ชอบ:** "
                f"{person.get('สิ่งที่ชอบ', '-')}"
            )
            st.write(
                f"**ความสามารถ/งานอดิเรก:** "
                f"{person.get('ความสามารถ/งานอดิเรก', '-')}"
            )
            st.write(
                f"**แท็กบุคลิก:** "
                f"{person.get('แท็กบุคลิก', '-')}"
            )