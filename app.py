import os
import shutil

import pandas as pd
import streamlit as st

from engines.yesbank.yesbank_main import (
    generate_yesbank_report
)


st.set_page_config(
    page_title="YesBank Report Generator",
    layout="centered"
)

st.title(
    "YesBank Audit Report Generator"
)

st.markdown(
    "Generate YesBank Audit Reports using the monthly data files."
)


os.makedirs(
    "temp",
    exist_ok=True
)

os.makedirs(
    "output",
    exist_ok=True
)


def save_uploaded_file(
    uploaded_file,
    destination
):

    with open(
        destination,
        "wb"
    ) as f:

        f.write(
            uploaded_file.getbuffer()
        )


def get_agency_list(
    uploaded_file
):

    df = pd.read_excel(
        uploaded_file,
        sheet_name="Schedule",
        keep_default_na=False
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    agencies = (
        df.iloc[:, 7]
        .astype(str)
        .str.strip()
    )

    uploaded_file.seek(0)

    return sorted(
        agencies[
            agencies != ""
        ].unique()
    )


REPORT_TYPE = st.selectbox(

    "Select Template",

    [

        "Collection",

        "Repo",

        "Stockyard"

    ]

)


template_map = {

    "Collection":

        "templates/YesBank/YesBank_Collection_Template_2026-27.xlsx",

    "Repo":

        "templates/YesBank/YesBank_Repossesion_Template_2026-27.xlsx",

    "Stockyard":

        "templates/YesBank/YesBank_Stockyard_Template_2026-27.xlsx"

}


TEMPLATE_FILE = template_map[
    REPORT_TYPE
]


st.divider()

st.subheader(
    "Upload Required Files"
)


base_data_file = st.file_uploader(

    "Base Data.xlsx",

    type=[
        "xlsx"
    ]

)


agency_name = None


if base_data_file is not None:

    try:

        agency_name = st.selectbox(

            "Select Agency",

            get_agency_list(
                base_data_file
            )

        )

    except Exception as e:

        st.error(
            f"Unable to read Schedule sheet.\n\n{e}"
        )


kaf_file = st.file_uploader(

    "KAF.xlsx",

    type=[
        "xlsx"
    ]

)


annexure_file = None

vehicle_details_file = None


if REPORT_TYPE == "Collection":

    annexure_file = st.file_uploader(

        "Annexure.xlsx",

        type=[
            "xlsx"
        ]

    )


if REPORT_TYPE == "Stockyard":

    vehicle_details_file = st.file_uploader(

        "Vehicle Details.xlsx",

        type=[
            "xlsx"
        ]

    )


st.divider()

if st.button(
    "Generate Report",
    use_container_width=True
):

    if base_data_file is None:

        st.error(
            "Please upload Base Data.xlsx."
        )

        st.stop()

    if agency_name is None:

        st.error(
            "Please select an Agency."
        )

        st.stop()

    if kaf_file is None:

        st.error(
            "Please upload KAF.xlsx."
        )

        st.stop()

    if (
        REPORT_TYPE == "Collection"
        and
        annexure_file is None
    ):

        st.error(
            "Please upload Annexure.xlsx."
        )

        st.stop()

    if (
        REPORT_TYPE == "Stockyard"
        and
        vehicle_details_file is None
    ):

        st.error(
            "Please upload Vehicle Details.xlsx."
        )

        st.stop()

    shutil.rmtree(
        "temp",
        ignore_errors=True
    )

    os.makedirs(
        "temp",
        exist_ok=True
    )

    base_path = os.path.join(
        "temp",
        "Base_Data.xlsx"
    )

    kaf_path = os.path.join(
        "temp",
        "KAF.xlsx"
    )

    save_uploaded_file(
        base_data_file,
        base_path
    )

    save_uploaded_file(
        kaf_file,
        kaf_path
    )

    annexure_path = None

    vehicle_path = None

    if REPORT_TYPE == "Collection":

        annexure_path = os.path.join(
            "temp",
            "Annexure.xlsx"
        )

        save_uploaded_file(
            annexure_file,
            annexure_path
        )

    if REPORT_TYPE == "Stockyard":

        vehicle_path = os.path.join(
            "temp",
            "Vehicle Details.xlsx"
        )

        save_uploaded_file(
            vehicle_details_file,
            vehicle_path
        )

    try:

        with st.spinner(
            "Generating Report..."
        ):

            output_file = generate_yesbank_report(

                report_type=REPORT_TYPE,

                base_data_file=base_path,

                kaf_file=kaf_path,

                annexure_file=annexure_path,

                vehicle_details_file=vehicle_path,

                template_file=TEMPLATE_FILE,

                agency_name=agency_name

            )

        st.success(
            "Report Generated Successfully!"
        )

        with open(
            output_file,
            "rb"
        ) as file:

            st.download_button(

                label="📥 Download Report",

                data=file,

                file_name=os.path.basename(
                    output_file
                ),

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                use_container_width=True

            )

    except Exception as e:

        st.exception(
            e
        )