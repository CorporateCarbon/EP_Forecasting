# -*- coding: utf-8 -*-
# Merges the Baseline and Project Scenario CEAs for PF Schedule 4.
# Uses shared utilities for header normalization, metric padding, and date parsing.
#%%##
from sch4_merge_utils import merge_workbooks

# ---------- INPUTS ----------
baseline_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\Coalara_Schedule4_NewModel\FC24\combined_output_2025-11-12.xlsx"
project_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\project_combined.xlsx"

output_folder = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\outputs"
output_name_prefix = "fc24_combined_output"

LABEL_A = "Baseline"
LABEL_B = "Project"
DATE_FMT = "%d/%m/%Y"


def main():
    merge_workbooks(
        baseline_path=baseline_path,
        project_path=project_path,
        output_folder=output_folder,
        output_name_prefix=output_name_prefix,
        label_a=LABEL_A,
        label_b=LABEL_B,
        date_fmt=DATE_FMT,
        normalize_month_end=False,
    )


if __name__ == "__main__":
    main()

