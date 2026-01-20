# -*- coding: utf-8 -*-
# Combine two multi-sheet Excel workbooks (Baseline vs Project) into one workbook.
# Uses shared utilities for header normalization, metric padding, and date parsing.
#%%##
from sch4_merge_utils import merge_workbooks

# ---------- INPUTS ----------
baseline_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\baseline_combined.xlsx"
project_path = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\project_combined.xlsx"

output_folder = r"C:\Users\GeorginaDoyle\github\EP_FullCAM\outputs"
output_name_prefix = "combined_output"

LABEL_A = "Baseline"
LABEL_B = "Project"
DATE_FMT = "%d/%m/%Y"
NORMALIZE_MONTH_END = False  # set True if you want month-end keys


def main():
    merge_workbooks(
        baseline_path=baseline_path,
        project_path=project_path,
        output_folder=output_folder,
        output_name_prefix=output_name_prefix,
        label_a=LABEL_A,
        label_b=LABEL_B,
        date_fmt=DATE_FMT,
        normalize_month_end=NORMALIZE_MONTH_END,
    )


if __name__ == "__main__":
    main()

