"""
report_generator.py — Wrapper around generate_sprint_reports.py logic
---------------------------------------------------------------------
Adapts the existing script for use inside an Azure Function
(no argparse, no sys.exit, returns list of generated file paths).
"""

import logging
import os
import re
import warnings
from datetime import datetime

import openpyxl

logger = logging.getLogger(__name__)

# ── Import all the core logic from the existing script ───────────────────────
# We put the existing generate_sprint_reports.py alongside this file.
# This module re-exports its functions and adds a clean entry point.

from generate_sprint_reports import (
    NON_CUSTOMER_SHEETS,
    parse_sheet,
    get_project_info,
    build_file1,
    build_file2,
    compute_latest_sprint_rag,
    build_all_projects,
    build_regional_master_index,
    build_heatmap_table,
    _region_group,
    _sanitize_ascii,
    METADATA_DIVIDER,
)


def generate_all_reports(workbook_paths: list[str], output_dir: str) -> list[str]:
    """
    Process multiple Excel workbooks and generate all text report files.

    Args:
        workbook_paths: List of paths to downloaded .xlsm files.
        output_dir:     Directory to write output text files.

    Returns:
        List of absolute paths to generated files.
    """
    os.makedirs(output_dir, exist_ok=True)
    all_project_data = []
    generated_files = []

    for wb_path in workbook_paths:
        if not os.path.exists(wb_path):
            logger.warning("Workbook not found, skipping: %s", wb_path)
            continue

        logger.info("Loading workbook: %s", wb_path)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)

        sheets = [s for s in wb.sheetnames if s not in NON_CUSTOMER_SHEETS]
        logger.info("Processing %d sheets from %s", len(sheets), os.path.basename(wb_path))

        for sheet_name in sheets:
            parsed = parse_sheet(wb[sheet_name])
            if not parsed or not parsed["sprints"]:
                logger.info("  SKIP: %s – no sprint data", sheet_name)
                continue

            info = get_project_info(parsed["metadata"])
            project = info["project"] or sheet_name
            safe_name = re.sub(r'[\\/*?:"<>|]', "_", project)

            # Write File1 – Health Summary
            f1_path = os.path.join(output_dir, f"{safe_name}_File1_Health_Summary.txt")
            with open(f1_path, "w", encoding="utf-8") as f:
                f.write(build_file1(parsed, info))
            generated_files.append(f1_path)

            # Write File2 – Parameter Trends
            f2_path = os.path.join(output_dir, f"{safe_name}_File2_Parameter_Trends.txt")
            with open(f2_path, "w", encoding="utf-8") as f:
                f.write(build_file2(parsed, info))
            generated_files.append(f2_path)

            # Collect for master index
            rag_data = compute_latest_sprint_rag(parsed)
            total_sprints = len(parsed["sprints"])
            total_metrics = len(parsed["metrics"])
            all_project_data.append((info, parsed["metadata"], total_sprints, total_metrics, rag_data))

            logger.info("  OK: %s (%d sprints, %d metrics)", sheet_name, total_sprints, total_metrics)

    # ── Generate index files ─────────────────────────────────────────────────
    if all_project_data:
        # all_projects.txt
        all_path = os.path.join(output_dir, "all_projects.txt")
        with open(all_path, "w", encoding="utf-8") as f:
            f.write(build_all_projects(all_project_data))
        generated_files.append(all_path)

        # Per-region master indexes
        region_buckets = {"Americas": [], "EMEA": [], "Asia": []}
        for entry in all_project_data:
            rg = _region_group(entry[0].get("region", ""))
            if rg in region_buckets:
                region_buckets[rg].append(entry)

        region_keywords = {
            "Americas": "MASTER_INDEX_AMERICAS",
            "EMEA": "MASTER_INDEX_EMEA",
            "Asia": "MASTER_INDEX_ASIA",
        }
        start = 1
        for region_name in ("Americas", "EMEA", "Asia"):
            entries = region_buckets[region_name]
            if not entries:
                continue
            keyword = region_keywords[region_name]
            content = build_regional_master_index(region_name, keyword, entries, start)
            content = _sanitize_ascii(content)
            fpath = os.path.join(output_dir, f"{region_name}_MasterIndex.txt")
            with open(fpath, "w", encoding="utf-8") as f:
                f.write(content)
            generated_files.append(fpath)
            logger.info("  Index: %s (%d projects)", fpath, len(entries))
            start += len(entries)

        # Heatmap
        heatmap_content = build_heatmap_table(all_project_data)
        heatmap_path = os.path.join(output_dir, "latest_sprint_heatmap.txt")
        with open(heatmap_path, "w", encoding="utf-8") as f:
            f.write(heatmap_content)
        generated_files.append(heatmap_path)

    logger.info("Total generated: %d files from %d projects.", len(generated_files), len(all_project_data))
    return generated_files
