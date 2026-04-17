#!/usr/bin/env python3
"""Wrapper script to run batch_equipment.py with UTF-8 output."""
import sys
import io
import time
import json
from pathlib import Path

# Force UTF-8 output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from batch_equipment import run_pipeline

def run_folder(folder_path: str, log_path: str):
    folder = Path(folder_path).resolve()
    log_file = open(log_path, 'w', encoding='utf-8')

    def dual_log(msg):
        print(msg)
        log_file.write(msg + '\n')
        log_file.flush()

    t0 = time.time()
    try:
        report = run_pipeline(
            folder,
            convert_dwg=True,
            keep_converted=True,
            parse_dxf=True,
            parse_pdf=True,
            generate_png=True,
            png_dpi=200,
            output_json='equipment_report.json',
            log=dual_log,
        )

        # Try VOR generation
        try:
            from vor_generator import generate_vor
            generate_vor(folder, output_path=None)
            dual_log("  VOR: generated successfully")
        except ImportError as e:
            dual_log(f"  VOR skipped (missing dependency): {e}")
        except Exception as e:
            dual_log(f"  VOR error: {e}")

    except Exception as e:
        dual_log(f"\nFATAL ERROR: {e}")
        import traceback
        dual_log(traceback.format_exc())
        report = {}

    elapsed = time.time() - t0
    dual_log(f"\nTime elapsed: {elapsed:.1f}s")

    # Summary
    if report and 'files' in report:
        files_processed = report.get('files_processed', len(report['files']))
        files_total = report.get('files_total', files_processed)
        total_equip = sum(
            fdata.get('total_count', 0) for fdata in report['files'].values()
        )
        total_cable = sum(
            fdata.get('total_cable_length_m', 0) for fdata in report['files'].values()
        )
        dual_log(f"\nSUMMARY: {files_processed}/{files_total} files, {total_equip} equipment, {total_cable}m cables")

    log_file.close()
    return report

if __name__ == '__main__':
    folder = sys.argv[1]
    log_path = sys.argv[2] if len(sys.argv) > 2 else 'pipeline_log.txt'
    run_folder(folder, log_path)
