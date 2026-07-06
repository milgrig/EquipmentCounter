# -*- coding: utf-8 -*-
"""Run the no-СО pipeline LOCALLY on one folder and write a UTF-8 log.

Spawn-safe: all work is inside the __main__ guard so ProcessPool workers
that re-import __main__ do NOT re-run the pipeline.

Usage:
    python run_local_folder.py <folder> [out_log]
"""
import multiprocessing as mp


def main():
    import sys, os, time, io

    here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, here)

    folder = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(here, "uploads", "2026-06-17_11-31-14", "PDF")
    out_log = sys.argv[2] if len(sys.argv) > 2 else \
        os.path.join(here, "run_local_folder.log")

    logf = io.open(out_log, "w", encoding="utf-8")

    def log(msg=""):
        line = str(msg)
        logf.write(line + "\n")
        logf.flush()

    from pdf_vor_pipeline import generate_vor_from_pdfs

    log("=== RUN folder: %s ===" % folder)
    t0 = time.time()
    sections = generate_vor_from_pdfs(folder, log=log)
    dt = time.time() - t0

    log("")
    log("=== SUMMARY ===")
    log("ELAPSED_SEC=%.1f" % dt)
    n_sec = len(sections) if sections else 0
    log("SECTIONS=%d" % n_sec)
    total_rows = 0
    for i, sec in enumerate(sections or []):
        title = getattr(sec, "title", getattr(sec, "name", "?"))
        rows = getattr(sec, "rows", None)
        if rows is None:
            rows = getattr(sec, "items", []) or []
        total_rows += len(rows)
        log("SEC[%d] rows=%d title=%s" % (i, len(rows), title))
    log("TOTAL_ROWS=%d" % total_rows)
    logf.close()

    # ASCII-only progress to console
    print("ELAPSED_SEC=%.1f" % dt)
    print("SECTIONS=%d TOTAL_ROWS=%d" % (n_sec, total_rows))
    print("LOG=%s" % out_log)
    print("DONE")


if __name__ == "__main__":
    mp.freeze_support()
    main()
