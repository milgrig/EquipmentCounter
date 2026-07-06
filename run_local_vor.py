import os, sys, time

# Force UTF-8 output so Cyrillic logs don't crash on Windows cp125x consoles
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

FOLDERS = {
    "eo":  "uploads/2026-06-17_11-39-38/02_PDF",
    "em":  "uploads/2026-06-17_11-34-32/02_PDF",
    "eg":  "uploads/2026-06-17_11-31-14/PDF",
}

def run_one(key, folder):
    from pdf_vor_pipeline import generate_vor_from_pdfs, export_vor_xlsx
    t0 = time.time()
    def L(m):
        line = f"[{time.time()-t0:7.1f}s] {m}"
        try:
            print(line, flush=True)
        except Exception:
            sys.stdout.buffer.write((line + "\n").encode("utf-8", "replace"))
            sys.stdout.flush()
    print(f"\n==================== {key}: {folder} ====================", flush=True)
    if not os.path.isdir(folder):
        print(f"  !! folder missing: {folder}", flush=True)
        return
    try:
        sec = generate_vor_from_pdfs(folder, log=L)
    except Exception as e:
        import traceback
        print("EXCEPTION:", e, flush=True)
        traceback.print_exc()
        return
    total = round(time.time() - t0, 1)
    nrows = sum(len(s.rows) for s in sec)
    print(f">>> RESULT [{key}]: {len(sec)} sections, {nrows} rows, {total}s", flush=True)
    for s in sec:
        print(f"    - {s.title}: {len(s.rows)} rows", flush=True)
    out = f"vor_local_{key}.xlsx"
    try:
        p = export_vor_xlsx(sec, out, project_name=key, log=lambda m: None)
        print(f">>> XLSX: {os.path.abspath(p)}", flush=True)
    except Exception as e:
        import traceback
        print("XLSX FAILED:", e, flush=True)
        traceback.print_exc()

def main():
    keys = sys.argv[1:] or list(FOLDERS.keys())
    grand = time.time()
    for k in keys:
        run_one(k, FOLDERS[k])
    print(f"\n========== ALL DONE {round(time.time()-grand,1)}s ==========", flush=True)

if __name__ == "__main__":
    import multiprocessing as mp
    mp.freeze_support()
    main()
