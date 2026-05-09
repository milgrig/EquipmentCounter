"""T047 ID-collision recovery:

Two tasks share id='T047' in tasks.json:
  (a) S005 executor=tester, status=questions (pre-existing)
  (b) S010 executor=boss,   status=new       (new diag-script fix)

task_manager.py result/status/get all match on `id`, picking the first
record. My earlier `result T047 "<diag fix text>"` wrote the diag-fix
text to record (a), clobbering the prior tester-baseline-mismatch
notes. Record (b) is still empty.

This script:
  1. Restores (a).result to the recovered S005/T047 halt notes.
  2. Writes (b).result to the diag-fix text and sets (b).status=done.
  3. Preserves all other fields.

We discriminate by (id, sprint_id, executor).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from datetime import datetime

TASKS = Path(".tayfa/common/tasks.json")
data = json.loads(TASKS.read_text(encoding="utf-8"))

S005_RESULT = (
    "[BOSS NOTE — reconstructed after T047 ID collision on 2026-04-24] "
    "T047 was paused in status=questions by boss after tester reported a "
    "BLOCKING baseline-mismatch issue: sprint/S005 branch is based on v0.14.0 "
    "(predates S003 release), benchmark numbers therefore regress vs the S003 "
    "v0.16.0 baseline. Root cause traced by boss via git diagnostics: release "
    "pipeline is partially broken (KB-006/KB-013 class) — S003 never merged to "
    "main, no v0.15.0/v0.16.0/v0.17.0 tags, sprint/S005 has 0 commits beyond "
    "v0.14.0 while 820 lines of CV work remain uncommitted in the working "
    "tree. Three repair tasks were filed and T047 blocked on them: "
    "T048 (developer_pdf release repair), T049 (developer_cv uncommitted "
    "triage), T050 (boss rebase plan). T047 will be re-triggered after T050 "
    "completes; gate text was also corrected (KB-012: no pytest-collectable "
    "suite — run benchmark scripts directly and quote exit codes). "
    "Depends_on: [B025, B026, B027, T048, T049, T050]."
)

S010_RESULT = (
    "Fixed _diag_005.py importlib crash (T047, S010). "
    "Root cause: spec_from_file_location + module_from_spec + exec_module "
    "pattern did not register the module in sys.modules before exec, so "
    "@dataclass decorators in test_vor_cross_format.py failed at import-time "
    "with AttributeError: 'NoneType' object has no attribute '__dict__' in "
    "dataclasses._is_type (sys.modules.get(cls.__module__) returned None). "
    "Fix: added `sys.modules['tvcf'] = tvcf` between module_from_spec and "
    "exec_module (3 lines including explanatory comment). "
    "Verified end-to-end: script now completes; [FINAL] block prints 33 "
    "items returned by _count_equipment_in_pdf for 005-Plany osveshcheniya "
    "PDF. Evidence: (a) B028 wiring reducer now emits "
    "'Provodka prokladyvaemaya na lotke' row (legend [13] reached output), "
    "(b) B029 fallback emits unit='sht' cable-laying rows instead of "
    "skipping when length_m is None. Side observations (not T047 scope): "
    "work_name prefix garbling on wiring row ('Prokladka provoda ka "
    "prokladyvaemaya...') and duplicate 'Prokladka kabelya' rows "
    "between legend-path unit='m' and cable_derivation unit='sht' -- these "
    "belong to B028/B029/dedupe follow-ups, not diag-script fix. "
    "No code commit needed (diag script is internal tooling, not release "
    "artefact)."
)

fixed_s005 = False
fixed_s010 = False
now = datetime.now().isoformat(timespec="seconds")

for t in data.get("tasks", []):
    if t.get("id") != "T047":
        continue
    sid = t.get("sprint_id")
    ex = t.get("executor")
    if sid == "S005" and ex == "tester":
        t["result"] = S005_RESULT
        t["updated_at"] = now
        fixed_s005 = True
        print(f"[OK] Restored S005/T047 (tester). result_len={len(S005_RESULT)}")
    elif sid == "S010" and ex == "boss":
        t["result"] = S010_RESULT
        t["status"] = "done"
        t["updated_at"] = now
        fixed_s010 = True
        print(f"[OK] Completed S010/T047 (boss). result_len={len(S010_RESULT)} status=done")

if not (fixed_s005 and fixed_s010):
    print(f"[ERR] Expected to fix both; got s005={fixed_s005} s010={fixed_s010}")
    sys.exit(1)

TASKS.write_text(
    json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
)
print("[OK] tasks.json written")
