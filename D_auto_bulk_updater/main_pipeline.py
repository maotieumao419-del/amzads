"""
main_pipeline.py  —  Pipeline Orchestrator
==========================================
Executes the full PPC data pipeline sequentially:

  Step 1  rebuild   io_handlers/rebuild_internal_file.py
              raw_xlsx/*.xlsx  →  working_json/*.json

  Step 2  sync      core_logic/auto_bulk_sync.py
              working_json/BulkSheetExport_*.json
              + working_json/PPC_*.json
              →  working_json/PPC_*_synced.json
              →  final_xlsx/PPC_*.xlsx

  Step 3  tracker   core_logic/end_to_end_ppc_tracker.py
              working_json/BulkSheetExport_*.json
              + working_json/PPC_*_synced.json 
              →  working_json/PPC_*_UPDATED.json
              →  final_xlsx/PPC_*_UPDATED.xlsx

  Step 4  update    core_logic/auto_bulk_updater.py
              working_json/PPC_*_UPDATED.json
              + rules.json
              →  working_json/Amazon_Upload_Ready_*.json
              →  final_xlsx/Amazon_Upload_Ready_*.xlsx

  Step 5  validate  core_logic/validate_metrics.py
              2-level integrity check (metrics + data-drift)

Usage:
  python main_pipeline.py              # run all steps
  python main_pipeline.py --from sync  # start from step 2
  python main_pipeline.py --only validate
"""

import os
import sys
import time
import logging
import argparse
import importlib
import importlib.util

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s'
)

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
IO_DIR      = os.path.join(BASE_DIR, 'io_handlers')
LOGIC_DIR   = os.path.join(BASE_DIR, 'core_logic')
RAW_DIR     = os.path.join(BASE_DIR, 'data', 'raw_xlsx')
JSON_DIR    = os.path.join(BASE_DIR, 'data', 'working_json')
FINAL_DIR   = os.path.join(BASE_DIR, 'data', 'final_xlsx')

for d in [RAW_DIR, JSON_DIR, FINAL_DIR]:
    os.makedirs(d, exist_ok=True)

sys.path.insert(0, IO_DIR)
sys.path.insert(0, LOGIC_DIR)


# ─── Helpers ────────────────────────────────────────────────────────────────

class _Timer:
    def __init__(self, label: str):
        self.label = label
    def __enter__(self):
        self._start = time.time()
        return self
    def __exit__(self, *_):
        elapsed = time.time() - self._start
        logging.info(f"[{self.label}] completed in {elapsed:.2f}s")


def _header(title: str):
    bar = '=' * 62
    print(f"\n{bar}\n  {title}\n{bar}")


def _run_step(step_name: str, module_path: str) -> bool:
    """Import and run the main() of a pipeline module."""
    _header(f"Step: {step_name}")
    try:
        # Force fresh import each call so state doesn't bleed between steps
        spec   = importlib.util.spec_from_file_location(step_name, module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        if hasattr(module, 'main'):
            with _Timer(step_name):
                module.main()
        return True
    except SystemExit:
        return True   # scripts that call sys.exit(0) on success
    except Exception as e:
        logging.error(f"[{step_name}] FAILED: {e}")
        import traceback; traceback.print_exc()
        return False


# ─── Step definitions ────────────────────────────────────────────────────────

STEPS = {
    'rebuild':  os.path.join(IO_DIR,    'rebuild_internal_file.py'),
    'sync':     os.path.join(LOGIC_DIR, 'auto_bulk_sync.py'),
    'tracker':  os.path.join(LOGIC_DIR, 'end_to_end_ppc_tracker.py'),
    'update':   os.path.join(LOGIC_DIR, 'auto_bulk_updater.py'),
    'validate': os.path.join(LOGIC_DIR, 'validate_metrics.py'),
}

STEP_ORDER = ['rebuild', 'sync', 'tracker', 'update', 'validate']


# ─── Pre-flight checks ───────────────────────────────────────────────────────

def preflight() -> bool:
    """Verify that raw_xlsx/ contains at least one bulk file before starting."""
    import glob
    bulk = glob.glob(os.path.join(RAW_DIR, 'BulkSheetExport_*.xlsx'))
    ppc  = glob.glob(os.path.join(RAW_DIR, 'PPC_*.xlsx'))

    ok = True
    if not bulk:
        logging.error(
            f"Pre-flight FAIL: No BulkSheetExport_*.xlsx in {RAW_DIR}\n"
            f"  → Copy your Amazon Bulk file there and re-run."
        )
        ok = False
    if not ppc:
        logging.warning(
            f"Pre-flight WARN: No PPC_*.xlsx in {RAW_DIR}\n"
            f"  → Sync step will have nothing to update."
        )
    return ok


# ─── Main orchestrator ───────────────────────────────────────────────────────

def run_pipeline(start_from: str = 'rebuild', only: str = None) -> bool:
    """
    Execute steps in order, optionally starting from a specific step or
    running only one step.
    Returns True if all executed steps passed.
    """
    if only:
        steps_to_run = [only] if only in STEPS else []
    else:
        idx = STEP_ORDER.index(start_from) if start_from in STEP_ORDER else 0
        steps_to_run = STEP_ORDER[idx:]

    if not steps_to_run:
        logging.error(
            f"Unknown step '{only or start_from}'. "
            f"Valid steps: {', '.join(STEP_ORDER)}"
        )
        return False

    t_total   = time.time()
    results   = {}
    all_pass  = True

    # Pre-flight only when rebuilding
    if 'rebuild' in steps_to_run:
        if not preflight():
            return False

    for step in steps_to_run:
        ok = _run_step(step, STEPS[step])
        results[step] = ok
        if not ok:
            all_pass = False
            logging.error(f"Pipeline halted at step '{step}'.")
            break

    # ── Final summary ─────────────────────────────────────────────────
    _header("PIPELINE SUMMARY")
    for step in steps_to_run:
        status = 'PASS' if results.get(step) else 'FAIL'
        print(f"  {step:<12}  {status}")

    elapsed = time.time() - t_total
    print(f"\n  Total runtime  : {elapsed:.1f}s")
    print(f"  Final status   : {'SUCCESS' if all_pass else 'FAILED'}")
    print('-' * 62 + '\n')

    return all_pass


# ─── CLI ────────────────────────────────────────────────────────────────────

def _parse_args():
    p = argparse.ArgumentParser(
        description='Amazon PPC Pipeline Orchestrator',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main_pipeline.py                    # full pipeline
  python main_pipeline.py --from sync        # skip rebuild
  python main_pipeline.py --from tracker     # re-inject metrics only
  python main_pipeline.py --only validate    # validation only
        """
    )
    p.add_argument(
        '--from', dest='start_from', default='rebuild',
        choices=STEP_ORDER,
        help='Start pipeline from this step (default: rebuild)'
    )
    p.add_argument(
        '--only', default=None,
        choices=STEP_ORDER,
        help='Run only this one step'
    )
    return p.parse_args()


if __name__ == '__main__':
    args    = _parse_args()
    success = run_pipeline(start_from=args.start_from, only=args.only)
    sys.exit(0 if success else 1)
