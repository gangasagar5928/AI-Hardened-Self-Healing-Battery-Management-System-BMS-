# =============================================================================
# capture_run.py  —  Live test-run metric capture for post-processing
#
#  Captures serial telemetry from bms_master and logs:
#    run_log.csv      — EKF / BMS state per 100 ms tick
#    security_log.csv — Layer 5 security events (attacks, auth failures)
#    bt_session.csv   — Layer 1 BLE session events
#
#  Usage:
#    python capture_run.py COM3         (BMS master port)
#    python capture_run.py COM3 COM5    (BMS on COM3, attacker on COM5)
#
#  Expected serial line types from bms_master:
#    EKF data:   DATA,<soc>,<v_meas>,<I>,<anomaly>,<R_eff>
#    IDS result: IDS,<ts>,<class>,<dt>,<freq>,<var>,<ent>,<score>
#    L5 log:     [L5] LOG  type=...  score=...  can=...  mac=...
#    L1 event:   [L1] Auth OK: AA:BB:CC:DD:EE:FF
#    L2 event:   [L2] CMD: ...  or  [L2] Tier-2 cmd BLOCKED
# =============================================================================

import serial
import serial.threaded
import csv
import sys
import time
import threading
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
BMS_PORT  = sys.argv[1] if len(sys.argv) > 1 else "COM3"
ATK_PORT  = sys.argv[2] if len(sys.argv) > 2 else None
BAUD      = 115200

RUN_OUT  = "run_log.csv"
SEC_OUT  = "security_log.csv"
BT_OUT   = "bt_session.csv"

RUN_COLS = [
    "t_s", "t_ms",
    "SoC_pct", "V_meas", "I_A",
    "anomaly_score", "R_eff",
    "layer3_class"
]
SEC_COLS = ["t_ms", "attack_type", "anomaly_score", "can_id", "mac"]
BT_COLS  = ["t_ms", "event_type",  "mac", "result"]

# ── Shared state ──────────────────────────────────────────────────────────────
lock           = threading.Lock()
run_rows       = []
sec_rows       = []
bt_rows        = []
last_ids_class = "Normal"
t0             = time.time()

def elapsed_s():  return time.time() - t0
def elapsed_ms(): return int((time.time() - t0) * 1000)

# ── Line parser ───────────────────────────────────────────────────────────────
def parse_line(raw: str):
    global last_ids_class

    # ── EKF data line: DATA,<soc>,<v>,<I>,<anomaly>,<R_eff> ──────────────────
    if raw.startswith("DATA,"):
        parts = raw.split(",")
        if len(parts) >= 5:
            try:
                row = {
                    "t_s":           round(elapsed_s(),  3),
                    "t_ms":          elapsed_ms(),
                    "SoC_pct":       float(parts[1]),
                    "V_meas":        float(parts[2]),
                    "I_A":           float(parts[3]),
                    "anomaly_score": float(parts[4]),
                    "R_eff":         float(parts[5]) if len(parts) > 5 else None,
                    "layer3_class":  last_ids_class,
                }
                with lock: run_rows.append(row)
                print(f"[EKF] SoC={parts[1]:>6}%  "
                      f"anomaly={float(parts[4]):.2f}  "
                      f"R_eff={parts[5] if len(parts)>5 else '?':<12}")
            except ValueError:
                pass
        return

    # ── IDS line: IDS,<ts>,<class>,<dt>,<freq>,<var>,<ent>,<score> ───────────
    if raw.startswith("IDS,"):
        parts = raw.split(",")
        if len(parts) >= 8:
            try:
                cls_map = {0:"Normal", 1:"DoS", 2:"Spoof", 3:"Replay", 4:"Fuzz"}
                cls_id  = int(parts[2])
                last_ids_class = cls_map.get(cls_id, "Unknown")
                score   = float(parts[7])
                if cls_id != 0:
                    sec_row = {
                        "t_ms":          int(parts[1]),
                        "attack_type":   last_ids_class,
                        "anomaly_score": score,
                        "can_id":        "",
                        "mac":           ""
                    }
                    with lock: sec_rows.append(sec_row)
                    print(f"[L3]  Attack={last_ids_class:<8}  score={score:.2f}")
            except (ValueError, IndexError):
                pass
        return

    # ── Layer 5 security log: [L5] LOG  type=...  score=...  can=...  mac=... ─
    if "[L5] LOG" in raw:
        row = {"t_ms": elapsed_ms(), "attack_type": "", "anomaly_score": 0.0,
               "can_id": "", "mac": ""}
        for token in raw.split():
            if token.startswith("type="):   row["attack_type"]   = token[5:]
            if token.startswith("score="):  row["anomaly_score"] = float(token[6:])
            if token.startswith("can="):    row["can_id"]        = token[4:]
            if token.startswith("mac="):    row["mac"]           = token[4:]
        with lock: sec_rows.append(row)
        print(f"[L5]  {raw}")
        return

    # ── Layer 1 BLE event ─────────────────────────────────────────────────────
    if "[L1]" in raw or "[L2]" in raw:
        mac    = ""
        parts2 = raw.split("MAC:")
        if len(parts2) > 1: mac = parts2[1].strip().split()[0]
        evt    = "AUTH_OK"   if "Auth OK"    in raw else \
                 "AUTH_FAIL" if "Auth FAIL"  in raw else \
                 "TIMEOUT"   if "TIMEOUT"    in raw else \
                 "L2_BLOCK"  if "BLOCKED"    in raw else \
                 "L2_EXEC"   if "CMD:"       in raw else raw[:25]
        result = "BLOCKED" if any(k in raw for k in ["FAIL","BLOCK","EXPIRE"]) \
                 else "ALLOW"
        row = {"t_ms": elapsed_ms(), "event_type": evt, "mac": mac, "result": result}
        with lock: bt_rows.append(row)
        print(f"[BT]  {evt:<20} mac={mac:<18} → {result}")
        return

# ── Serial reader thread ──────────────────────────────────────────────────────
def reader_thread(port: str, label: str):
    try:
        ser = serial.Serial(port, BAUD, timeout=1)
        print(f"  Connected to {port} ({label})")
        while running:
            try:
                raw = ser.readline().decode("utf-8", errors="ignore").strip()
            except serial.SerialException:
                break
            if raw:
                parse_line(raw)
    except serial.SerialException as e:
        print(f"  [{label}] Serial error: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
running = True
threads = []

print("═" * 60)
print("  Cyber-Hardened BMS — 5-Layer Live Run Capture")
print("═" * 60)
print(f"  BMS port   : {BMS_PORT}")
print(f"  Attack port: {ATK_PORT or 'N/A (single-node mode)'}")
print(f"  Outputs    : {RUN_OUT}  {SEC_OUT}  {BT_OUT}")
print("  Press Ctrl+C to stop and save.\n")

t0 = time.time()

bms_thread = threading.Thread(target=reader_thread,
                              args=(BMS_PORT, "BMS"), daemon=True)
bms_thread.start()
threads.append(bms_thread)

if ATK_PORT:
    atk_thread = threading.Thread(target=reader_thread,
                                  args=(ATK_PORT, "ATK"), daemon=True)
    atk_thread.start()
    threads.append(atk_thread)

try:
    while True:
        time.sleep(10)
        with lock:
            print(f"\n[TICK] t={elapsed_s():.0f}s  "
                  f"run_rows={len(run_rows)}  "
                  f"sec_events={len(sec_rows)}  "
                  f"bt_events={len(bt_rows)}")
except KeyboardInterrupt:
    running = False

# ── Save CSV files ────────────────────────────────────────────────────────────
print("\n── Saving logs ──────────────────────────────────────────────────")
with lock:
    _run  = list(run_rows)
    _sec  = list(sec_rows)
    _bt   = list(bt_rows)

with open(RUN_OUT, "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=RUN_COLS)
    w.writeheader(); w.writerows(_run)
print(f"  {RUN_OUT:<25}  {len(_run):,} rows")

with open(SEC_OUT, "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=SEC_COLS)
    w.writeheader(); w.writerows(_sec)
print(f"  {SEC_OUT:<25}  {len(_sec):,} rows")

with open(BT_OUT, "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=BT_COLS)
    w.writeheader(); w.writerows(_bt)
print(f"  {BT_OUT:<25}  {len(_bt):,} rows")

# ── Summary statistics ────────────────────────────────────────────────────────
if _run:
    import statistics
    scores = [r["anomaly_score"] for r in _run if r["anomaly_score"] is not None]
    socs   = [r["SoC_pct"]       for r in _run if r["SoC_pct"]       is not None]
    print(f"\n── Run Summary ──────────────────────────────────────────────")
    print(f"  Duration          : {elapsed_s():.1f} s")
    print(f"  EKF samples       : {len(_run)}")
    print(f"  Security events   : {len(_sec)}")
    print(f"  BT session events : {len(_bt)}")
    if scores:
        print(f"  Anomaly score     : "
              f"min={min(scores):.3f}  "
              f"max={max(scores):.3f}  "
              f"mean={statistics.mean(scores):.3f}")
    if socs:
        print(f"  SoC range         : "
              f"{min(socs):.1f}%  →  {max(socs):.1f}%")

    attack_types = {}
    for r in _sec:
        t = r["attack_type"]
        attack_types[t] = attack_types.get(t, 0) + 1
    if attack_types:
        print(f"\n  Detected attacks:")
        for atype, cnt in sorted(attack_types.items(), key=lambda x: -x[1]):
            print(f"    {atype:<20} {cnt:4} events")

print(f"\n  Next step: open run_log.csv in Excel / pandas for SoC vs attack plots")
