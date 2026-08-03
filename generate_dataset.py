# =============================================================================
# generate_dataset.py  —  Live CAN + Bluetooth telemetry capture
#
#  Captures serial output from the attacker_node ESP32 into:
#    can_dataset.csv     — CAN frame log with features + labels
#    bt_events.csv       — BLE authentication + command events
#
#  Usage:
#    python generate_dataset.py COM3           (Windows)
#    python generate_dataset.py /dev/ttyUSB0  (Linux)
#
#  Expected attacker_node serial format (one line per event):
#    CAN frame:
#      Timestamp,CAN_ID,DLC,D0,D1,D2,D3,D4,D5,D6,D7,InterArrival_ms,Label
#    BT scan event:
#      [ATTACK] BT_SCAN attempt from MAC: AA:BB:CC:DD:EE:FF
#    IDS telemetry from bms_master:
#      IDS,<ts>,<class>,<dt>,<freq>,<var>,<ent>,<score>
#    EKF telemetry from bms_master:
#      DATA,<soc>,<v_meas>,<I>,<anomaly>,<R_eff>
#
#  Features derived per frame:
#    InterArrival_ms   — time since last frame on same CAN ID (ms)
#    msg_freq          — 20-frame rolling message rate
#    id_variance       — 20-frame rolling variance of CAN ID values
#    entropy           — Shannon entropy of 8 data bytes
#
#  Pressing Ctrl+C stops capture and saves both files.
# =============================================================================

import serial
import csv
import sys
import time
import math
from collections import deque

# ── Config ────────────────────────────────────────────────────────────────────
PORT = sys.argv[1] if len(sys.argv) > 1 else "COM3"
BAUD = 115200
CAN_OUT  = "can_dataset.csv"
BT_OUT   = "bt_events.csv"

# ── Rolling windows for feature computation on raw bytes ─────────────────────
WINDOW    = 20
id_window = deque(maxlen=WINDOW)
t_window  = deque(maxlen=WINDOW)

def rolling_freq():
    if len(t_window) < 2: return 0.0
    span = max(t_window[-1] - t_window[0], 1e-3) / 1000.0  # ms → s
    return len(t_window) / span

def rolling_id_variance():
    if len(id_window) < 2: return 0.0
    arr = list(id_window)
    mean = sum(arr) / len(arr)
    return sum((x - mean) ** 2 for x in arr) / len(arr)

def byte_entropy(data_bytes):
    counts = {}
    for b in data_bytes:
        counts[b] = counts.get(b, 0) + 1
    n = len(data_bytes)
    H = 0.0
    for c in counts.values():
        p = c / n
        H -= p * math.log2(p + 1e-9)
    return H

# ── Column definitions ────────────────────────────────────────────────────────
CAN_COLS = [
    "Timestamp_ms", "CAN_ID", "DLC",
    "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7",
    "InterArrival_ms", "msg_freq", "id_variance", "entropy",
    "Label"
]

BT_COLS = [
    "Timestamp_ms", "Event", "MAC", "Detail"
]

# ── Main capture loop ─────────────────────────────────────────────────────────
print(f"Connecting to {PORT} at {BAUD} baud...")
print(f"  CAN frames  → {CAN_OUT}")
print(f"  BT events   → {BT_OUT}")
print("Press Ctrl+C to stop and save.\n")

can_rows = 0
bt_rows  = 0
t_start  = time.time()

try:
    ser = serial.Serial(PORT, BAUD, timeout=1)
    with open(CAN_OUT, "w", newline="") as can_f, \
         open(BT_OUT,  "w", newline="") as bt_f:

        can_writer = csv.writer(can_f)
        bt_writer  = csv.writer(bt_f)
        can_writer.writerow(CAN_COLS)
        bt_writer.writerow(BT_COLS)

        while True:
            raw = ser.readline().decode("utf-8", errors="ignore").strip()
            if not raw:
                continue

            ts = int((time.time() - t_start) * 1000)

            # ── BLE / Layer 1 event ───────────────────────────────────────────
            if "[ATTACK] BT_SCAN" in raw or "[L1]" in raw or "[L2]" in raw:
                # Extract MAC if present
                mac   = ""
                parts = raw.split("MAC:")
                if len(parts) > 1:
                    mac = parts[1].strip().split()[0]
                bt_writer.writerow([ts, raw[:30], mac, raw])
                bt_rows += 1
                print(f"  [BT]  {raw}")
                continue

            # ── UDS event ─────────────────────────────────────────────────────
            if "[UDS]" in raw or "SecurityAccess" in raw:
                print(f"  [UDS] {raw}")
                continue

            # ── IDS telemetry line (from bms_master securityTask) ─────────────
            if raw.startswith("IDS,"):
                parts = raw.split(",")
                if len(parts) >= 8:
                    print(f"  [IDS] class={parts[2]}  score={parts[7]}")
                continue

            # ── EKF DATA line (from bms_master controlTask) ───────────────────
            if raw.startswith("DATA,"):
                parts = raw.split(",")
                if len(parts) >= 5:
                    print(f"  [EKF] SoC={parts[1]}%  "
                          f"anomaly={parts[4]}  R_eff={parts[5] if len(parts)>5 else '?'}")
                continue

            # ── CAN frame line ────────────────────────────────────────────────
            if "," not in raw:
                continue

            cols = raw.split(",")
            if len(cols) < 13:
                continue

            try:
                frame_ts  = int(cols[0])
                can_id    = int(cols[1], 16) if cols[1].startswith("0x") \
                            else int(cols[1])
                dlc       = int(cols[2])
                data      = [int(cols[3 + i], 16) for i in range(8)]
                iat_ms    = float(cols[11])
                label     = int(cols[12]) if len(cols) > 12 else 0
            except (ValueError, IndexError):
                continue

            # Update rolling windows
            id_window.append(can_id)
            t_window.append(frame_ts)

            freq   = rolling_freq()
            id_var = rolling_id_variance()
            ent    = byte_entropy(data)

            row = [
                frame_ts,
                hex(can_id), dlc,
                *[hex(b) for b in data],
                iat_ms, round(freq, 2), round(id_var, 2), round(ent, 4),
                label
            ]
            can_writer.writerow(row)
            can_rows += 1

            # Progress ticker
            if can_rows % 200 == 0:
                elapsed = time.time() - t_start
                print(f"  CAN: {can_rows:6,} frames  |  "
                      f"BT: {bt_rows} events  |  "
                      f"elapsed: {elapsed:.0f}s")

except KeyboardInterrupt:
    elapsed = time.time() - t_start
    print(f"\n── Capture stopped ──────────────────────────────────────────")
    print(f"  CAN frames saved : {can_rows:,}  → {CAN_OUT}")
    print(f"  BT events saved  : {bt_rows}     → {BT_OUT}")
    print(f"  Duration         : {elapsed:.1f} s")
    print(f"  Next step: run train_ids.py to train the Layer 3 IDS model")

except serial.SerialException as e:
    print(f"Serial port error: {e}")
    print(f"Check that {PORT} is correct and attacker_node is connected.")
    sys.exit(1)
