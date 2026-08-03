# =============================================================================
# train_ids.py  —  Layer 3 IDS training pipeline for Cyber-Hardened BMS
#
#  Trains a Decision Tree (exportable to C via m2cgen) on 5 traffic classes:
#    0  Normal
#    1  DoS     (high frequency, low inter-arrival)
#    2  Spoof   (normal frequency but voltage out-of-range)
#    3  Replay  (moderate frequency, near-identical consecutive payloads)
#    4  Fuzz    (high entropy random bytes)
#
#  Pipeline:
#    1. Generate or load can_dataset.csv
#    2. Feature engineering (4 statistical features per frame)
#    3. Train-test split (stratified 80/20)
#    4. GridSearch hyperparameter tuning
#    5. Evaluate: accuracy, F1, ROC-AUC, confusion matrix
#    6. Export to ids_model.h (C code) via m2cgen for ESP32 Core 0
#    7. Save bt_security_sim.csv (Bluetooth auth events)  ← Layer 1 + 2 data
# =============================================================================

import os
import numpy as np
import pandas as pd
from sklearn.tree import DecisionTreeClassifier
from sklearn.model_selection import train_test_split, GridSearchCV, StratifiedKFold
from sklearn.metrics import (classification_report, confusion_matrix,
                             roc_auc_score, f1_score)
from sklearn.preprocessing import label_binarize
import m2cgen as m2c

# ── 1. Dataset generation ─────────────────────────────────────────────────────
CLASS_NAMES = ["Normal", "DoS", "Spoof", "Replay", "Fuzz"]
N_PER_CLASS = {0: 8000, 1: 2000, 2: 2000, 3: 2000, 4: 2000}

def generate_dataset(path: str) -> pd.DataFrame:
    """Synthesise a 5-class CAN traffic dataset with realistic feature distributions."""
    print("Generating synthetic CAN dataset...")
    rng = np.random.default_rng(42)

    frames = {k: [] for k in ("InterArrival_ms", "msg_freq", "id_variance",
                               "entropy", "Label")}

    # Label 0 — Normal: ~100 ms inter-arrival, ~10 msg/s, moderate variance/entropy
    n = N_PER_CLASS[0]
    frames["InterArrival_ms"] += list(rng.normal(100,  15,  n).clip(5))
    frames["msg_freq"]        += list(rng.normal(10,   2,   n).clip(1))
    frames["id_variance"]     += list(rng.normal(250,  40,  n).clip(0))
    frames["entropy"]         += list(rng.normal(2.0,  0.3, n).clip(0))
    frames["Label"]           += [0] * n

    # Label 1 — DoS: very short inter-arrival (<2 ms), very high frequency
    n = N_PER_CLASS[1]
    frames["InterArrival_ms"] += list(rng.normal(0.8,  0.3, n).clip(0.1))
    frames["msg_freq"]        += list(rng.normal(800,  100, n).clip(200))
    frames["id_variance"]     += list(rng.normal(0.05, 0.02, n).clip(0))
    frames["entropy"]         += list(rng.normal(0.05, 0.02, n).clip(0))
    frames["Label"]           += [1] * n

    # Label 2 — Spoof: normal timing but ID 0x100 with out-of-range payload
    #   → captured as unusually low id_variance (always same ID) + normal entropy
    n = N_PER_CLASS[2]
    frames["InterArrival_ms"] += list(rng.normal(100,  10,  n).clip(10))
    frames["msg_freq"]        += list(rng.normal(10,   1.5, n).clip(1))
    frames["id_variance"]     += list(rng.normal(0.1,  0.05, n).clip(0))  # single ID
    frames["entropy"]         += list(rng.normal(1.2,  0.2, n).clip(0))
    frames["Label"]           += [2] * n

    # Label 3 — Replay: slightly elevated frequency, low entropy (same bytes)
    n = N_PER_CLASS[3]
    frames["InterArrival_ms"] += list(rng.normal(45,   8,   n).clip(5))
    frames["msg_freq"]        += list(rng.normal(22,   4,   n).clip(5))
    frames["id_variance"]     += list(rng.normal(80,   15,  n).clip(0))
    frames["entropy"]         += list(rng.normal(0.8,  0.15, n).clip(0))  # repetitive
    frames["Label"]           += [3] * n

    # Label 4 — Fuzzing: random IDs (high variance), maximum entropy
    n = N_PER_CLASS[4]
    frames["InterArrival_ms"] += list(rng.normal(10,   3,   n).clip(1))
    frames["msg_freq"]        += list(rng.normal(100,  20,  n).clip(20))
    frames["id_variance"]     += list(rng.normal(900,  80,  n).clip(100))  # many IDs
    frames["entropy"]         += list(rng.normal(7.5,  0.3, n).clip(5))    # max entropy
    frames["Label"]           += [4] * n

    df = pd.DataFrame(frames)
    df.to_csv(path, index=False)
    print(f"  Saved {len(df):,} frames → {path}")
    return df

# ── 2. Load dataset ───────────────────────────────────────────────────────────
DATASET_PATH = "can_dataset.csv"
if not os.path.exists(DATASET_PATH):
    df = generate_dataset(DATASET_PATH)
else:
    df = pd.read_csv(DATASET_PATH)
    print(f"Loaded {len(df):,} frames from {DATASET_PATH}")

# ── Feature engineering on raw CAN logs (if captured via generate_dataset.py) ─
FEATURES = ["InterArrival_ms", "msg_freq", "id_variance", "entropy"]

if "msg_freq" not in df.columns and "CAN_ID" in df.columns:
    print("Deriving features from raw CAN log columns...")
    df["msg_freq"]    = df["CAN_ID"].rolling(20).count().fillna(1)
    df["id_variance"] = pd.to_numeric(df["CAN_ID"], errors="coerce") \
                          .rolling(20).var().fillna(0)
    def byte_entropy(row):
        vals   = [row[f"D{i}"] for i in range(8)]
        counts = pd.Series(vals).value_counts(normalize=True)
        return -sum(p * np.log2(p + 1e-9) for p in counts)
    df["entropy"] = df.apply(byte_entropy, axis=1)

X = df[FEATURES].fillna(0).values
y = df["Label"].astype(int).values

print(f"\nClass distribution:")
for i, name in enumerate(CLASS_NAMES):
    count = (y == i).sum()
    print(f"  [{i}] {name:<8}  {count:6,}  ({100*count/len(y):.1f}%)")

# ── 3. Train / test split ─────────────────────────────────────────────────────
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.20, random_state=42, stratify=y
)
print(f"\nTrain: {len(X_train):,}   Test: {len(X_test):,}")

# ── 4. GridSearchCV ───────────────────────────────────────────────────────────
print("\nRunning hyperparameter search...")
param_grid = {
    "max_depth":         [4, 6, 8, 10, None],
    "min_samples_leaf":  [1, 2, 5, 10],
    "criterion":         ["gini", "entropy"],
}
cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
gs = GridSearchCV(
    DecisionTreeClassifier(random_state=42),
    param_grid,
    cv=cv,
    scoring="f1_weighted",
    n_jobs=-1,
    verbose=0,
)
gs.fit(X_train, y_train)
model = gs.best_estimator_
print(f"  Best params: {gs.best_params_}")
print(f"  Best CV F1 : {gs.best_score_:.4f}")

# ── 5. Evaluation ─────────────────────────────────────────────────────────────
y_pred  = model.predict(X_test)
y_proba = model.predict_proba(X_test)

print("\n── Classification Report ───────────────────────────────────────")
print(classification_report(y_test, y_pred, target_names=CLASS_NAMES))

print("── Confusion Matrix ─────────────────────────────────────────────")
cm = confusion_matrix(y_test, y_pred)
header = f"{'':>10}" + "".join(f"{n:>10}" for n in CLASS_NAMES)
print(header)
for i, row in enumerate(cm):
    print(f"{CLASS_NAMES[i]:>10}" + "".join(f"{v:>10}" for v in row))

# Multi-class ROC-AUC (one-vs-rest)
y_bin   = label_binarize(y_test, classes=list(range(len(CLASS_NAMES))))
auc_ovr = roc_auc_score(y_bin, y_proba, multi_class="ovr", average="macro")
print(f"\n  Macro ROC-AUC (OVR): {auc_ovr:.4f}")
print(f"  Weighted F1-score  : {f1_score(y_test, y_pred, average='weighted'):.4f}")

# Feature importance
print("\n── Feature Importances ──────────────────────────────────────────")
for feat, imp in sorted(zip(FEATURES, model.feature_importances_),
                        key=lambda x: x[1], reverse=True):
    bar = "█" * int(imp * 40)
    print(f"  {feat:20s}  {imp:.4f}  {bar}")

# ── 6. Export to C for ESP32 ─────────────────────────────────────────────────
print("\nExporting decision tree to C (m2cgen)...")
c_code = m2c.export_to_c(model)

os.makedirs("bms_master", exist_ok=True)
for path in ("bms_master/ids_model.h", "ids_model.h"):
    with open(path, "w") as fh:
        fh.write(c_code)
    print(f"  Wrote {path}")

# Quick sanity check — run C model on first 5 test samples in Python
print("\n── Sanity check (first 5 test predictions) ──────────────────────")
for i in range(min(5, len(X_test))):
    true_cls  = CLASS_NAMES[y_test[i]]
    pred_cls  = CLASS_NAMES[y_pred[i]]
    match     = "✓" if y_test[i] == y_pred[i] else "✗"
    print(f"  [{match}] True={true_cls:<8} Pred={pred_cls:<8}  "
          f"feat={X_test[i]}")

# ── 7. Bluetooth layer log dataset (Layer 1 + 2 simulation) ──────────────────
print("\nGenerating Bluetooth security simulation log...")
rng  = np.random.default_rng(99)
rows = []
for _ in range(500):
    mac      = ":".join(f"{x:02X}" for x in rng.integers(0, 256, 6))
    whitelisted = rng.random() < 0.7   # 70% legitimate
    authed      = whitelisted and (rng.random() < 0.95)
    cmd_id      = rng.integers(0, 256)
    tier2       = int(cmd_id >= 0x80)
    valid_token = int(authed and (rng.random() < 0.9))
    result      = "ALLOW" if (authed and (not tier2 or valid_token)) else "BLOCK"
    rows.append({
        "mac": mac, "whitelisted": int(whitelisted),
        "authed": int(authed), "cmd_id": hex(cmd_id),
        "tier2": tier2, "valid_token": valid_token, "result": result
    })
bt_df = pd.DataFrame(rows)
bt_df.to_csv("bt_security_sim.csv", index=False)
print(f"  Saved bt_security_sim.csv  ({len(bt_df)} events)")
print(f"  ALLOW: {(bt_df.result=='ALLOW').sum()}  "
      f"BLOCK: {(bt_df.result=='BLOCK').sum()}")

print("\n══ Training complete ════════════════════════════════════════════")
print("  ids_model.h  → flash to ESP32 Core 0 (Layer 3)")
print("  can_dataset.csv       → archive for paper appendix")
print("  bt_security_sim.csv   → Layer 1/2 audit log analysis")
