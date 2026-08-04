# Cyber-Hardened BMS — API & Protocol Documentation

This document specifies the CAN message database, UDS diagnostic protocols, BLE GATT services, C++ firmware functions, and Python Machine Learning APIs.

---

## 1. CAN Bus Message Specification (500 kbps)

### Message Matrix

| CAN ID (Hex) | Transmitting Node | Receiving Node | Cycle Time | Length (DLC) | Description |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `0x180` | `bms_master` | TCU / Vehicle ECU | $100\text{ ms}$ | 8 Bytes | BMS Telemetry & Anomaly Status |
| `0x120` | Motor Controller / Charger | `bms_master` | $100\text{ ms}$ | 8 Bytes | Vehicle Load Current & Charge Commands |
| `0x188` | `bms_master` | Cloud Gateway / Fleet | Asynchronous | 8 Bytes | Federated Learning Model Weight Delta |
| `0x190` | `bms_master` | Inspector / App | $1000\text{ ms}$ | 8 Bytes | Digital Battery Health Passport Digest |
| `0x198` | Smart Grid Charger | `bms_master` | $1000\text{ ms}$ | 8 Bytes | Grid Demand-Response Peak Stress Signal |
| `0x000` | Attacker Node | CAN Bus | Variable | 8 Bytes | DoS High-Frequency Flood Vector |
| `0x7E0` | OBD-II Diagnostic Tester | `bms_master` | Asynchronous | 8 Bytes | UDS Request Frame |
| `0x7E8` | `bms_master` | Diagnostic Tester | Asynchronous | 8 Bytes | UDS Response Frame |

---

### CAN ID `0x180` — BMS Master Telemetry Frame Layout

| Byte Offset | Data Field | Data Type | Scale / Format | Range |
| :--- | :--- | :--- | :--- | :--- |
| `Byte 0-1` | State of Charge (SoC) | `uint16_t` | `val / 100.0` (%) | $0.0\% - 100.0\%$ |
| `Byte 2-3` | Pack Voltage | `uint16_t` | `val / 1000.0` (V) | $0.000\text{V} - 20.000\text{V}$ |
| `Byte 4-5` | Pack Current | `int16_t` | `val / 1000.0` (A) | $-50.000\text{A} - +50.000\text{A}$ |
| `Byte 6` | Anomaly Score $S$ | `uint8_t` | `val / 100.0` | $0.00 - 1.00$ |
| `Byte 7` | Self-Healing Action | `uint8_t` | Enum ID ($0-4$) | `0:NORMAL, 1:RED_CHG, 2:DIS_BAL, 3:LIMP, 4:ISOLATE` |

---

## 2. UDS (ISO 14229 / OBD-II) Protocol API

The BMS monitors diagnostic messages over CAN ID `0x7E0`. Layer 3 intercepts unauthorized session requests:

| Service ID (SID) | Service Name | Expected Subfunction | Security Policy | Action on Unauthorized Access |
| :--- | :--- | :--- | :--- | :--- |
| `0x10` | DiagnosticSessionControl | `0x01` (Default), `0x03` (Extended) | Read-only session allowed | Log event |
| `0x27` | **SecurityAccess** | `0x01` (Request Seed), `0x02` (Send Key) | **Authentication Required** | **Override $S = 0.96$, block seed release** |
| `0x3E` | **TesterPresent** | `0x00` (Zero Subfunction) | **Session Active Check** | **Flag session hijack if unauthenticated** |
| `0x22` | ReadDataByIdentifier | `0x0100` (SoC), `0x0101` (Cell V) | Public Read | Return data frame |
| `0x2E` | WriteDataByIdentifier | `0x0200` (Clear NVS / Calibration) | **Admin Authenticated Only** | **Block write & trip alarm** |

---

## 3. BLE GATT Telemetry & Auth Specification

* **Device Name:** `BMS_SecureBLE`
* **Service UUID:** `0000180F-0000-1000-8000-00805F9B34FB`

### Characteristics

| Characteristic UUID | Properties | Format | Description |
| :--- | :--- | :--- | :--- |
| `00002A19-0000-1000-8000-00805F9B34FB` | Read, Notify | 8-byte Binary | Encrypted BMS Status (`SoC`, `Voltage`, `Current`, `Anomaly Score`) |
| `00002A20-0000-1000-8000-00805F9B34FB` | Write | 64-byte Hex | ECDH Public Key Handshake & HMAC Authentication Request |

---

## 4. Firmware C++ Function API Reference (`bms_master.ino`)

### Security & IDS (`Core 0`)

```cpp
float compute_interval_ms(uint32_t now_ms);
```
Calculates inter-arrival time ($\Delta t$) between consecutive CAN frames.

```cpp
float compute_frequency();
```
Computes rolling message transmission frequency ($f$) over 32-frame buffer.

```cpp
float compute_payload_variance(const uint8_t *data, uint8_t len);
```
Calculates variance of payload data bytes across CAN frame.

```cpp
float compute_entropy(const uint8_t *data, uint8_t len);
```
Computes Shannon entropy ($H = -\sum p \log_2 p$) over 8 data bytes.

```cpp
bool is_uds_session_attack(const twai_message_t *msg);
```
Inspects CAN IDs `0x7E0`/`0x7E8` for unauthorized SecurityAccess (`0x27`) or TesterPresent (`0x3E`). Returns `true` if attack detected.

```cpp
float random_forest_predict(float dt, float freq, float var, float entropy);
```
Executes TinyML decision tree ensemble. Returns anomaly score $S \in [0.0, 1.0]$.

---

### EKF State Estimation & Control (`Core 1`)

```cpp
void ekf_set_R_scale(float score);
```
Sets measurement noise scaling $R_{\text{eff}} = R_{\text{base}} \cdot e^{10 \cdot S}$.

```cpp
void ekf_update(float I_meas, float V_meas);
```
Executes EKF prediction and measurement update steps.

```cpp
void selfheal_trigger(float score);
```
Executes closed-loop mitigation actions. When `score >= 0.90f`, drives `SSR_CUTOFF_PIN` (GPIO 17) HIGH to physically open high-side pack SSR contactor.

---

## 5. Attacker Node Control API (`attacker_node.ino`)

Set `ATTACK_MODE` variable in `attacker_node.ino` to execute target injection mode:

| Mode ID | Attack Name | Injection Behavior |
| :--- | :--- | :--- |
| `0` | Normal Baseline | Broadcasts healthy BMS telemetry (`0x180`) at $100\text{ ms}$ interval. |
| `1` | DoS Flood | Transmits high-rate zero-ID frames (`0x000`) at $500\,\mu\text{s}$ interval ($2,000\text{ msg/s}$). |
| `2` | Voltage Spoof | Injects fake depleted cell voltage frames (`0x180` with $10.0\text{V}$). |
| `3` | Replay Attack | Re-sends captured legitimate status frame 20 times rapidly. |
| `4` | CAN Fuzzing | Transmits random CAN IDs (`0x001` - `0x7FF`) with random 8-byte payloads. |
| `5` | Mixed Cycling | Cycles through attack modes 0 to 6 every 5 seconds. |
| `6` | UDS Session Hijack | Injects unauthorized SecurityAccess `0x27` & TesterPresent `0x3E` requests on `0x7E0`. |
| `7` | High-Severity Emergency | Forces score $S > 0.90$ payload to test GPIO 17 High-Side SSR cutoff trip actuation. |

---

## 6. Python ML & Simulation Pipeline API

### `generate_dataset.py`
```bash
python generate_dataset.py [PORT]
```
Captures serial output from `attacker_node` or `bms_master` into `can_dataset.csv` and `bt_events.csv`.

### `train_ids.py`
```bash
python train_ids.py
```
Loads `can_dataset.csv`, trains a 20-tree Random Forest classifier, evaluates metrics, and exports C++ header `ids_model.h` via `m2cgen`.

### `run_all_simulations.py`
```bash
python run_all_simulations.py
```
Executes the master 5-part simulation suite and generates high-resolution waveform plots in `simulations/matlab/`.
