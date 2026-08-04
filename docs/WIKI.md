# Self-Healing Cyber-Hardened Battery Management System (BMS) — Complete Technical Wiki

Welcome to the official technical wiki for the **Self-Healing Cyber-Hardened Battery Management System**. This comprehensive guide details the mathematical foundations, zero-cost novelties, firmware implementations, and security protocols of the system.

---

## Table of Contents
1. [System Vision & 7-Layer Architecture](#1-system-vision--7-layer-architecture)
2. [2RC Equivalent Circuit & Extended Kalman Filter (EKF) Math](#2-2rc-equivalent-circuit--extended-kalman-filter-ekf-math)
3. [Dynamic Covariance Scaling Law](#3-dynamic-covariance-scaling-law)
4. [TinyML Random Forest IDS & UDS (ISO 14229) Security](#4-tinyml-random-forest-ids--uds-iso-14229-security)
5. [Digital Battery Health Passport (SHA-256 Resale Trust)](#5-digital-battery-health-passport-sha-256-resale-trust)
6. [Battery Theft & Physical Cell-Bypass Tamper Detection](#6-battery-theft--physical-cell-bypass-tamper-detection)
7. [Grid-Aware Adaptive Demand-Response Charging](#7-grid-aware-adaptive-demand-response-charging)
8. [Federated Learning Fleet Intelligence](#8-federated-learning-fleet-intelligence)
9. [Dual-Core FreeRTOS Task Isolation](#9-dual-core-freertos-task-isolation)
10. [Hardware Secure Boot V2 & eFuse Encryption Setup](#10-hardware-secure-boot-v2--efuse-encryption-setup)

---

## 1. System Vision & 7-Layer Architecture

Indian 2-wheeler and 3-wheeler Electric Vehicles (EVs) suffer from unauthenticated CAN (Controller Area Network) and Bluetooth (BLE) buses. Malicious nodes can inject false state-of-charge (SoC) frames, launch Denial of Service (DoS) floods, or execute unauthorized diagnostic UDS sessions, causing thermal runaway or sudden vehicle shutdown.

The 7-Layer Cyber-Hardened BMS resolves this by combining real-time anomaly detection with self-healing firmware controls.

### 7-Layer Security Map

| Security Layer | Name | Function / Algorithm | Output Action |
| :--- | :--- | :--- | :--- |
| **Layer 1** | BLE Peripheral Auth | ECDH SECP256R1 + AES-128-GCM + 64-nonce cache | Disconnects unauthenticated BLE pairs |
| **Layer 2** | HMAC Telemetry Auth | HMAC-SHA256 frame signature over CAN payloads | Rejects unauthorized write requests |
| **Layer 3** | TinyML IDS & UDS Inspector | Decision Tree ensemble (`m2cgen` C++) + `0x7E0` OBD-II filter | Computes Anomaly Score $S \in [0, 1]$ |
| **Layer 4** | Adaptive EKF Covariance | $R_{\text{eff}} = R_{\text{base}} \cdot e^{\lambda \cdot S}$ ($\lambda = 10.0$) | Suppresses gain ($K \to 0$) during attack |
| **Layer 5** | Self-Healing Decision Engine | 5-stage threshold state machine | Charges reduced, balancing off, Limp Mode |
| **Layer 6** | High-Side SSR Isolation | Automated P-FET gate driver on GPIO 17 ($S > 0.90$) | Disconnects physical traction pack in $<1.2\text{ ms}$ |
| **Layer 7** | Cloud TCU Gateway | Secure MQTT / TLS 1.3 telemetry & Node-RED bridge | Logs digital twin state & remote alerts |

---

## 2. 2RC Equivalent Circuit & Extended Kalman Filter (EKF) Math

### Battery Cell Model (Thevenin 2-RC Model)

The battery is modeled as a 2-RC network consisting of open-circuit voltage $V_{\text{oc}}(z)$, internal ohmic resistance $R_0$, and two polarization RC pairs $(R_1, C_1)$ and $(R_2, C_2)$:

$$V_k = V_{\text{oc}}(z_k) - I_k R_0 - V_{1, k} - V_{2, k}$$

### Continuous-Time State-Space Equations

State vector:
$$x_k = \begin{bmatrix} z_k \\ V_{1, k} \\ V_{2, k} \end{bmatrix}$$

where $z_k$ is the State of Charge (SoC), $V_{1, k}$ is the short-term RC polarization voltage, and $V_{2, k}$ is the long-term RC polarization voltage.

State Transition Equations:
$$z_{k+1} = z_k - \frac{\eta \cdot \Delta t}{Q_n} I_k$$

$$V_{1, k+1} = e^{-\frac{\Delta t}{R_1 C_1}} V_{1, k} + R_1 \left(1 - e^{-\frac{\Delta t}{R_1 C_1}}\right) I_k$$

$$V_{2, k+1} = e^{-\frac{\Delta t}{R_2 C_2}} V_{2, k} + R_2 \left(1 - e^{-\frac{\Delta t}{R_2 C_2}}\right) I_k$$

---

## 3. Dynamic Covariance Scaling Law

Under normal operation, the measurement noise covariance matrix is set to a static baseline $R = R_{\text{base}} = 0.01$. However, during a CAN spoofing or replay attack, false sensor measurements cause standard EKFs to update state estimates incorrectly.

### Exponential Covariance Scaling Formula

We dynamically scale the measurement covariance matrix $R_{\text{eff}}$ using the Layer 3 Anomaly Score $S \in [0, 1]$:

$$R_{\text{eff}} = R_{\text{base}} \cdot e^{\lambda \cdot S}$$

where $\lambda = 10.0$ is the sensitivity tuning constant.

### Kalman Gain Behavior

$$K_k = P_k^- H_k^T \left( H_k P_k^- H_k^T + R_{\text{eff}} \right)^{-1}$$

* **Normal State ($S \approx 0.0$):** $R_{\text{eff}} \approx 0.01$. The filter updates normally.
* **Under Cyber Attack ($S \approx 1.0$):** $R_{\text{eff}} = 0.01 \cdot e^{10.0} \approx 220.26$.
* As $R_{\text{eff}} \to \infty$, the Kalman Gain approaches zero ($K_k \to 0$).
* State update equation:
  $$\hat{x}_k = \hat{x}_k^- + K_k \left( y_k - h(\hat{x}_k^-) \right) \approx \hat{x}_k^-$$
* **Result:** The system ignores corrupted sensor values entirely while continuing smooth Coulomb counting, preventing controller collapse.

---

## 4. TinyML Random Forest IDS & UDS (ISO 14229) Security

The Layer 3 Intrusion Detection System classifies CAN frames into 6 classes in $<0.35\text{ ms}$:
1. **Class 0:** Normal Telemetry
2. **Class 1:** Denial of Service (DoS) Flood
3. **Class 2:** False Data Injection (Spoofing)
4. **Class 3:** Replay Attack
5. **Class 4:** CAN Fuzzing
6. **Class 5:** UDS Session Hijack (`0x27 SecurityAccess` & `0x3E TesterPresent`)

The model is exported via `m2cgen` into C++ decision trees embedded directly into `bms_master/ids_model.h`. Zero Python runtime or GPU required!

---

## 5. Digital Battery Health Passport (SHA-256 Resale Trust)

### The Resale Trust Problem in EV Markets
Used EV buyers in India currently have no reliable way to verify a battery pack's real state of health (SoH), history of cell balancing, thermal over-temperatures, or past cyber-attack interventions. This depresses resale values and slows adoption.

### Cryptographic SHA-256 Battery Certificate
The BMS generates a tamper-evident, cryptographically signed digital passport hash over non-volatile flash memory logs:

$$\text{Passport Digest} = \text{SHA-256}\Big(\text{SoH} \,||\, \text{Cycle Count} \,||\, R_{\text{int}} \,||\, \text{Thermal Spikes} \,||\, \text{Intrusion Log}\Big)$$

This 256-bit hexadecimal digest is published over CAN ID `0x190` and stored in NVS flash. Used-EV inspectors can verify the digest against cloud records using a smartphone app without trusting seller claims.

---

## 6. Battery Theft & Physical Cell-Bypass Tamper Detection

Battery packs constitute 40%–50% of an EV's cost, making them primary targets for pack theft or cell-harvesting bypass attacks.

### Detection Mechanism
The Layer 3 IDS is extended with dual physical anomaly heads:
1. **Abrupt Pack Removal:** Detects sudden $V_{\text{pack}} < 4.0\text{V}$ drop while current was non-zero, flagging physical connector disconnection.
2. **Cell-Bypass Jump:** Detects instantaneous voltage step changes ($\Delta V > 3.5\text{V}$) while $I_{\text{pack}} \approx 0\text{A}$, indicating a stolen cell being bypassed with a jumper wire.

Upon detection, the BMS logs a non-resettable tamper flag to eFuse/NVS and locks the high-side SSR isolation relay.

---

## 7. Grid-Aware Adaptive Demand-Response Charging

### Grid Stability in India's EV Transition
Uncoordinated peak-hour EV charging strains local transformers and grid distribution infrastructure in tier-2 and tier-3 Indian cities.

### Demand-Response Firmware Controller
When receiving a grid stress signal (CAN ID `0x198` or peak hour schedule):
* The BMS dynamically scales the maximum allowable charging current ceiling:
  $$I_{\text{charge\_max}} = 0.50 \cdot I_{\text{nominal}}$$
* If grid stress persists ($S_{\text{grid}} > 0.8$), charging is postponed until off-peak hours (e.g., 11:00 PM – 5:00 AM).
* Reduces peak grid load by up to 50% without requiring additional hardware.

---

## 8. Federated Learning Fleet Intelligence

### Privacy-Preserving Fleet Intelligence
A static ML model trained on offline synthetic datasets cannot adapt to novel attack vectors once deployed across thousands of vehicles.

### Local Gradient Update & Delta Aggregation
1. Each deployed BMS node evaluates local CAN traffic and fine-tunes decision tree thresholds on-device.
2. Rather than uploading raw vehicle telemetry (which violates user privacy), each BMS broadcasts 8-byte model parameter gradient deltas ($\Delta W$) over CAN ID `0x188` / BLE.
3. The TCU Cloud Gateway aggregates fleet deltas using Federated Averaging (FedAvg):
   $$W_{\text{global}}^{(t+1)} = W_{\text{global}}^{(t)} + \frac{1}{N} \sum_{i=1}^{N} \Delta W_i$$
4. Updated global weights are pushed back to the fleet via OTA firmware updates.

---

## 9. Dual-Core FreeRTOS Task Isolation

The ESP32 dual-core architecture separates critical safety control from security processing:

* **Core 0 (Security & AI Engine):**
  - Executes Layer 3 TinyML IDS classification (`<0.35 ms`).
  - Handles BLE auth, nonces, and Layer 3 UDS inspection.
  - Broadcasts Federated Learning deltas.
* **Core 1 (Deterministic Control & EKF):**
  - Reads BQ76920 AFE cell voltages & current every $100\text{ ms}$.
  - Executes Layer 4 Adaptive EKF state estimation.
  - Manages passive cell balancing and GPIO 17 High-Side SSR cutoff driver.

Inter-core communication relies on lock-free `volatile` atomic registers, preventing thread deadlocks.

---

## 10. Hardware Secure Boot V2 & eFuse Encryption Setup

To prevent physical attacks (e.g. UART dumping, JTAG probing, or custom firmware flashing):
1. **ESP32 Secure Boot V2:** Enforces RSA-3072 signature verification of the bootloader and application binary on eFuse key block 0.
2. **AES-256 Flash Encryption:** Encrypts external SPI flash contents with an internal eFuse key.
3. **JTAG & UART Disable:** Permanently disables JTAG debugging (`DISABLE_JTAG = 1`) and bootloader UART output via eFuse fuses.

---

*Documentation maintained by GCET EEE Department | 2026*
