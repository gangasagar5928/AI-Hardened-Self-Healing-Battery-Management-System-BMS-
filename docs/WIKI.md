# Cyber-Hardened BMS — Project Wiki & Knowledge Base

Welcome to the **Cyber-Hardened Battery Management System (BMS)** Wiki. This document provides complete technical specifications, mathematical foundations, architectural explanations, and security guidelines.

---

## Table of Contents
1. [Core Theory & Mathematical Foundation](#1-core-theory--mathematical-foundation)
2. [Dual-Core FreeRTOS Task Isolation](#2-dual-core-freertos-task-isolation)
3. [4S Bench Sub-Module vs 16S Pack Scalability](#3-4s-bench-sub-module-vs-16s-pack-scalability)
4. [7-Layer Security Architecture](#4-7-layer-security-architecture)
5. [Self-Healing Decision Engine State Machine](#5-self-healing-decision-engine-state-machine)
6. [ESP32 Hardware Secure Boot & Flash Encryption](#6-esp32-hardware-secure-boot--flash-encryption)

---

## 1. Core Theory & Mathematical Foundation

State-of-Charge (SoC) estimation relies on a discrete 2RC Thevenin equivalent circuit battery model:

### State Transition Vector
$$x_k = \begin{bmatrix} \text{SoC}_k \\ V_{C1, k} \\ V_{C2, k} \end{bmatrix}$$

1. **State Propagation:**
   $$\text{SoC}_{k+1} = \text{SoC}_k - \frac{\eta \cdot I_k \cdot \Delta t}{Q_{\text{nom}}}$$
   $$V_{C1, k+1} = V_{C1, k} \cdot e^{-\frac{\Delta t}{R_1 C_1}} + I_k R_1 \left(1 - e^{-\frac{\Delta t}{R_1 C_1}}\right)$$
   $$V_{C2, k+1} = V_{C2, k} \cdot e^{-\frac{\Delta t}{R_2 C_2}} + I_k R_2 \left(1 - e^{-\frac{\Delta t}{R_2 C_2}}\right)$$

2. **Measurement Equation:**
   $$V_{\text{pred}, k} = \text{OCV}(\text{SoC}_k) - I_k R_0 - V_{C1, k} - V_{C2, k}$$

3. **Dynamic Covariance Scaling Law (Patent Core):**
   $$R_{\text{eff}} = R_{\text{base}} \cdot e^{\lambda \cdot S}$$
   * $R_{\text{base}} = 0.01\text{ V}^2$ (nominal sensor noise variance)
   * $\lambda = 10.0$ (scaling factor)
   * $S \in [0, 1]$ (real-time ML anomaly score)

4. **Kalman Gain & Suppression:**
   $$K_k = \frac{P_k^- H^T}{H P_k^- H^T + R_{\text{eff}}}$$
   * When $S \to 0$ (clean traffic): $R_{\text{eff}} = R_{\text{base}}$, $K_k$ is optimal, sensor updates filter measurement noise.
   * When $S \to 1$ (cyber-attack): $R_{\text{eff}} \to R_{\text{base}} \cdot e^{10} \approx 22,026 \cdot R_{\text{base}}$, driving $K_k \to 0$. The filter ignores corrupted voltage sensor readings and relies strictly on internal model prediction propagation.

---

## 2. Dual-Core FreeRTOS Task Isolation

The system leverages the ESP32 dual-core Xtensa LX6 architecture to guarantee real-time safety:

| Task Name | Pinned Core | Priority | Period | Responsibilities |
| :--- | :--- | :--- | :--- | :--- |
| `SecurityTask` | **Core 0** | 5 (High) | $10\text{ ms}$ | CAN frame ring buffer, feature extraction ($\Delta t, f, \text{Var}, H$), UDS ISO 14229 inspection, Random Forest prediction, EKF $R_{\text{eff}}$ scaling. |
| `ControlTask` | **Core 1** | 4 (Med) | $100\text{ ms}$ | BQ76920 I2C AFE polling, EKF state update, Digital Twin SoH/thermal tracking, active/passive cell balancing, OLED update, High-Side SSR cutoff drive on GPIO 17. |

**Inter-Core Communication:** Shared volatile variables (`anomaly_score`, `ekf_R_scale`) transferred via lock-free atomic registers, preventing thread lockouts or timing jitter.

---

## 3. 4S Bench Sub-Module vs 16S Pack Scalability

* **Hardware Prototype Setup:** 4S 18650 Li-ion module ($14.8\text{V}$ nominal / $16.8\text{V}$ max) managed by TI BQ76920 AFE IC.
* **Traction Pack Scaling:** Standard Indian EV 2W/3W traction packs operate at $57.6\text{V}$ nominal (16S).
* **Modular Sub-Unit Daisy-Chaining:** Full 16S or 100S packs are constructed by daisy-chaining 4x 4S sub-modules over isolated SPI or CAN communication buses. Each sub-module runs an identical local EKF and Layer 3 security monitor, providing distributed resilience across the entire battery pack.

---

## 4. 7-Layer Security Architecture

1. **Layer 1: Secure BLE Access:** ECDH key exchange + AES-128-GCM encrypted BLE telemetry channel with 5-minute session timeout.
2. **Layer 2: HMAC Command Authentication:** HMAC-SHA256 signature verification with 64-nonce replay cache preventing command replay.
3. **Layer 3: TinyML Intrusion Detection & UDS Inspector:** On-device Random Forest classifier ($<0.35\text{ ms}$) evaluating CAN frame inter-arrival time, message frequency, payload variance, and Shannon entropy. Inspects OBD-II/UDS CAN IDs `0x7E0`/`0x7E8` for unauthorized `0x27` SecurityAccess and `0x3E` TesterPresent requests.
4. **Layer 4: Trust-Scaled Adaptive EKF:** Dynamic measurement noise covariance scaling ($R_{\text{eff}} = R_{\text{base}} \cdot e^{10 \cdot S}$) suppressing corrupted sensor gain ($K \to 0$).
5. **Layer 5: Self-Healing Decision Engine:** Automated closed-loop mitigation state machine executing active charge reduction, balancing disablement, limp-home mode, or automated high-side SSR pack isolation.
6. **Layer 6: Cell Reconfiguration & Passive Balancing:** Low-side IRLML2502 MOSFETs + $47\,\Omega$ bleed resistors ($89\text{ mA}$ bleed) for cell equalization, backed by GPIO 17 High-Side SSR pack isolation.
7. **Layer 7: Secure TCU Cloud Gateway:** Encrypted MQTT forwarding over cellular/Wi-Fi to Node-RED and Grafana real-time forensic monitoring dashboards.

---

## 5. Self-Healing Decision Engine State Machine

| Anomaly Score Range | State Action | Hardware Actuation | Safety Rationale |
| :--- | :--- | :--- | :--- |
| $S < 0.30$ | `HEAL_NORMAL` | SSR Cutoff LOW, Normal Operation | System healthy; standard balancing & charge enabled. |
| $0.30 \le S < 0.50$ | `HEAL_REDUCE_CHARGE` | PWM duty reduced on charger pin | Minor anomaly detected; reduce thermal stress. |
| $0.50 \le S < 0.70$ | `HEAL_DISABLE_BALANCING` | Balancing MOSFET gates LOW | Moderate threat; freeze cell balancing to prevent erroneous cell discharge. |
| $0.70 \le S < 0.90$ | `HEAL_LIMP_HOME` | Speed-limit frame sent on CAN `0x180` | High threat; notify vehicle ECU to cap speed & power. |
| $S \ge 0.90$ | `HEAL_ISOLATE_CELL` | **GPIO 17 driven HIGH (SSR Trip)** | **Critical emergency; physical high-side pack disconnect within $<1.2\text{ ms}$.** |

---

## 6. ESP32 Hardware Secure Boot & Flash Encryption

To prevent physical UART/JTAG flash dumping, firmware reverse engineering, or bootloader tampering on the bench setup:

1. **Flash Encryption:** AES-256 eFuse key generation encrypts internal flash memory.
2. **Secure Boot V2:** RSA-3072 signature verification ensures only cryptographically signed binaries boot.

### Enablement Commands:
```bash
# Generate Secure Boot signing key
espsecure.py generate_signing_key --version 2 secure_boot_signing_key.pem

# Burn AES-256 Flash Encryption key to ESP32 eFuse
espefuse.py --port COM3 burn_key flash_encryption my_flash_key.bin

# Build signed binary
idf.py build
```
