# Cyber-Hardened BMS — Flowcharts & Diagrams

This document presents all system flowcharts, FreeRTOS task sequence diagrams, state machines, and mathematical control graphs in interactive GitHub-flavored Mermaid notation.

---

## 1. Complete 7-Layer Security Flowchart

```mermaid
flowchart TD
    subgraph L1 ["Layer 1: Wireless Security"]
        A["Incoming BLE Telemetry Connection"] --> B{"Authenticated MAC & Session Active?"}
        B -- No --> C["ECDH Key Exchange & AES-128 Challenge"]
        C --> D{"HMAC Signature Valid?"}
        D -- No --> E["Terminate Connection & Blacklist MAC"]
        D -- Yes --> F["Grant BLE Telemetry Session (5-min Timeout)"]
        B -- Yes --> F
    end

    subgraph L2 ["Layer 2: Bus Command Auth"]
        F --> G["CAN / BLE Command Received"]
        G --> H{"Valid HMAC-SHA256 & Nonce Fresh?"}
        H -- No --> I["Drop Command & Log Replay Vector"]
        H -- Yes --> J["Forward Command to Core 0 Engine"]
    end

    subgraph L3 ["Layer 3: TinyML IDS & UDS Inspector"]
        J --> K["CAN Receiver Ring Buffer"]
        K --> L{"CAN ID == 0x7E0 / 0x7E8 (UDS)?"}
        L -- Yes --> M{"Service ID == 0x27 or 0x3E?"}
        M -- Yes --> N["Flag UDS Hijack: Set Anomaly Score S = 0.96"]
        M -- No --> O["Extract Features: dt, freq, var, entropy"]
        L -- No --> O
        O --> P["Random Forest Classifier (m2cgen C++)"]
        P --> Q["Compute Anomaly Score S in [0.0, 1.0]"]
    end

    subgraph L4 ["Layer 4: Adaptive EKF Trust Scaling"]
        N & Q --> R["Scale Measurement Covariance: R_eff = R_base * e^(10*S)"]
        R --> S["Calculate Kalman Gain: K = P * H^T / (H * P * H^T + R_eff)"]
        S --> T{"S -> 1.0 (Under Attack)?"}
        T -- Yes --> U["Kalman Gain K -> 0: Suppress Corrupted Sensor Gain"]
        T -- No --> V["Standard Kalman Filter Sensor Weighting"]
    end

    subgraph L5 ["Layer 5 & 6: Self-Healing & Actuation"]
        U & V --> W["Self-Healing Decision Engine"]
        W --> X{"Anomaly Score S Range"}
        X -- "S < 0.30" --> Y1["Action 0: Normal Operation"]
        X -- "0.30 <= S < 0.50" --> Y2["Action 1: Reduce Charge Current (PWM)"]
        X -- "0.50 <= S < 0.70" --> Y3["Action 2: Disable Passive Cell Balancing"]
        X -- "0.70 <= S < 0.90" --> Y4["Action 3: Limp Home Mode (Throttle 50%)"]
        X -- "S >= 0.90" --> Y5["Action 4: High-Side SSR Pack Isolation (GPIO 17 HIGH)"]
    end
```

---

## 2. Advanced Features Flowcharts

### Digital Battery Health Passport Flowchart
```mermaid
flowchart LR
    A["BMS State Data (SoH, Cycle Count, Thermal Spikes)"] --> B["Serialize Log Format: SOH|CYC|RINT|TAMP|RUL"]
    B --> C["Compute Cryptographic SHA-256 Digest"]
    C --> D["Store Digest in NVS / MicroSD & Publish CAN 0x190"]
    D --> E["Resale Inspector / App Scans Passport Digest"]
    E --> F{"Digest Matches Cloud Record?"}
    F -- Yes --> G["Verified Untampered Battery — High Resale Trust"]
    F -- No --> H["Tamper Detected — Flag Abused Battery"]
```

### Battery Theft & Physical Cell-Bypass Tamper Detection
```mermaid
flowchart TD
    A["Monitor Pack Voltage & Current"] --> B{"V_pack < 4.0V while I_pack != 0?"}
    B -- Yes --> C["Flag Abrupt Pack Removal (Theft Vector 1)"]
    B -- No --> D{"Instantaneous Delta V > 3.5V while I_pack == 0?"}
    D -- Yes --> E["Flag Cell-Bypass Jumpering (Theft Vector 2)"]
    D -- No --> F["Normal Operating State"]
    C & E --> G["Log Non-Resettable Theft Flag to eFuse & Trip GPIO 17 SSR Cutoff"]
```

### Grid-Aware Demand-Response Charging
```mermaid
flowchart TD
    A["Grid Stress Signal Received (CAN 0x198 / Peak Hours)"] --> B{"Grid Stress Score > 0.50?"}
    B -- Yes --> C["Activate Demand-Response Mode"]
    C --> D["Cap Maximum Charge Current: I_max = 0.5 * I_nominal"]
    B -- No --> E["Maintain Standard Fast Charging"]
```

### Federated Learning Fleet Intelligence Update Cycle
```mermaid
flowchart LR
    A["Local BMS Node (On-Device IDS)"] --> B["Compute Local Gradient Delta (Delta W)"]
    B --> C["Broadcast Delta W via CAN 0x188 / BLE"]
    C --> D["Cloud TCU Gateway Aggregates Fleet Deltas (FedAvg)"]
    D --> E["Update Global Model Weights (W_global)"]
    E --> F["OTA Push Updated Model Weights to Fleet"]
```
