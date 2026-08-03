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
        X -- "0.50 <= S < 0.70" --> Y3["Action 2: Disable Cell Balancing MOSFETs"]
        X -- "0.70 <= S < 0.90" --> Y4["Action 3: Limp Home Mode (CAN 0x180 Speed Limit)"]
        X -- "S >= 0.90" --> Y5["Action 4: High-Side SSR Cutoff (GPIO 17 HIGH)"]
    end

    subgraph L7 ["Layer 7: Cloud Telematics"]
        Y1 & Y2 & Y3 & Y4 & Y5 --> Z["MQTT Gateway Forwarding to Node-RED / Grafana"]
    end
```

---

## 2. FreeRTOS Dual-Core Execution Diagram

```mermaid
sequenceDiagram
    autonumber
    participant TWAI as CAN Bus / Transceiver
    participant C0 as Core 0: SecurityTask (Priority 5, 10ms)
    participant REG as Inter-Core Shared Memory
    participant C1 as Core 1: ControlTask (Priority 4, 100ms)
    participant AFE as TI BQ76920 AFE / I2C
    participant SSR as High-Side SSR (GPIO 17)

    loop Every 10ms (Core 0 Execution)
        TWAI->>C0: Receive CAN Frame (twai_receive)
        C0->>C0: Extract Features (dt, freq, var, entropy)
        C0->>C0: Check UDS SIDs (0x27 / 0x3E)
        C0->>C0: Run Random Forest (m2cgen C++ tree)
        C0->>C0: Calculate score S & R_eff = R_base * exp(10*S)
        C0->>REG: Atomic Write anomaly_score & ekf_R_scale
    end

    loop Every 100ms (Core 1 Execution)
        C1->>AFE: Read Cell Voltages & Current via I2C
        AFE-->>C1: Raw Voltage & Current ADC Data
        REG->>C1: Read ekf_R_scale & anomaly_score
        C1->>C1: Run Adaptive EKF Update (Kalman Gain K scaled)
        C1->>C1: Update Digital Twin (SoH, Rint, Thermal Pred)
        C1->>C1: Evaluate Self-Healing Action
        alt Anomaly Score S >= 0.90
            C1->>SSR: Set GPIO 17 HIGH (High-Side SSR Pack Cutoff)
        else Anomaly Score S < 0.90
            C1->>SSR: Maintain GPIO 17 LOW (Connected)
        end
        C1->>TWAI: Publish Status Frame 0x180
    end
```

---

## 3. EKF Covariance Scaling & Gain Suppression Flow

```mermaid
graph LR
    A["Raw Cell Voltage Reading (V_meas)"] --> B["Compute Innovation: y = V_meas - V_pred"]
    C["Layer 3 Anomaly Score S"] --> D["R_eff = R_base * e^(10*S)"]
    D --> E["Kalman Gain: K = P * H^T / (H * P * H^T + R_eff)"]
    B --> F["State Update: x_hat = x_pred + K * y"]
    E --> F
    
    subgraph Behavior ["Dynamic Filter Behavior"]
        G["S = 0.0 (Clean Traffic)"] --> H["R_eff = 0.01 V^2"] --> I["K = Optimal Gain (~0.45)"] --> J["Trust Sensors"]
        K["S = 1.0 (Cyber Attack)"] --> L["R_eff = 220.26 V^2"] --> M["K -> 0.00002"] --> N["Ignore Sensors / Trust Internal Model"]
    end
```

---

## 4. Self-Healing State Machine Transitions

```mermaid
stateDiagram-v2
    [*] --> HEAL_NORMAL : Power On / System Healthy (S < 0.30)
    
    HEAL_NORMAL --> HEAL_REDUCE_CHARGE : Anomaly Score 0.30 <= S < 0.50
    HEAL_REDUCE_CHARGE --> HEAL_NORMAL : Anomaly Score S < 0.30
    
    HEAL_REDUCE_CHARGE --> HEAL_DISABLE_BALANCING : Anomaly Score 0.50 <= S < 0.70
    HEAL_DISABLE_BALANCING --> HEAL_REDUCE_CHARGE : Anomaly Score 0.30 <= S < 0.50
    
    HEAL_DISABLE_BALANCING --> HEAL_LIMP_HOME : Anomaly Score 0.70 <= S < 0.90
    HEAL_LIMP_HOME --> HEAL_DISABLE_BALANCING : Anomaly Score 0.50 <= S < 0.70
    
    HEAL_LIMP_HOME --> HEAL_ISOLATE_CELL : Anomaly Score S >= 0.90 / UDS Hijack
    HEAL_ISOLATE_CELL --> [*] : High-Side SSR Cutoff Tripped (GPIO 17 HIGH) / Manual Lockout Reset
```

---

## 5. UDS Session Hijack Inspection Flow

```mermaid
flowchart TD
    A["CAN Message Arrives"] --> B{"CAN ID == 0x7E0 or 0x7E8?"}
    B -- No --> C["Proceed to Standard Random Forest Feature Extraction"]
    B -- Yes --> D{"DLC >= 2?"}
    D -- No --> C
    D -- Yes --> E{"Data[1] (SID) == 0x27 or 0x3E?"}
    E -- No --> C
    E -- Yes --> F["Unauthorized UDS Session Request Detected!"]
    F --> G["Override Anomaly Score S = 0.96"]
    G --> H["Scale R_eff = R_base * exp(9.6) = 14,764 * R_base"]
    H --> I["Drive Kalman Gain K -> 0"]
    I --> J["Trip High-Side SSR Pack Isolation (GPIO 17 HIGH)"]
```
